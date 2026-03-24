/**
 * Shared token capture logic via Chrome DevTools Protocol (CDP).
 *
 * Used by both the auto-login and interactive login flows (Playwright-based).
 * Captures skype, bearer, and substrate tokens from intercepted network requests.
 */

import { detectTeamsRegionFromUrl } from "../region.js";
import { diagnosePageState } from "./page-diagnostics.js";

export const TOKEN_INTERCEPT_TIMEOUT = 60 * 1_000;
export const PAGE_RELOAD_TIMEOUT = 30 * 1_000;
export const PAGE_STATE_POLL_INTERVAL = 3 * 1_000;

export type LogFunction = (...arguments_: unknown[]) => void;

/**
 * Wait for the page to reach Teams and capture authentication tokens via CDP.
 *
 * Polls the page state periodically to provide diagnostic feedback instead of
 * relying solely on timeouts. Throws a descriptive error if the login flow
 * stalls or reaches an unexpected state.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function captureTokensFromPage(
  page: any,
  log: LogFunction,
  interceptTimeout: number,
): Promise<{
  skypeToken: string;
  region: string | undefined;
  bearerToken: string | undefined;
  substrateToken: string | undefined;
}> {
  const cdpSession = (await page.context().newCDPSession(page)) as {
    send: (
      method: string,
      parameters?: Record<string, unknown>,
    ) => Promise<unknown>;
    on: (
      event: string,
      handler: (event: Record<string, unknown>) => void,
    ) => void;
    detach: () => Promise<void>;
  };
  let skypeToken: string | null = null;
  let region: string | null = null;
  let bearerToken: string | null = null;
  let substrateToken: string | null = null;

  await cdpSession.send("Fetch.enable", {
    patterns: [
      { urlPattern: "*teams*", requestStage: "Request" },
      { urlPattern: "*substrate.office.com*", requestStage: "Request" },
    ],
  });

  log("Listening for tokens in network requests...");
  const tokenPromise = new Promise<void>((resolve) => {
    const timeout = setTimeout(() => {
      log("Token intercept timed out");
      resolve();
    }, interceptTimeout);

    cdpSession.on(
      "Fetch.requestPaused",
      async (event: Record<string, unknown>) => {
        const request = event.request as {
          url?: string;
          headers?: Record<string, string>;
        };
        const requestId = event.requestId as string;
        const requestUrl = request.url ?? "";
        const detectedRegion = detectTeamsRegionFromUrl(requestUrl);

        if (detectedRegion && !region) {
          region = detectedRegion;
          log(`Detected Teams region from request URL: ${region}`);
        }

        for (const [name, value] of Object.entries(request.headers ?? {})) {
          if (name.toLowerCase() === "x-skypetoken" && !skypeToken) {
            skypeToken = value;
          }
          if (
            name.toLowerCase() === "authorization" &&
            value.startsWith("Bearer ")
          ) {
            if (requestUrl.includes("/api/mt/") && !bearerToken) {
              bearerToken = value.slice("Bearer ".length);
            }
            if (
              requestUrl.includes("substrate.office.com") &&
              !substrateToken
            ) {
              substrateToken = value.slice("Bearer ".length);
            }
          }
        }

        try {
          await cdpSession.send("Fetch.continueRequest", { requestId });
        } catch {
          // Request may have already been handled
        }

        if (skypeToken) {
          log("Skype token captured from request headers");
          clearTimeout(timeout);
          resolve();
        }
      },
    );
  });

  log("Reloading page to intercept token...");
  page
    .reload({ waitUntil: "domcontentloaded", timeout: PAGE_RELOAD_TIMEOUT })
    .catch(() => {
      // Reload may time out, but token should be captured by then
    });

  await tokenPromise;

  // If we have the skype token but not the substrate token yet,
  // wait a bit longer for substrate requests to fire
  if (skypeToken && !substrateToken) {
    log("Waiting for substrate token...");
    await new Promise<void>((resolve) => {
      const extraTimeout = setTimeout(resolve, 5_000);
      const checkInterval = setInterval(() => {
        if (substrateToken) {
          clearTimeout(extraTimeout);
          clearInterval(checkInterval);
          resolve();
        }
      }, 200);
      extraTimeout.unref?.();
    });
  }

  await cdpSession.send("Fetch.disable");
  await cdpSession.detach();

  if (!skypeToken) {
    const state = await diagnosePageState(page);
    throw new Error(
      `Failed to capture skype token. Page state: ${state.phase} — ${state.detail} (URL: ${state.url})`,
    );
  }

  const capturedSkypeToken: string = skypeToken;
  log(`Token captured (${capturedSkypeToken.length} chars)`);
  if (bearerToken) {
    log("Bearer token also captured for profile resolution");
  }
  if (substrateToken) {
    log("Substrate token also captured for people/chat search");
  }

  return {
    skypeToken,
    region: region ?? undefined,
    bearerToken: bearerToken ?? undefined,
    substrateToken: substrateToken ?? undefined,
  };
}
