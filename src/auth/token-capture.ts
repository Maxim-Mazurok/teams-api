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
  amsToken: string | undefined;
  sharePointToken: string | undefined;
  sharePointHost: string | undefined;
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
      const checkInterval = setInterval(() => {
        if (substrateToken) {
          clearInterval(checkInterval);
          resolve();
        }
      }, 200);
      const extraTimeout = setTimeout(() => {
        clearInterval(checkInterval);
        resolve();
      }, 5_000);
      extraTimeout.unref?.();
      checkInterval.unref?.();
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

  // Read the AMS (IC3) token from the browser's MSAL cache in localStorage.
  // This token has audience https://ic3.teams.office.com and is used for image uploads
  // to the Async Media Service. It's acquired by MSAL during login but may not appear
  // in network requests during a page reload (AMS requests only happen with images).
  let amsToken: string | undefined;
  let sharePointToken: string | undefined;
  let sharePointHost: string | undefined;
  try {
    const msalTokens = await page.evaluate(() => {
      let foundAmsToken: string | null = null;
      let foundSharePointToken: string | null = null;
      let foundSharePointHost: string | null = null;
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.includes("accesstoken")) {
          try {
            const value = JSON.parse(localStorage.getItem(key) ?? "");
            if (
              value.target &&
              value.target.includes("ic3.teams.office.com") &&
              value.secret
            ) {
              foundAmsToken = value.secret as string;
            }
            if (
              value.target &&
              value.target.includes("sharepoint.com") &&
              value.secret
            ) {
              foundSharePointToken = value.secret as string;
              // Extract the SharePoint host from the MSAL target field.
              // The target looks like "https://contoso-my.sharepoint.com/.default"
              // or "contoso-my.sharepoint.com/AllSites.Write AllSites.Read".
              // We also check the environment/realm fields from the MSAL cache key.
              const targetString = value.target as string;
              const hostMatch = targetString.match(
                /([a-z0-9-]+\.sharepoint\.com)/i,
              );
              if (hostMatch) {
                foundSharePointHost = hostMatch[1];
              }
            }
          } catch {
            // Skip malformed entries
          }
        }
      }
      return {
        amsToken: foundAmsToken,
        sharePointToken: foundSharePointToken,
        sharePointHost: foundSharePointHost,
      };
    });
    amsToken = msalTokens?.amsToken ?? undefined;
    sharePointToken = msalTokens?.sharePointToken ?? undefined;
    sharePointHost = msalTokens?.sharePointHost ?? undefined;
    if (amsToken) {
      log("AMS token captured from MSAL cache for image upload");
    }
    if (sharePointToken) {
      log("SharePoint token captured from MSAL cache for file download");
      if (sharePointHost) {
        log(`SharePoint host: ${sharePointHost}`);
      }
    }
  } catch {
    // page.evaluate may fail if the page context is unavailable
  }

  return {
    skypeToken,
    region: region ?? undefined,
    bearerToken: bearerToken ?? undefined,
    substrateToken: substrateToken ?? undefined,
    amsToken: amsToken ?? undefined,
    sharePointToken: sharePointToken ?? undefined,
    sharePointHost: sharePointHost ?? undefined,
  };
}
