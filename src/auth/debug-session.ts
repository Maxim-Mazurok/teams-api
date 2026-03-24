/**
 * Debug session token capture for Teams.
 *
 * Connects via puppeteer-core to a running Chrome instance with Teams open,
 * intercepts network headers during a page reload, and extracts tokens.
 */

import type { TeamsToken, ManualTokenOptions } from "../types.js";
import { detectTeamsRegionFromUrl, resolveTeamsRegion } from "../region.js";

const FETCH_INTERCEPT_TIMEOUT = 30 * 1_000;
const PAGE_RELOAD_TIMEOUT = 30 * 1_000;

/**
 * Capture a Teams skype token from an existing Chrome debug session.
 *
 * Connects via puppeteer-core to a running Chrome instance with Teams open,
 * intercepts network headers during a page reload, and extracts the token.
 *
 * Usage:
 *   1. Start Chrome with --remote-debugging-port=9222
 *   2. Navigate to Teams and log in
 *   3. Call this function
 */
export async function acquireTokenViaDebugSession(
  options?: ManualTokenOptions,
): Promise<TeamsToken> {
  const puppeteer = await import("puppeteer-core");
  const debugPort = options?.debugPort ?? 9222;
  const browserUrl = `http://127.0.0.1:${debugPort}`;

  const browser = await puppeteer.default.connect({ browserURL: browserUrl });

  try {
    const pages = await browser.pages();
    const teamsPage = pages.find((page) =>
      /teams\.(microsoft|cloud\.microsoft)/i.test(page.url()),
    );

    if (!teamsPage) {
      throw new Error(
        "No Teams page found in the browser. Navigate to Teams and log in first.",
      );
    }

    const cdpSession = await teamsPage.createCDPSession();
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

    const tokenPromise = new Promise<void>((resolve) => {
      const timeout = setTimeout(resolve, FETCH_INTERCEPT_TIMEOUT);

      cdpSession.on(
        "Fetch.requestPaused",
        async (
          event: import("puppeteer-core").Protocol.Fetch.RequestPausedEvent,
        ) => {
          const headers = event.request.headers ?? {};
          const requestUrl = event.request.url ?? "";
          const detectedRegion = detectTeamsRegionFromUrl(requestUrl);

          if (detectedRegion && !region) {
            region = detectedRegion;
          }

          for (const [name, value] of Object.entries(headers)) {
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
            await cdpSession.send("Fetch.continueRequest", {
              requestId: event.requestId,
            });
          } catch {
            // Request may have already been handled
          }

          if (skypeToken) {
            clearTimeout(timeout);
            resolve();
          }
        },
      );
    });

    teamsPage
      .reload({ waitUntil: "networkidle2", timeout: PAGE_RELOAD_TIMEOUT })
      .catch(() => {
        // Reload may time out
      });

    await tokenPromise;

    // If we have the skype token but not the substrate token yet,
    // wait a bit longer for substrate requests to fire
    if (skypeToken && !substrateToken) {
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
      throw new Error(
        "Failed to capture skype token. Ensure Teams is loaded and authenticated.",
      );
    }

    return {
      skypeToken,
      region: resolveTeamsRegion(options?.region, region ?? undefined),
      bearerToken: bearerToken ?? undefined,
      substrateToken: substrateToken ?? undefined,
    };
  } finally {
    browser.disconnect();
  }
}
