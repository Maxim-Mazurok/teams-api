/**
 * Interactive browser login for Teams.
 *
 * Opens a visible Chromium window (Playwright's bundled browser),
 * navigates to Teams, and waits for the user to complete login manually.
 * Works on all platforms (macOS, Windows, Linux).
 */

import type { TeamsToken, InteractiveLoginOptions } from "../types.js";
import { resolveTeamsRegion } from "../region.js";
import { diagnosePageState } from "./page-diagnostics.js";
import { launchInteractiveBrowser } from "../browser-runtime.js";
import {
  captureTokensFromPage,
  TOKEN_INTERCEPT_TIMEOUT,
  PAGE_STATE_POLL_INTERVAL,
  type LogFunction,
} from "./token-capture.js";

const TEAMS_URL = "https://teams.cloud.microsoft/";
const INTERACTIVE_LOGIN_TIMEOUT = 5 * 60 * 1_000;

/**
 * Acquire a Teams skype token via interactive browser login.
 *
 * Opens a visible Chromium window (Playwright's bundled browser),
 * navigates to Teams, and waits for the user to complete login
 * manually. Once Teams loads, the skype token is captured via
 * CDP Fetch interception.
 *
 * Works on all platforms (macOS, Windows, Linux). No FIDO2 passkey
 * or system Chrome required.
 */
export async function acquireTokenViaInteractiveLogin(
  options?: InteractiveLoginOptions,
): Promise<TeamsToken> {
  const { chromium } = await import("playwright");
  const log: LogFunction = options?.verbose
    ? console.log.bind(console)
    : () => {};

  log("Launching browser for interactive login...");
  const browser = await launchInteractiveBrowser(chromium, log);
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    log("Navigating to Teams...");
    await page.goto(TEAMS_URL, {
      waitUntil: "domcontentloaded",
      timeout: INTERACTIVE_LOGIN_TIMEOUT,
    });

    // Pre-fill email if provided
    if (options?.email) {
      try {
        await page.waitForURL(
          /login\.microsoftonline\.com|login\.microsoft\.com/,
          { timeout: 30_000 },
        );

        const emailInput = await page
          .waitForSelector(
            '#i0116, input[name="loginfmt"], input[type="email"]',
            { timeout: 10_000 },
          )
          .catch(() => null);

        if (emailInput) {
          await emailInput.fill(options.email);
          log(`Pre-filled email: ${options.email}`);
        }
      } catch {
        log("Could not pre-fill email (login page may have changed)");
      }
    }

    // Wait for user to complete login, polling page state for diagnostics
    log("Waiting for you to complete login in the browser window...");
    const loginDeadline = Date.now() + INTERACTIVE_LOGIN_TIMEOUT;
    while (Date.now() < loginDeadline) {
      const state = await diagnosePageState(page);
      if (state.phase === "teams-loaded" || state.phase === "teams-loading") {
        log(`Reached Teams: ${state.detail}`);
        break;
      }
      log(`Login progress: ${state.detail}`);
      await new Promise((resolve) =>
        setTimeout(resolve, PAGE_STATE_POLL_INTERVAL),
      );
    }

    if (
      !page.url().includes("teams.cloud.microsoft") &&
      !page.url().includes("teams.microsoft.com")
    ) {
      const state = await diagnosePageState(page);
      throw new Error(
        `Login did not complete within ${INTERACTIVE_LOGIN_TIMEOUT / 1_000}s. ` +
          `Current state: ${state.phase} — ${state.detail} (URL: ${state.url})`,
      );
    }

    log("Login detected, capturing token...");

    const {
      skypeToken,
      region,
      bearerToken,
      substrateToken,
      amsToken,
      sharePointToken,
    } = await captureTokensFromPage(page, log, TOKEN_INTERCEPT_TIMEOUT);

    return {
      skypeToken,
      region: resolveTeamsRegion(options?.region, region),
      bearerToken,
      substrateToken,
      amsToken,
      sharePointToken,
    };
  } finally {
    await context.close();
    await browser.close();
  }
}
