/**
 * Automatic FIDO2 passkey login for Teams.
 *
 * Launches system Chrome with a fresh profile, completes the Microsoft
 * Entra ID FIDO2 login flow using a platform authenticator, and captures
 * the skype token from network requests. macOS only.
 */

import type { TeamsToken, AutoLoginOptions } from "../types.js";
import { resolveTeamsRegion } from "../region.js";
import { diagnosePageState } from "./page-diagnostics.js";
import {
  captureTokensFromPage,
  TOKEN_INTERCEPT_TIMEOUT,
  PAGE_STATE_POLL_INTERVAL,
  type LogFunction,
} from "./token-capture.js";

const TEAMS_URL = "https://teams.cloud.microsoft/";
const DEFAULT_SYSTEM_CHROME_PATH =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const DEFAULT_PROFILE_DIRECTORY = "/tmp/teams-api-chrome-profile";
const AUTO_LOGIN_TIMEOUT = 2 * 60 * 1_000;
const BROWSER_LAUNCH_TIMEOUT = 30 * 1_000;

/**
 * Acquire a Teams skype token with zero user interaction.
 *
 * Launches system Chrome with a fresh profile, navigates to Teams,
 * enters the corporate email, lets the platform authenticator complete
 * the FIDO2 passkey challenge, then intercepts the skype token from
 * the authenticated network requests.
 *
 * Requirements:
 *   - macOS with a FIDO2 platform authenticator (e.g. Intune/Company Portal)
 *   - System Chrome installed (or custom chromePath)
 *   - playwright package
 */
export async function acquireTokenViaAutoLogin(
  options: AutoLoginOptions,
): Promise<TeamsToken> {
  const { chromium } = await import("playwright");

  const chromePath = options.chromePath ?? DEFAULT_SYSTEM_CHROME_PATH;
  const profileDirectory =
    options.profileDirectory ?? DEFAULT_PROFILE_DIRECTORY;
  const headless = options.headless ?? true;
  const log: LogFunction = options.verbose
    ? console.log.bind(console)
    : () => {};

  // Clean up any previous profile to ensure a fresh session
  const { execSync } = await import("node:child_process");
  execSync(`rm -rf "${profileDirectory}"`);

  log("Launching Chrome with fresh profile...");
  const launchPromise = chromium.launchPersistentContext(profileDirectory, {
    headless,
    executablePath: chromePath,
    args: [
      "--disable-blink-features=AutomationControlled",
      "--disable-features=VirtualAuthenticators",
      "--enable-features=WebAuthenticationMacPlatform",
      "--no-first-run",
      "--no-default-browser-check",
    ],
    ignoreDefaultArgs: ["--disable-component-extensions-with-background-pages"],
  });

  const context = await Promise.race([
    launchPromise,
    new Promise<never>((_, reject) =>
      setTimeout(
        () =>
          reject(
            new Error(
              `Chrome launch timed out after ${BROWSER_LAUNCH_TIMEOUT / 1_000}s`,
            ),
          ),
        BROWSER_LAUNCH_TIMEOUT,
      ),
    ),
  ] as const);
  log("Chrome launched successfully");

  try {
    const page = context.pages()[0] || (await context.newPage());

    log("Navigating to Teams...");
    await page.goto(TEAMS_URL, {
      waitUntil: "domcontentloaded",
      timeout: AUTO_LOGIN_TIMEOUT,
    });

    // Wait for Entra ID login page
    log(`Current URL: ${page.url()}`);
    log("Waiting for Entra ID login page...");
    await page
      .waitForURL(/login\.microsoftonline\.com|login\.microsoft\.com/, {
        timeout: AUTO_LOGIN_TIMEOUT,
      })
      .catch(() => {
        log("Did not reach login page (may already be at Teams)");
      });

    // If at the login page, enter email and submit
    if (
      page.url().includes("login.microsoftonline.com") ||
      page.url().includes("login.microsoft.com")
    ) {
      log("At login page, entering email...");

      const emailInput = await page
        .waitForSelector(
          '#i0116, input[name="loginfmt"], input[type="email"]',
          { timeout: 15_000 },
        )
        .catch(() => null);

      if (emailInput) {
        await emailInput.fill(options.email);

        const nextButton = await page
          .waitForSelector("#idSIButton9", { timeout: 5_000 })
          .catch(() => null);
        if (nextButton) {
          log("Submitting email...");
          await nextButton.click();
        }
      }

      // Wait for login flow to complete, polling page state for diagnostics
      log("Waiting for authentication to complete...");
      const loginDeadline = Date.now() + AUTO_LOGIN_TIMEOUT;
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

      // Final check after poll loop
      if (
        !page.url().includes("teams.cloud.microsoft") &&
        !page.url().includes("teams.microsoft.com")
      ) {
        const state = await diagnosePageState(page);
        throw new Error(
          `Login did not complete within ${AUTO_LOGIN_TIMEOUT / 1_000}s. ` +
            `Current state: ${state.phase} — ${state.detail} (URL: ${state.url})`,
        );
      }
    }

    log(`Login flow finished. URL: ${page.url()}`);

    // Capture tokens via shared helper
    log("Capturing tokens...");
    const { skypeToken, region, bearerToken, substrateToken, amsToken, sharePointToken, sharePointHost } =
      await captureTokensFromPage(page, log, TOKEN_INTERCEPT_TIMEOUT);

    return {
      skypeToken,
      region: resolveTeamsRegion(options.region, region),
      bearerToken,
      substrateToken,
      amsToken,
      sharePointToken,
      sharePointHost,
    };
  } finally {
    await context.close();
    try {
      execSync(`rm -rf "${profileDirectory}"`);
    } catch {
      // Best-effort cleanup
    }
  }
}
