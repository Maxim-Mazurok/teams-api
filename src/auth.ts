/**
 * Authentication module for Teams API.
 *
 * Two strategies:
 *   1. Auto-login: Launch system Chrome, complete FIDO2/passkey via platform authenticator,
 *      capture skype token from CDP Fetch interception.
 *   2. Manual: Connect to a running Chrome debug session with Teams open,
 *      intercept the token during a page reload.
 *
 * Both return a TeamsToken that can be used with the API layer.
 */

import type {
  TeamsToken,
  AutoLoginOptions,
  ManualTokenOptions,
} from "./types.js";

const TEAMS_URL = "https://teams.cloud.microsoft/";
const DEFAULT_SYSTEM_CHROME_PATH =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const DEFAULT_PROFILE_DIRECTORY = "/tmp/teams-api-chrome-profile";
const LOGIN_TIMEOUT = 60_000;
const TOKEN_INTERCEPT_TIMEOUT = 30_000;
const FETCH_INTERCEPT_TIMEOUT = 20_000;
const PAGE_RELOAD_TIMEOUT = 25_000;
const BROWSER_LAUNCH_TIMEOUT = 30_000;

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
  const log = options.verbose ? console.log.bind(console) : () => {};

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
  ]);
  log("Chrome launched successfully");

  try {
    const page = context.pages()[0] || (await context.newPage());

    log("Navigating to Teams...");
    await page.goto(TEAMS_URL, {
      waitUntil: "domcontentloaded",
      timeout: LOGIN_TIMEOUT,
    });

    // Wait for Entra ID login page
    log(`Current URL: ${page.url()}`);
    log("Waiting for Entra ID login page...");
    await page
      .waitForURL(/login\.microsoftonline\.com|login\.microsoft\.com/, {
        timeout: LOGIN_TIMEOUT,
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

      // Wait for FIDO/passkey flow to complete → redirect back to Teams
      log("Waiting for passkey authentication...");
      await page.waitForURL(/teams\.cloud\.microsoft/, {
        timeout: LOGIN_TIMEOUT,
        waitUntil: "domcontentloaded",
      });
    }

    log(`Login flow finished. URL: ${page.url()}`);

    // Capture the skype token via CDP Fetch interception
    log("Capturing skype token...");
    const cdpSession = await page.context().newCDPSession(page);
    let skypeToken: string | null = null;

    await cdpSession.send("Fetch.enable", {
      patterns: [{ urlPattern: "*teams*", requestStage: "Request" }],
    });

    log("Listening for skype token in network requests...");
    const tokenPromise = new Promise<void>((resolve) => {
      const timeout = setTimeout(() => {
        log("Token intercept timed out");
        resolve();
      }, TOKEN_INTERCEPT_TIMEOUT);

      cdpSession.on(
        "Fetch.requestPaused",
        async (event: Record<string, unknown>) => {
          const request = event.request as {
            headers?: Record<string, string>;
          };
          const requestId = event.requestId as string;

          for (const [name, value] of Object.entries(request.headers ?? {})) {
            if (name.toLowerCase() === "x-skypetoken" && !skypeToken) {
              skypeToken = value;
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
      .reload({ waitUntil: "domcontentloaded", timeout: LOGIN_TIMEOUT })
      .catch(() => {
        // Reload may time out, but token should be captured by then
      });

    await tokenPromise;

    await cdpSession.send("Fetch.disable");
    await cdpSession.detach();

    if (!skypeToken) {
      throw new Error(
        "Failed to capture skype token after login. Teams may not have fully loaded.",
      );
    }

    const capturedToken: string = skypeToken;
    log(`Token captured (${capturedToken.length} chars)`);

    return { skypeToken: capturedToken, region: "apac" };
  } finally {
    await context.close();
    try {
      execSync(`rm -rf "${profileDirectory}"`);
    } catch {
      // Best-effort cleanup
    }
  }
}

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

    await cdpSession.send("Fetch.enable", {
      patterns: [{ urlPattern: "*teams*", requestStage: "Request" }],
    });

    const tokenPromise = new Promise<void>((resolve) => {
      const timeout = setTimeout(resolve, FETCH_INTERCEPT_TIMEOUT);

      cdpSession.on(
        "Fetch.requestPaused",
        async (
          event: import("puppeteer-core").Protocol.Fetch.RequestPausedEvent,
        ) => {
          const headers = event.request.headers ?? {};

          for (const [name, value] of Object.entries(headers)) {
            if (name.toLowerCase() === "x-skypetoken" && !skypeToken) {
              skypeToken = value;
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

    await cdpSession.send("Fetch.disable");
    await cdpSession.detach();

    if (!skypeToken) {
      throw new Error(
        "Failed to capture skype token. Ensure Teams is loaded and authenticated.",
      );
    }

    return { skypeToken, region: "apac" };
  } finally {
    browser.disconnect();
  }
}
