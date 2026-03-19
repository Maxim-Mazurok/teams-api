/**
 * Authentication module for Teams API.
 *
 * Three strategies:
 *   1. Interactive login: Open a visible browser, let the user log in manually,
 *      capture skype token via CDP Fetch interception. Works on all platforms.
 *   2. Auto-login: Launch system Chrome, complete FIDO2/passkey via platform authenticator,
 *      capture skype token from CDP Fetch interception. macOS only.
 *   3. Debug session: Connect to a running Chrome debug session with Teams open,
 *      intercept the token during a page reload.
 *
 * All strategies return a TeamsToken that can be used with the API layer.
 */

import type {
  TeamsToken,
  AutoLoginOptions,
  InteractiveLoginOptions,
  ManualTokenOptions,
} from "./types.js";
import { detectTeamsRegionFromUrl, resolveTeamsRegion } from "./region.js";

const TEAMS_URL = "https://teams.cloud.microsoft/";
const DEFAULT_SYSTEM_CHROME_PATH =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const DEFAULT_PROFILE_DIRECTORY = "/tmp/teams-api-chrome-profile";
const AUTO_LOGIN_TIMEOUT = 2 * 60 * 1_000;
const TOKEN_INTERCEPT_TIMEOUT = 60 * 1_000;
const FETCH_INTERCEPT_TIMEOUT = 30 * 1_000;
const PAGE_RELOAD_TIMEOUT = 30 * 1_000;
const BROWSER_LAUNCH_TIMEOUT = 30 * 1_000;
const PAGE_STATE_POLL_INTERVAL = 3 * 1_000;

type LogFunction = (...arguments_: unknown[]) => void;

/** Describes the current state of the page during the login flow. */
interface PageState {
  url: string;
  phase:
    | "login-page"
    | "mfa-challenge"
    | "redirect"
    | "teams-loading"
    | "teams-loaded"
    | "unknown";
  detail: string;
}

/**
 * Diagnose the current page state during authentication.
 * Returns a structured description of where the user is in the login flow.
 */
async function diagnosePageState(page: {
  url: () => string;
  evaluate: (script: string) => Promise<unknown>;
}): Promise<PageState> {
  const url = page.url();

  if (
    url.includes("login.microsoftonline.com") ||
    url.includes("login.microsoft.com")
  ) {
    const pageContent = (await page
      .evaluate("document.body?.innerText?.slice(0, 500) ?? ''")
      .catch(() => "")) as string;

    if (
      pageContent.includes("Approve sign in request") ||
      pageContent.includes("Enter code") ||
      pageContent.includes("Verify your identity") ||
      pageContent.includes("approve") ||
      url.includes("/common/SAS")
    ) {
      return {
        url,
        phase: "mfa-challenge",
        detail: "Waiting for MFA/2FA approval",
      };
    }

    return {
      url,
      phase: "login-page",
      detail: "On the Microsoft login page",
    };
  }

  if (
    url.includes("teams.cloud.microsoft") ||
    url.includes("teams.microsoft.com")
  ) {
    const isLoaded = (await page
      .evaluate(
        "(document.querySelector('#app') ?? document.querySelector(\"[data-tid='app-layout']\") ?? document.querySelector('main')) !== null",
      )
      .catch(() => false)) as boolean;

    if (isLoaded) {
      return {
        url,
        phase: "teams-loaded",
        detail: "Teams application loaded",
      };
    }

    return {
      url,
      phase: "teams-loading",
      detail: "Teams page reached, application loading",
    };
  }

  if (url.includes("microsoftonline.com") || url.includes("microsoft.com")) {
    return {
      url,
      phase: "redirect",
      detail: "Redirecting through Microsoft authentication",
    };
  }

  return { url, phase: "unknown", detail: `Unexpected page: ${url}` };
}

/**
 * Wait for the page to reach Teams and capture authentication tokens via CDP.
 *
 * Polls the page state periodically to provide diagnostic feedback instead of
 * relying solely on timeouts. Throws a descriptive error if the login flow
 * stalls or reaches an unexpected state.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function captureTokensFromPage(
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
    const { skypeToken, region, bearerToken, substrateToken } =
      await captureTokensFromPage(page, log, TOKEN_INTERCEPT_TIMEOUT);

    return {
      skypeToken,
      region: resolveTeamsRegion(options.region, region),
      bearerToken,
      substrateToken,
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
  const browser = await chromium.launch({ headless: false });
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

    const { skypeToken, region, bearerToken, substrateToken } =
      await captureTokensFromPage(page, log, TOKEN_INTERCEPT_TIMEOUT);

    return {
      skypeToken,
      region: resolveTeamsRegion(options?.region, region),
      bearerToken,
      substrateToken,
    };
  } finally {
    await context.close();
    await browser.close();
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
