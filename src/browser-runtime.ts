/**
 * Browser runtime helpers for interactive login.
 *
 * Interactive login depends on Chromium DevTools Protocol interception,
 * so we prefer installed Chromium-based browsers (Edge/Chrome) and
 * fall back to Playwright's bundled Chromium when needed.
 */

import { spawnSync } from "node:child_process";
import { lstatSync, readlinkSync, rmSync } from "node:fs";
import { createRequire } from "node:module";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import type { Browser, BrowserContext } from "playwright";

const requireFromHere = createRequire(__filename);

type LogFunction = (...arguments_: unknown[]) => void;
type InteractiveBrowserChannel = "chrome" | "msedge";

export interface ChromiumBrowserLauncher {
  launch(options: {
    headless: false;
    channel?: InteractiveBrowserChannel;
  }): Promise<Browser>;
}

export interface PersistentBrowserLauncher extends ChromiumBrowserLauncher {
  launchPersistentContext(
    userDataDir: string,
    options: {
      headless: false;
      channel?: InteractiveBrowserChannel;
    },
  ): Promise<BrowserContext>;
}

export function getDefaultBrowserProfileDir(): string {
  return join(homedir(), ".teams-api", "browser-profile");
}

const INTERACTIVE_BROWSER_CHANNELS: Partial<
  Record<NodeJS.Platform, InteractiveBrowserChannel[]>
> = {
  darwin: ["chrome"],
  win32: ["msedge", "chrome"],
  linux: ["chrome", "msedge"],
};

function formatChannelName(channel: InteractiveBrowserChannel): string {
  switch (channel) {
    case "msedge":
      return "Microsoft Edge";
    case "chrome":
      return "Google Chrome";
  }
}

export function getInteractiveBrowserChannels(
  platform: NodeJS.Platform = process.platform,
): InteractiveBrowserChannel[] {
  return INTERACTIVE_BROWSER_CHANNELS[platform] ?? ["chrome", "msedge"];
}

export function isMissingPlaywrightBrowserError(error: unknown): boolean {
  return (
    error instanceof Error && error.message.includes("Executable doesn't exist")
  );
}

export function installBundledChromium(log: LogFunction): void {
  const playwrightPackageJson = requireFromHere.resolve(
    "playwright/package.json",
  );
  const playwrightCliPath = join(dirname(playwrightPackageJson), "cli.js");

  log(
    "Playwright Chromium is not installed. Downloading it now (one-time setup)...",
  );

  const result = spawnSync(
    process.execPath,
    [playwrightCliPath, "install", "chromium"],
    {
      encoding: "utf-8",
    },
  );

  if (result.status === 0) {
    log("Playwright Chromium installed successfully.");
    return;
  }

  const detail =
    result.error?.message ||
    result.stderr?.trim() ||
    result.stdout?.trim() ||
    "Unknown error";

  throw new Error(
    `Failed to install Playwright Chromium automatically. ${detail}`,
  );
}

const PROFILE_LOCK_MAX_RETRIES = 3;
const PROFILE_LOCK_RETRY_DELAY_MS = 500;

/**
 * Check if a file or symlink exists (without following the symlink).
 * Unlike `existsSync`, this returns true for broken symlinks.
 */
function fileOrSymlinkExists(path: string): boolean {
  try {
    lstatSync(path);
    return true;
  } catch {
    return false;
  }
}

/**
 * Detect whether a browser launch error is caused by a locked profile
 * (Chromium exit code 21 or SingletonLock contention).
 */
export function isProfileLockError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  const message = error.message.toLowerCase();
  return (
    message.includes("exit code 21") ||
    message.includes("process singleton") ||
    message.includes("profile directory is already in use") ||
    message.includes("singletonlock")
  );
}

/**
 * Attempt to remove a stale SingletonLock file from a browser profile.
 * Only removes the lock if the PID it references is not running.
 * Returns true if the lock was successfully cleaned.
 */
export function cleanStaleSingletonLock(
  profileDir: string,
  log: LogFunction,
): boolean {
  const lockPath = join(profileDir, "SingletonLock");
  if (!fileOrSymlinkExists(lockPath)) {
    log("No SingletonLock file found — nothing to clean.");
    return false;
  }

  // SingletonLock is a symlink whose target contains the hostname and PID
  // Format: "hostname-pid" (e.g., "MacBook-Pro.local-12345")
  try {
    const linkTarget = readlinkSync(lockPath);
    const pidMatch = linkTarget.match(/-(\d+)$/);
    if (pidMatch) {
      const pid = Number(pidMatch[1]);
      if (isProcessRunning(pid)) {
        log(`SingletonLock held by live process (PID ${pid}) — cannot remove.`);
        return false;
      }
      log(`SingletonLock held by dead process (PID ${pid}) — removing.`);
    } else {
      log("SingletonLock has unrecognized format — removing.");
    }
  } catch {
    log("Could not read SingletonLock — attempting removal anyway.");
  }

  try {
    rmSync(lockPath, { force: true });
    log("Removed stale SingletonLock.");
    return true;
  } catch (removeError) {
    log(`Failed to remove SingletonLock: ${(removeError as Error).message}`);
    return false;
  }
}

function isProcessRunning(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

export async function launchInteractiveBrowser(
  chromium: ChromiumBrowserLauncher,
  log: LogFunction,
  options?: {
    platform?: NodeJS.Platform;
    installBundledChromium?: (log: LogFunction) => void;
  },
): Promise<Browser> {
  for (const channel of getInteractiveBrowserChannels(options?.platform)) {
    try {
      log(`Trying installed ${formatChannelName(channel)}...`);
      const browser = await chromium.launch({
        headless: false,
        channel,
      });
      log(
        `Using installed ${formatChannelName(channel)} for interactive login.`,
      );
      return browser;
    } catch (error) {
      log(
        `Could not launch ${formatChannelName(channel)}: ${(error as Error).message}`,
      );
    }
  }

  try {
    log("Trying Playwright bundled Chromium...");
    const browser = await chromium.launch({ headless: false });
    log("Using Playwright bundled Chromium for interactive login.");
    return browser;
  } catch (error) {
    if (!isMissingPlaywrightBrowserError(error)) {
      throw error;
    }

    const installChromium =
      options?.installBundledChromium ?? installBundledChromium;
    installChromium(log);

    log("Retrying Playwright bundled Chromium after install...");
    return chromium.launch({ headless: false });
  }
}

export async function launchInteractiveBrowserContext(
  chromium: PersistentBrowserLauncher,
  log: LogFunction,
  userDataDir: string,
  options?: {
    platform?: NodeJS.Platform;
    installBundledChromium?: (log: LogFunction) => void;
  },
): Promise<BrowserContext> {
  for (let attempt = 0; attempt <= PROFILE_LOCK_MAX_RETRIES; attempt++) {
    try {
      return await attemptLaunchPersistentContext(
        chromium,
        log,
        userDataDir,
        options,
      );
    } catch (error) {
      if (!isProfileLockError(error) || attempt === PROFILE_LOCK_MAX_RETRIES) {
        throw error;
      }
      log(
        `Browser launch failed with profile lock error (attempt ${attempt + 1}/${PROFILE_LOCK_MAX_RETRIES + 1}): ${(error as Error).message}`,
      );
      const cleaned = cleanStaleSingletonLock(userDataDir, log);
      if (!cleaned && attempt === PROFILE_LOCK_MAX_RETRIES - 1) {
        throw new Error(
          `Browser profile at "${userDataDir}" is locked and could not be cleaned. ` +
            `Try closing other browser instances or deleting "${join(userDataDir, "SingletonLock")}" manually.`,
        );
      }
      await new Promise((resolve) =>
        setTimeout(resolve, PROFILE_LOCK_RETRY_DELAY_MS),
      );
    }
  }

  // Unreachable, but satisfies TypeScript
  throw new Error("Browser launch failed after all retries");
}

async function attemptLaunchPersistentContext(
  chromium: PersistentBrowserLauncher,
  log: LogFunction,
  userDataDir: string,
  options?: {
    platform?: NodeJS.Platform;
    installBundledChromium?: (log: LogFunction) => void;
  },
): Promise<BrowserContext> {
  for (const channel of getInteractiveBrowserChannels(options?.platform)) {
    try {
      log(`Trying installed ${formatChannelName(channel)}...`);
      const context = await chromium.launchPersistentContext(userDataDir, {
        headless: false,
        channel,
      });
      log(
        `Using installed ${formatChannelName(channel)} for interactive login.`,
      );
      return context;
    } catch (error) {
      if (isProfileLockError(error)) throw error;
      log(
        `Could not launch ${formatChannelName(channel)}: ${(error as Error).message}`,
      );
    }
  }

  try {
    log("Trying Playwright bundled Chromium...");
    const context = await chromium.launchPersistentContext(userDataDir, {
      headless: false,
    });
    log("Using Playwright bundled Chromium for interactive login.");
    return context;
  } catch (error) {
    if (isProfileLockError(error)) throw error;
    if (!isMissingPlaywrightBrowserError(error)) {
      throw error;
    }

    const installChromium =
      options?.installBundledChromium ?? installBundledChromium;
    installChromium(log);

    log("Retrying Playwright bundled Chromium after install...");
    return chromium.launchPersistentContext(userDataDir, { headless: false });
  }
}
