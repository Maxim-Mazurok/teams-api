/**
 * Browser runtime helpers for interactive login.
 *
 * Interactive login depends on Chromium DevTools Protocol interception,
 * so we prefer installed Chromium-based browsers (Edge/Chrome) and
 * fall back to Playwright's bundled Chromium when needed.
 */

import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import type { Browser } from "playwright";

const requireFromHere = createRequire(__filename);

type LogFunction = (...arguments_: unknown[]) => void;
type InteractiveBrowserChannel = "chrome" | "msedge";

export interface ChromiumBrowserLauncher {
  launch(options: {
    headless: false;
    channel?: InteractiveBrowserChannel;
  }): Promise<Browser>;
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
