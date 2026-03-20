/**
 * Platform detection utilities for smart login flow.
 *
 * Determines the current platform and whether auto-login prerequisites
 * are met (macOS + system Chrome installed).
 */

import { accessSync, constants } from "node:fs";

const DEFAULT_SYSTEM_CHROME_PATH =
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

export type Platform = "macos" | "windows" | "linux";

export function detectPlatform(): Platform {
  switch (process.platform) {
    case "darwin":
      return "macos";
    case "win32":
      return "windows";
    default:
      return "linux";
  }
}

/**
 * Check whether auto-login prerequisites are met:
 * macOS + system Chrome exists at the expected path.
 *
 * Does NOT check for FIDO2 passkey enrollment — auto-login will
 * simply fail and the smart login flow will fall back to interactive.
 */
export function canAttemptAutoLogin(chromePath?: string): boolean {
  if (detectPlatform() !== "macos") {
    return false;
  }

  const chromeExecutable = chromePath ?? DEFAULT_SYSTEM_CHROME_PATH;
  try {
    accessSync(chromeExecutable, constants.X_OK);
    return true;
  } catch {
    return false;
  }
}
