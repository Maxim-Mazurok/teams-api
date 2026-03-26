/**
 * Smart login — the zero-config default authentication strategy.
 *
 * Strategy:
 *   1. If on macOS + email provided + system Chrome exists → try auto-login
 *   2. If auto-login fails or prerequisites not met → interactive login
 *
 * Auto-login errors are caught and logged, then we immediately fall back
 * to interactive. Only if interactive also fails does the user see an error.
 */

import type { TeamsToken, SmartLoginOptions } from "./types.js";
import { canAttemptAutoLogin } from "./platform.js";
import { acquireTokenViaAutoLogin } from "./auth/auto-login.js";
import { acquireTokenViaInteractiveLogin } from "./auth/interactive.js";

type LogFunction = (...arguments_: unknown[]) => void;

export async function acquireTokenViaSmartLogin(
  options?: SmartLoginOptions,
): Promise<TeamsToken> {
  const log: LogFunction = options?.verbose
    ? console.error.bind(console)
    : () => {};

  // Try auto-login if prerequisites are met
  if (options?.email && canAttemptAutoLogin()) {
    log("Auto-login prerequisites met (macOS + Chrome), attempting...");
    try {
      return await acquireTokenViaAutoLogin({
        email: options.email,
        region: options.region,
        headless: true,
        verbose: options.verbose,
      });
    } catch (error) {
      log(
        `Auto-login failed: ${(error as Error).message}. Falling back to interactive login...`,
      );
    }
  }

  // Fall back to interactive login (works everywhere)
  log("Using interactive browser login...");
  return acquireTokenViaInteractiveLogin({
    region: options?.region,
    email: options?.email,
    verbose: options?.verbose,
  });
}
