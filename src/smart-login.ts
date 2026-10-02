/**
 * Smart login — headed interactive authentication by default.
 *
 * Strategy:
 *   1. If FIDO2 is explicitly enabled and prerequisites exist → try auto-login
 *   2. If auto-login fails or prerequisites not met → interactive login
 *
 * Auto-login errors are caught and logged, then we immediately fall back
 * to interactive. Only if interactive also fails does the user see an error.
 */

import type {
  AuthLogFunction,
  TeamsToken,
  SmartLoginOptions,
} from "./types.js";
import { canAttemptAutoLogin } from "./platform.js";
import { acquireTokenViaAutoLogin } from "./auth/auto-login.js";
import { acquireTokenViaInteractiveLogin } from "./auth/interactive.js";

export async function acquireTokenViaSmartLogin(
  options?: SmartLoginOptions,
): Promise<TeamsToken> {
  const log: AuthLogFunction =
    options?.log ?? (options?.verbose ? console.error.bind(console) : () => {});

  // Try FIDO2 auto-login only when explicitly enabled.
  if (options?.auto && options.email && canAttemptAutoLogin()) {
    log("Auto-login prerequisites met (macOS + Chrome), attempting...");
    try {
      return await acquireTokenViaAutoLogin({
        email: options.email,
        region: options.region,
        headless: options.headless ?? false,
        verbose: options.verbose,
        log,
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
    log,
  });
}
