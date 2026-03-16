/**
 * Secure token persistence using the macOS Keychain.
 *
 * Stores Teams tokens in the system Keychain via the `security` CLI tool.
 * Tokens are base64-encoded JSON with an acquisition timestamp.
 * Expired tokens (older than TOKEN_LIFETIME) are automatically cleared.
 *
 * Uses `execFileSync` (not `execSync`) to avoid shell injection risks.
 */

import { execFileSync } from "node:child_process";
import type { TeamsToken } from "./types.js";

const KEYCHAIN_SERVICE = "teams-api";
const TOKEN_LIFETIME = 23 * 60 * 60 * 1_000; // 23 hours (tokens last ~24h, 1h safety margin)

interface StoredToken {
  skypeToken: string;
  region: string;
  acquiredAt: number;
}

export function saveToken(email: string, token: TeamsToken): void {
  const storedToken: StoredToken = {
    skypeToken: token.skypeToken,
    region: token.region,
    acquiredAt: Date.now(),
  };
  const encoded = Buffer.from(JSON.stringify(storedToken)).toString("base64");

  execFileSync("security", [
    "add-generic-password",
    "-a",
    email,
    "-s",
    KEYCHAIN_SERVICE,
    "-w",
    encoded,
    "-U",
  ]);
}

export function loadToken(email: string): TeamsToken | null {
  let encoded: string;
  try {
    encoded = execFileSync(
      "security",
      ["find-generic-password", "-a", email, "-s", KEYCHAIN_SERVICE, "-w"],
      { encoding: "utf-8" },
    ).trim();
  } catch {
    return null;
  }

  let storedToken: StoredToken;
  try {
    storedToken = JSON.parse(
      Buffer.from(encoded, "base64").toString("utf-8"),
    ) as StoredToken;
  } catch {
    clearToken(email);
    return null;
  }

  if (Date.now() - storedToken.acquiredAt > TOKEN_LIFETIME) {
    clearToken(email);
    return null;
  }

  return { skypeToken: storedToken.skypeToken, region: storedToken.region };
}

export function clearToken(email: string): void {
  try {
    execFileSync("security", [
      "delete-generic-password",
      "-a",
      email,
      "-s",
      KEYCHAIN_SERVICE,
    ]);
  } catch {
    // Token may not exist in keychain, that's fine
  }
}
