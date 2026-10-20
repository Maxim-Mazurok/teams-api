/**
 * Secure token persistence using the platform credential store.
 *
 * Stores Teams tokens via the cross-platform credential store:
 * macOS Keychain, Windows DPAPI, or Linux secret-tool/file.
 * Tokens are base64-encoded JSON with an acquisition timestamp.
 * Expired tokens (older than TOKEN_LIFETIME) are automatically cleared.
 */

import type { TeamsToken } from "./types.js";
import { createCredentialStore } from "./credential-store.js";

const TOKEN_LIFETIME = 23 * 60 * 60 * 1_000; // 23 hours (tokens last ~24h, 1h safety margin)
const JWT_EXPIRY_MARGIN = 5 * 60 * 1_000;

interface StoredToken {
  skypeToken: string;
  region: string;
  chatBearerToken?: string;
  bearerToken?: string;
  substrateToken?: string;
  amsToken?: string;
  sharePointToken?: string;
  sharePointHost?: string;
  acquiredAt: number;
}

const store = createCredentialStore();

/** True when the token is a JWT whose `exp` claim is (almost) reached. */
function isJwtExpired(token: string): boolean {
  try {
    const payload = JSON.parse(
      Buffer.from(token.split(".")[1], "base64url").toString("utf-8"),
    ) as { exp?: number };
    return (
      typeof payload.exp === "number" &&
      payload.exp * 1_000 - Date.now() < JWT_EXPIRY_MARGIN
    );
  } catch {
    return false;
  }
}

export async function saveToken(
  email: string,
  token: TeamsToken,
): Promise<void> {
  const storedToken: StoredToken = {
    skypeToken: token.skypeToken,
    region: token.region,
    chatBearerToken: token.chatBearerToken,
    bearerToken: token.bearerToken,
    substrateToken: token.substrateToken,
    amsToken: token.amsToken,
    sharePointToken: token.sharePointToken,
    sharePointHost: token.sharePointHost,
    acquiredAt: Date.now(),
  };
  const encoded = Buffer.from(JSON.stringify(storedToken)).toString("base64");

  await store.save(email, encoded);
}

export async function loadToken(email: string): Promise<TeamsToken | null> {
  const encoded = await store.load(email);
  if (!encoded) {
    return null;
  }

  let storedToken: StoredToken;
  try {
    storedToken = JSON.parse(
      Buffer.from(encoded, "base64").toString("utf-8"),
    ) as StoredToken;
  } catch {
    await clearToken(email);
    return null;
  }

  if (Date.now() - storedToken.acquiredAt > TOKEN_LIFETIME) {
    await clearToken(email);
    return null;
  }

  if (
    !storedToken.skypeToken &&
    storedToken.chatBearerToken &&
    isJwtExpired(storedToken.chatBearerToken)
  ) {
    await clearToken(email);
    return null;
  }

  return {
    skypeToken: storedToken.skypeToken,
    region: storedToken.region,
    chatBearerToken: storedToken.chatBearerToken,
    bearerToken: storedToken.bearerToken,
    substrateToken: storedToken.substrateToken,
    amsToken: storedToken.amsToken,
    sharePointToken: storedToken.sharePointToken,
    sharePointHost: storedToken.sharePointHost,
  };
}

export async function clearToken(email: string): Promise<void> {
  await store.clear(email);
}
