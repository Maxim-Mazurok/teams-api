/**
 * Windows E2E tests — credential store and token persistence.
 *
 * Tests the Windows Credential Manager integration including chunked
 * storage for large payloads. No browser or network access needed.
 *
 * Skipped by default. Run on Windows with:
 *   TEAMS_TEST_WINDOWS=1 npx vitest run tests/e2e/windows-credential-store.test.ts
 */

import { describe, it, expect, afterAll } from "vitest";
import { createCredentialStore } from "../../src/credential-store.js";
import { saveToken, loadToken, clearToken } from "../../src/token-store.js";

const shouldRun =
  process.platform === "win32" && Boolean(process.env["TEAMS_TEST_WINDOWS"]);

const testAccounts: string[] = [];

afterAll(async () => {
  const store = createCredentialStore();
  for (const account of testAccounts) {
    await store.clear(account).catch(() => {});
  }
});

function track(account: string): string {
  testAccounts.push(account);
  return account;
}

describe.skipIf(!shouldRun)(
  "Windows Credential Store",
  { timeout: 30_000 },
  () => {
    const store = createCredentialStore();

    it("should save and load a small payload", async () => {
      const account = track("teams-api-e2e-small");
      await store.save(account, "hello-world");
      expect(await store.load(account)).toBe("hello-world");
    });

    it("should clear a small payload", async () => {
      const account = track("teams-api-e2e-small-clear");
      await store.save(account, "to-be-cleared");
      await store.clear(account);
      expect(await store.load(account)).toBeNull();
    });

    it("should save and load a large payload via chunking (5000 chars)", async () => {
      const account = track("teams-api-e2e-large");
      const data = "X".repeat(5000);
      await store.save(account, data);
      expect(await store.load(account)).toBe(data);
    });

    it("should clear all chunks of a large payload", async () => {
      const account = track("teams-api-e2e-large-clear");
      await store.save(account, "X".repeat(5000));
      await store.clear(account);
      expect(await store.load(account)).toBeNull();
    });

    it("should save and load a huge payload via chunking (8000 chars)", async () => {
      const account = track("teams-api-e2e-huge");
      const data = "Y".repeat(8000);
      await store.save(account, data);
      expect(await store.load(account)).toBe(data);
    });

    it("should overwrite a chunked payload with a small one", async () => {
      const account = track("teams-api-e2e-overwrite-1");
      await store.save(account, "Z".repeat(8000));
      await store.save(account, "tiny");
      expect(await store.load(account)).toBe("tiny");
    });

    it("should overwrite a small payload with a chunked one", async () => {
      const account = track("teams-api-e2e-overwrite-2");
      const huge = "W".repeat(8000);
      await store.save(account, "tiny");
      await store.save(account, huge);
      expect(await store.load(account)).toBe(huge);
    });
  },
);

describe.skipIf(!shouldRun)(
  "Windows Token Store",
  { timeout: 30_000 },
  () => {
    const tokenKey = "teams-api-e2e-token";

    afterAll(async () => {
      await clearToken(tokenKey).catch(() => {});
    });

    it("should save and load a realistic token payload", async () => {
      const token = {
        skypeToken: "sk_" + "a".repeat(1500),
        region: "apac",
        bearerToken: "bt_" + "b".repeat(1500),
        substrateToken: "st_" + "c".repeat(500),
      };
      await saveToken(tokenKey, token as any);
      const loaded = await loadToken(tokenKey);

      expect(loaded).not.toBeNull();
      expect(loaded!.skypeToken).toBe(token.skypeToken);
      expect(loaded!.bearerToken).toBe(token.bearerToken);
      expect(loaded!.substrateToken).toBe(token.substrateToken);
      expect(loaded!.region).toBe(token.region);
    });

    it("should clear a saved token", async () => {
      await saveToken(tokenKey, {
        skypeToken: "temp",
        region: "amer",
      } as any);
      await clearToken(tokenKey);
      expect(await loadToken(tokenKey)).toBeNull();
    });
  },
);
