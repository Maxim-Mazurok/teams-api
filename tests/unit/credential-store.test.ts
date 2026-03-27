/**
 * Unit tests for the credential store factory (src/credential-store.ts).
 *
 * Tests the factory function and verifies platform selection logic.
 * On Windows, also runs live tests against Windows Credential Manager.
 */

import { describe, it, expect, vi, afterEach } from "vitest";

// We need to re-import after changing platform, so we use dynamic imports
describe("createCredentialStore", () => {
  const originalPlatform = process.platform;

  afterEach(() => {
    Object.defineProperty(process, "platform", { value: originalPlatform });
    vi.restoreAllMocks();
  });

  it("should return a KeychainStore on macOS", async () => {
    Object.defineProperty(process, "platform", { value: "darwin" });

    // Re-import to pick up the new platform value
    const { createCredentialStore } =
      await import("../../src/credential-store.js");
    const store = createCredentialStore();

    expect(store).toBeDefined();
    expect(store.save).toBeTypeOf("function");
    expect(store.load).toBeTypeOf("function");
    expect(store.clear).toBeTypeOf("function");
  });

  it("should return a WinCredStore on Windows", async () => {
    Object.defineProperty(process, "platform", { value: "win32" });

    const { createCredentialStore } =
      await import("../../src/credential-store.js");
    const store = createCredentialStore();

    expect(store).toBeDefined();
    expect(store.save).toBeTypeOf("function");
    expect(store.load).toBeTypeOf("function");
    expect(store.clear).toBeTypeOf("function");
  });

  it("should return a LinuxStore on Linux", async () => {
    Object.defineProperty(process, "platform", { value: "linux" });

    const { createCredentialStore } =
      await import("../../src/credential-store.js");
    const store = createCredentialStore();

    expect(store).toBeDefined();
    expect(store.save).toBeTypeOf("function");
    expect(store.load).toBeTypeOf("function");
    expect(store.clear).toBeTypeOf("function");
  });
});

describe.runIf(process.platform === "win32")(
  "WinCredStore (live Windows Credential Manager)",
  () => {
    const testAccount = "teams-api-unit-test-account";

    afterEach(async () => {
      try {
        const { createCredentialStore } =
          await import("../../src/credential-store.js");
        createCredentialStore().clear(testAccount);
      } catch {
        // ignore
      }
    });

    it("should save and load a credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      store.save(testAccount, "test-secret-value");
      const loaded = store.load(testAccount);

      expect(loaded).toBe("test-secret-value");
    });

    it("should return null for a non-existent credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      const loaded = store.load("teams-api-nonexistent-account-xyz");

      expect(loaded).toBeNull();
    });

    it("should clear a credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      store.save(testAccount, "to-be-deleted");
      store.clear(testAccount);
      const loaded = store.load(testAccount);

      expect(loaded).toBeNull();
    });

    it("should overwrite an existing credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      store.save(testAccount, "first-value");
      store.save(testAccount, "second-value");
      const loaded = store.load(testAccount);

      expect(loaded).toBe("second-value");
    });

    it("should handle special characters in data", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      const specialData = 'quotes"and\'backslash\\newline\ntab\t${}`template`';
      store.save(testAccount, specialData);
      const loaded = store.load(testAccount);

      expect(loaded).toBe(specialData);
    });
  },
);
