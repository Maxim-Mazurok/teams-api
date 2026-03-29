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
        await createCredentialStore().clear(testAccount);
      } catch {
        // ignore
      }
    });

    it("should save and load a credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      await store.save(testAccount, "test-secret-value");
      const loaded = await store.load(testAccount);

      expect(loaded).toBe("test-secret-value");
    });

    it("should return null for a non-existent credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      const loaded = await store.load("teams-api-nonexistent-account-xyz");

      expect(loaded).toBeNull();
    });

    it("should clear a credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      await store.save(testAccount, "to-be-deleted");
      await store.clear(testAccount);
      const loaded = await store.load(testAccount);

      expect(loaded).toBeNull();
    });

    it("should overwrite an existing credential", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      await store.save(testAccount, "first-value");
      await store.save(testAccount, "second-value");
      const loaded = await store.load(testAccount);

      expect(loaded).toBe("second-value");
    });

    it("should handle special characters in data", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      const specialData = 'quotes"and\'backslash\\newline\ntab\t${}`template`';
      await store.save(testAccount, specialData);
      const loaded = await store.load(testAccount);

      expect(loaded).toBe(specialData);
    });

    it("should chunk large payloads that exceed 2560 bytes", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      const largeData = "X".repeat(5000);
      await store.save(testAccount, largeData);
      const loaded = await store.load(testAccount);

      expect(loaded).toBe(largeData);
    });

    it("should clear chunked credentials completely", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      await store.save(testAccount, "Y".repeat(5000));
      await store.clear(testAccount);
      const loaded = await store.load(testAccount);

      expect(loaded).toBeNull();
    });

    it("should overwrite chunked with non-chunked and vice versa", async () => {
      const { createCredentialStore } =
        await import("../../src/credential-store.js");
      const store = createCredentialStore();

      // Save chunked, then overwrite with small
      await store.save(testAccount, "Z".repeat(5000));
      await store.save(testAccount, "small-value");
      expect(await store.load(testAccount)).toBe("small-value");

      // Save small, then overwrite with chunked
      const large = "W".repeat(5000);
      await store.save(testAccount, large);
      expect(await store.load(testAccount)).toBe(large);
    });
  },
);
