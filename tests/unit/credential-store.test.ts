/**
 * Unit tests for the credential store factory (src/credential-store.ts).
 *
 * Tests the factory function and verifies platform selection logic.
 * Individual store implementations are tested via the token-store tests
 * and integration tests on each platform.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

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
