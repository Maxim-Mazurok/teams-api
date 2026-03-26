/**
 * Unit tests for the platform detection module (src/platform.ts).
 */

import { describe, it, expect, vi, afterEach } from "vitest";

describe("detectPlatform", () => {
  const originalPlatform = process.platform;

  afterEach(() => {
    Object.defineProperty(process, "platform", { value: originalPlatform });
  });

  it("should return 'macos' on darwin", async () => {
    Object.defineProperty(process, "platform", { value: "darwin" });
    const { detectPlatform } = await import("../../src/platform.js");
    expect(detectPlatform()).toBe("macos");
  });

  it("should return 'windows' on win32", async () => {
    Object.defineProperty(process, "platform", { value: "win32" });
    const { detectPlatform } = await import("../../src/platform.js");
    expect(detectPlatform()).toBe("windows");
  });

  it("should return 'linux' on linux", async () => {
    Object.defineProperty(process, "platform", { value: "linux" });
    const { detectPlatform } = await import("../../src/platform.js");
    expect(detectPlatform()).toBe("linux");
  });

  it("should return 'linux' on freebsd", async () => {
    Object.defineProperty(process, "platform", { value: "freebsd" });
    const { detectPlatform } = await import("../../src/platform.js");
    expect(detectPlatform()).toBe("linux");
  });
});

describe("canAttemptAutoLogin", () => {
  const originalPlatform = process.platform;

  afterEach(() => {
    Object.defineProperty(process, "platform", { value: originalPlatform });
    vi.restoreAllMocks();
  });

  it("should return false on Windows", async () => {
    Object.defineProperty(process, "platform", { value: "win32" });
    const { canAttemptAutoLogin } = await import("../../src/platform.js");
    expect(canAttemptAutoLogin()).toBe(false);
  });

  it("should return false on Linux", async () => {
    Object.defineProperty(process, "platform", { value: "linux" });
    const { canAttemptAutoLogin } = await import("../../src/platform.js");
    expect(canAttemptAutoLogin()).toBe(false);
  });

  it("should return false on macOS when Chrome is not found", async () => {
    Object.defineProperty(process, "platform", { value: "darwin" });
    const { canAttemptAutoLogin } = await import("../../src/platform.js");
    // Pass a path that definitely doesn't exist
    expect(canAttemptAutoLogin("/nonexistent/chrome")).toBe(false);
  });
});
