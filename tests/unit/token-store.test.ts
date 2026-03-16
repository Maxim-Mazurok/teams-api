/**
 * Unit tests for the token store (src/token-store.ts).
 *
 * These tests mock `execFileSync` from `node:child_process` to verify
 * that the token store correctly interacts with the macOS Keychain.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import type { TeamsToken } from "../../src/types.js";

vi.mock("node:child_process");

import { execFileSync } from "node:child_process";
import { saveToken, loadToken, clearToken } from "../../src/token-store.js";

const mockedExecFileSync = vi.mocked(execFileSync);

const testEmail = "user@company.com";
const testToken: TeamsToken = {
  skypeToken: "test-skype-token-abc123",
  region: "apac",
};

beforeEach(() => {
  vi.resetAllMocks();
  vi.useFakeTimers();
  vi.setSystemTime(new Date("2026-03-17T12:00:00.000Z"));
});

afterEach(() => {
  vi.useRealTimers();
});

function makeStoredTokenBase64(
  overrides: { acquiredAt?: number } = {},
): string {
  const storedToken = {
    skypeToken: testToken.skypeToken,
    region: testToken.region,
    acquiredAt:
      overrides.acquiredAt ?? new Date("2026-03-17T12:00:00.000Z").getTime(),
  };
  return Buffer.from(JSON.stringify(storedToken)).toString("base64");
}

describe("saveToken", () => {
  it("should call security CLI with correct arguments", () => {
    saveToken(testEmail, testToken);

    expect(mockedExecFileSync).toHaveBeenCalledTimes(1);
    expect(mockedExecFileSync).toHaveBeenCalledWith("security", [
      "add-generic-password",
      "-a",
      testEmail,
      "-s",
      "teams-api",
      "-w",
      expect.any(String),
      "-U",
    ]);
  });

  it("should store base64-encoded JSON with acquiredAt timestamp", () => {
    saveToken(testEmail, testToken);

    const encodedValue = (mockedExecFileSync as ReturnType<typeof vi.fn>).mock
      .calls[0][1][6] as string;
    const decoded = JSON.parse(
      Buffer.from(encodedValue, "base64").toString("utf-8"),
    ) as {
      skypeToken: string;
      region: string;
      acquiredAt: number;
    };

    expect(decoded.skypeToken).toBe(testToken.skypeToken);
    expect(decoded.region).toBe(testToken.region);
    expect(decoded.acquiredAt).toBe(
      new Date("2026-03-17T12:00:00.000Z").getTime(),
    );
  });
});

describe("loadToken", () => {
  it("should return token when cached and not expired", () => {
    const encoded = makeStoredTokenBase64();
    mockedExecFileSync.mockReturnValueOnce(encoded as never);

    const result = loadToken(testEmail);

    expect(result).toEqual(testToken);
    expect(mockedExecFileSync).toHaveBeenCalledWith(
      "security",
      ["find-generic-password", "-a", testEmail, "-s", "teams-api", "-w"],
      { encoding: "utf-8" },
    );
  });

  it("should return null when no token exists in keychain", () => {
    mockedExecFileSync.mockImplementation(() => {
      throw new Error(
        "security: SecKeychainSearchCopyNext: The specified item could not be found in the keychain.",
      );
    });

    const result = loadToken(testEmail);

    expect(result).toBeNull();
  });

  it("should return null and clear token when expired (older than 23 hours)", () => {
    const twentyFourHoursAgo =
      new Date("2026-03-17T12:00:00.000Z").getTime() - 24 * 60 * 60 * 1_000;
    const encoded = makeStoredTokenBase64({ acquiredAt: twentyFourHoursAgo });
    mockedExecFileSync.mockReturnValueOnce(encoded as never);

    const result = loadToken(testEmail);

    expect(result).toBeNull();
    // Should have called delete-generic-password to clear expired token
    expect(mockedExecFileSync).toHaveBeenCalledWith("security", [
      "delete-generic-password",
      "-a",
      testEmail,
      "-s",
      "teams-api",
    ]);
  });

  it("should return token when acquired exactly 22 hours ago", () => {
    const twentyTwoHoursAgo =
      new Date("2026-03-17T12:00:00.000Z").getTime() - 22 * 60 * 60 * 1_000;
    const encoded = makeStoredTokenBase64({
      acquiredAt: twentyTwoHoursAgo,
    });
    mockedExecFileSync.mockReturnValueOnce(encoded as never);

    const result = loadToken(testEmail);

    expect(result).toEqual(testToken);
  });

  it("should return null and clear token when data is corrupted", () => {
    mockedExecFileSync.mockReturnValueOnce("not-valid-base64!!!" as never);

    const result = loadToken(testEmail);

    expect(result).toBeNull();
    // Should have tried to clear the corrupted entry
    expect(mockedExecFileSync).toHaveBeenCalledWith("security", [
      "delete-generic-password",
      "-a",
      testEmail,
      "-s",
      "teams-api",
    ]);
  });
});

describe("clearToken", () => {
  it("should call security CLI to delete the keychain entry", () => {
    clearToken(testEmail);

    expect(mockedExecFileSync).toHaveBeenCalledWith("security", [
      "delete-generic-password",
      "-a",
      testEmail,
      "-s",
      "teams-api",
    ]);
  });

  it("should not throw when token does not exist in keychain", () => {
    mockedExecFileSync.mockImplementation(() => {
      throw new Error(
        "security: SecKeychainSearchCopyNext: The specified item could not be found in the keychain.",
      );
    });

    expect(() => clearToken(testEmail)).not.toThrow();
  });
});
