/**
 * Unit tests for the token store (src/token-store.ts).
 *
 * These tests mock the credential store to verify that the token store
 * correctly serializes/deserializes tokens and handles expiry.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import type { TeamsToken } from "../../src/types.js";

const mockStore = vi.hoisted(() => ({
  save: vi.fn(),
  load: vi.fn(),
  clear: vi.fn(),
}));

vi.mock("../../src/credential-store.js", () => ({
  createCredentialStore: () => mockStore,
}));

import { saveToken, loadToken, clearToken } from "../../src/token-store.js";

const testEmail = "user@company.com";
const testToken: TeamsToken = {
  skypeToken: "test-skype-token-abc123",
  region: "apac",
};

const testTokenWithAuxiliaryTokens: TeamsToken = {
  skypeToken: "test-skype-token-abc123",
  region: "apac",
  bearerToken: "test-bearer-token",
  substrateToken: "test-substrate-token",
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
  it("should call credential store save with correct arguments", async () => {
    await saveToken(testEmail, testToken);

    expect(mockStore.save).toHaveBeenCalledTimes(1);
    expect(mockStore.save).toHaveBeenCalledWith(
      testEmail,
      expect.any(String),
    );
  });

  it("should store base64-encoded JSON with acquiredAt timestamp", async () => {
    await saveToken(testEmail, testToken);

    const encodedValue = mockStore.save.mock.calls[0][1] as string;
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

  it("should persist optional bearer and substrate tokens", async () => {
    await saveToken(testEmail, testTokenWithAuxiliaryTokens);

    const encodedValue = mockStore.save.mock.calls[0][1] as string;
    const decoded = JSON.parse(
      Buffer.from(encodedValue, "base64").toString("utf-8"),
    ) as {
      skypeToken: string;
      region: string;
      bearerToken?: string;
      substrateToken?: string;
      acquiredAt: number;
    };

    expect(decoded.bearerToken).toBe("test-bearer-token");
    expect(decoded.substrateToken).toBe("test-substrate-token");
  });
});

describe("loadToken", () => {
  it("should return token when cached and not expired", async () => {
    const encoded = makeStoredTokenBase64();
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toEqual(testToken);
    expect(mockStore.load).toHaveBeenCalledWith(testEmail);
  });

  it("should restore optional bearer and substrate tokens", async () => {
    const encoded = Buffer.from(
      JSON.stringify({
        skypeToken: testTokenWithAuxiliaryTokens.skypeToken,
        region: testTokenWithAuxiliaryTokens.region,
        bearerToken: testTokenWithAuxiliaryTokens.bearerToken,
        substrateToken: testTokenWithAuxiliaryTokens.substrateToken,
        acquiredAt: new Date("2026-03-17T12:00:00.000Z").getTime(),
      }),
    ).toString("base64");
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toEqual(testTokenWithAuxiliaryTokens);
  });

  it("should return null when no token exists in credential store", async () => {
    mockStore.load.mockResolvedValueOnce(null);

    const result = await loadToken(testEmail);

    expect(result).toBeNull();
  });

  it("should return null and clear token when expired (older than 23 hours)", async () => {
    const twentyFourHoursAgo =
      new Date("2026-03-17T12:00:00.000Z").getTime() - 24 * 60 * 60 * 1_000;
    const encoded = makeStoredTokenBase64({ acquiredAt: twentyFourHoursAgo });
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toBeNull();
    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
  });

  it("should return token when acquired exactly 22 hours ago", async () => {
    const twentyTwoHoursAgo =
      new Date("2026-03-17T12:00:00.000Z").getTime() - 22 * 60 * 60 * 1_000;
    const encoded = makeStoredTokenBase64({
      acquiredAt: twentyTwoHoursAgo,
    });
    mockStore.load.mockResolvedValueOnce(encoded);

    const result = await loadToken(testEmail);

    expect(result).toEqual(testToken);
  });

  it("should return null and clear token when data is corrupted", async () => {
    mockStore.load.mockResolvedValueOnce("not-valid-base64!!!");

    const result = await loadToken(testEmail);

    expect(result).toBeNull();
    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
  });
});

describe("clearToken", () => {
  it("should call credential store clear", async () => {
    await clearToken(testEmail);

    expect(mockStore.clear).toHaveBeenCalledWith(testEmail);
  });

  it("should not throw when token does not exist in credential store", async () => {
    mockStore.clear.mockImplementation(() => {
      // no-op, simulating non-existent entry
    });

    await expect(clearToken(testEmail)).resolves.not.toThrow();
  });
});
