/**
 * Unit tests for the smart login flow (src/smart-login.ts).
 *
 * Tests the decision logic: when to try auto-login vs interactive,
 * and the fallback behavior when auto-login fails.
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import type { TeamsToken } from "../../src/types.js";

vi.mock("../../src/platform.js");
vi.mock("../../src/auth/auto-login.js");
vi.mock("../../src/auth/interactive.js");

import { acquireTokenViaSmartLogin } from "../../src/smart-login.js";
import * as autoLogin from "../../src/auth/auto-login.js";
import * as interactive from "../../src/auth/interactive.js";
import * as platform from "../../src/platform.js";

const mockedAutoLogin = vi.mocked(autoLogin);
const mockedInteractive = vi.mocked(interactive);
const mockedPlatform = vi.mocked(platform);

const testToken: TeamsToken = {
  skypeToken: "test-token",
  region: "apac",
  bearerToken: "bearer",
  substrateToken: "substrate",
};

beforeEach(() => {
  vi.resetAllMocks();
});

describe("acquireTokenViaSmartLogin", () => {
  it("should use headed interactive login when only email is provided", async () => {
    mockedPlatform.canAttemptAutoLogin.mockReturnValue(true);
    mockedInteractive.acquireTokenViaInteractiveLogin.mockResolvedValue(
      testToken,
    );

    const result = await acquireTokenViaSmartLogin({
      email: "user@company.com",
    });

    expect(result).toEqual(testToken);
    expect(mockedAutoLogin.acquireTokenViaAutoLogin).not.toHaveBeenCalled();
    expect(
      mockedInteractive.acquireTokenViaInteractiveLogin,
    ).toHaveBeenCalledWith(
      expect.objectContaining({ email: "user@company.com" }),
    );
  });

  it("should try headed auto-login when FIDO2 is explicitly enabled", async () => {
    mockedPlatform.canAttemptAutoLogin.mockReturnValue(true);
    mockedAutoLogin.acquireTokenViaAutoLogin.mockResolvedValue(testToken);

    const result = await acquireTokenViaSmartLogin({
      auto: true,
      email: "user@company.com",
    });

    expect(result).toEqual(testToken);
    expect(mockedAutoLogin.acquireTokenViaAutoLogin).toHaveBeenCalledWith(
      expect.objectContaining({
        email: "user@company.com",
        headless: false,
      }),
    );
    expect(
      mockedInteractive.acquireTokenViaInteractiveLogin,
    ).not.toHaveBeenCalled();
  });

  it("should fall back to interactive when auto-login fails", async () => {
    mockedPlatform.canAttemptAutoLogin.mockReturnValue(true);
    mockedAutoLogin.acquireTokenViaAutoLogin.mockRejectedValue(
      new Error("FIDO2 not enrolled"),
    );
    mockedInteractive.acquireTokenViaInteractiveLogin.mockResolvedValue(
      testToken,
    );

    const result = await acquireTokenViaSmartLogin({
      auto: true,
      email: "user@company.com",
    });

    expect(result).toEqual(testToken);
    expect(mockedAutoLogin.acquireTokenViaAutoLogin).toHaveBeenCalled();
    expect(
      mockedInteractive.acquireTokenViaInteractiveLogin,
    ).toHaveBeenCalledWith(
      expect.objectContaining({ email: "user@company.com" }),
    );
  });

  it("should go straight to interactive when no email provided", async () => {
    mockedPlatform.canAttemptAutoLogin.mockReturnValue(true);
    mockedInteractive.acquireTokenViaInteractiveLogin.mockResolvedValue(
      testToken,
    );

    const result = await acquireTokenViaSmartLogin({});

    expect(result).toEqual(testToken);
    expect(mockedAutoLogin.acquireTokenViaAutoLogin).not.toHaveBeenCalled();
    expect(
      mockedInteractive.acquireTokenViaInteractiveLogin,
    ).toHaveBeenCalled();
  });

  it("should go straight to interactive when auto-login prerequisites not met", async () => {
    mockedPlatform.canAttemptAutoLogin.mockReturnValue(false);
    mockedInteractive.acquireTokenViaInteractiveLogin.mockResolvedValue(
      testToken,
    );

    const result = await acquireTokenViaSmartLogin({
      auto: true,
      email: "user@company.com",
    });

    expect(result).toEqual(testToken);
    expect(mockedAutoLogin.acquireTokenViaAutoLogin).not.toHaveBeenCalled();
    expect(
      mockedInteractive.acquireTokenViaInteractiveLogin,
    ).toHaveBeenCalled();
  });

  it("should pass region to both auto and interactive login", async () => {
    mockedPlatform.canAttemptAutoLogin.mockReturnValue(true);
    mockedAutoLogin.acquireTokenViaAutoLogin.mockRejectedValue(
      new Error("fail"),
    );
    mockedInteractive.acquireTokenViaInteractiveLogin.mockResolvedValue(
      testToken,
    );

    await acquireTokenViaSmartLogin({
      auto: true,
      email: "user@company.com",
      region: "emea",
    });

    expect(mockedAutoLogin.acquireTokenViaAutoLogin).toHaveBeenCalledWith(
      expect.objectContaining({ region: "emea" }),
    );
    expect(
      mockedInteractive.acquireTokenViaInteractiveLogin,
    ).toHaveBeenCalledWith(expect.objectContaining({ region: "emea" }));
  });

  it("should work with no options at all", async () => {
    mockedPlatform.canAttemptAutoLogin.mockReturnValue(false);
    mockedInteractive.acquireTokenViaInteractiveLogin.mockResolvedValue(
      testToken,
    );

    const result = await acquireTokenViaSmartLogin();

    expect(result).toEqual(testToken);
    expect(
      mockedInteractive.acquireTokenViaInteractiveLogin,
    ).toHaveBeenCalled();
  });
});
