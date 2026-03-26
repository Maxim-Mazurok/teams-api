/**
 * Unit tests for browser runtime helpers (src/browser-runtime.ts).
 */

import { describe, it, expect, vi } from "vitest";
import type { Browser } from "playwright";
import {
  getInteractiveBrowserChannels,
  isMissingPlaywrightBrowserError,
  launchInteractiveBrowser,
} from "../../src/browser-runtime.js";

function createBrowserStub(): Browser {
  return {} as Browser;
}

describe("getInteractiveBrowserChannels", () => {
  it("should prefer Edge then Chrome on Windows", () => {
    expect(getInteractiveBrowserChannels("win32")).toEqual([
      "msedge",
      "chrome",
    ]);
  });

  it("should prefer Chrome on macOS", () => {
    expect(getInteractiveBrowserChannels("darwin")).toEqual(["chrome"]);
  });

  it("should prefer Chrome then Edge on Linux", () => {
    expect(getInteractiveBrowserChannels("linux")).toEqual([
      "chrome",
      "msedge",
    ]);
  });
});

describe("isMissingPlaywrightBrowserError", () => {
  it("should detect a missing bundled-browser executable", () => {
    expect(
      isMissingPlaywrightBrowserError(
        new Error("Executable doesn't exist at C:\\Users\\me\\browser.exe"),
      ),
    ).toBe(true);
  });

  it("should ignore unrelated launch errors", () => {
    expect(isMissingPlaywrightBrowserError(new Error("Launch crashed"))).toBe(
      false,
    );
  });
});

describe("launchInteractiveBrowser", () => {
  it("should use the first available installed browser channel", async () => {
    const browser = createBrowserStub();
    const chromium = {
      launch: vi.fn().mockResolvedValue(browser),
    };

    const result = await launchInteractiveBrowser(chromium, vi.fn(), {
      platform: "win32",
      installBundledChromium: vi.fn(),
    });

    expect(result).toBe(browser);
    expect(chromium.launch).toHaveBeenCalledTimes(1);
    expect(chromium.launch).toHaveBeenCalledWith({
      headless: false,
      channel: "msedge",
    });
  });

  it("should fall back to the next installed browser channel", async () => {
    const browser = createBrowserStub();
    const chromium = {
      launch: vi
        .fn()
        .mockRejectedValueOnce(new Error("Edge not installed"))
        .mockResolvedValueOnce(browser),
    };

    const result = await launchInteractiveBrowser(chromium, vi.fn(), {
      platform: "win32",
      installBundledChromium: vi.fn(),
    });

    expect(result).toBe(browser);
    expect(chromium.launch).toHaveBeenNthCalledWith(1, {
      headless: false,
      channel: "msedge",
    });
    expect(chromium.launch).toHaveBeenNthCalledWith(2, {
      headless: false,
      channel: "chrome",
    });
  });

  it("should fall back to bundled Chromium when installed channels fail", async () => {
    const browser = createBrowserStub();
    const chromium = {
      launch: vi
        .fn()
        .mockRejectedValueOnce(new Error("Edge not installed"))
        .mockRejectedValueOnce(new Error("Chrome not installed"))
        .mockResolvedValueOnce(browser),
    };

    const result = await launchInteractiveBrowser(chromium, vi.fn(), {
      platform: "win32",
      installBundledChromium: vi.fn(),
    });

    expect(result).toBe(browser);
    expect(chromium.launch).toHaveBeenNthCalledWith(3, {
      headless: false,
    });
  });

  it("should install bundled Chromium and retry when it is missing", async () => {
    const browser = createBrowserStub();
    const installBundledChromium = vi.fn();
    const chromium = {
      launch: vi
        .fn()
        .mockRejectedValueOnce(new Error("Edge not installed"))
        .mockRejectedValueOnce(new Error("Chrome not installed"))
        .mockRejectedValueOnce(
          new Error("Executable doesn't exist at C:\\Users\\me\\browser.exe"),
        )
        .mockResolvedValueOnce(browser),
    };

    const result = await launchInteractiveBrowser(chromium, vi.fn(), {
      platform: "win32",
      installBundledChromium,
    });

    expect(result).toBe(browser);
    expect(installBundledChromium).toHaveBeenCalledTimes(1);
    expect(chromium.launch).toHaveBeenNthCalledWith(4, {
      headless: false,
    });
  });

  it("should rethrow non-installable bundled Chromium launch errors", async () => {
    const chromium = {
      launch: vi
        .fn()
        .mockRejectedValueOnce(new Error("Chrome not installed"))
        .mockRejectedValueOnce(new Error("Edge not installed"))
        .mockRejectedValueOnce(new Error("Browser sandbox failure")),
    };

    await expect(
      launchInteractiveBrowser(chromium, vi.fn(), {
        platform: "linux",
        installBundledChromium: vi.fn(),
      }),
    ).rejects.toThrow("Browser sandbox failure");
  });
});
