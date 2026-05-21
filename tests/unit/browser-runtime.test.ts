/**
 * Unit tests for browser runtime helpers (src/browser-runtime.ts).
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { mkdirSync, symlinkSync, lstatSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import type { Browser, BrowserContext } from "playwright";
import {
  getInteractiveBrowserChannels,
  getDefaultBrowserProfileDir,
  isMissingPlaywrightBrowserError,
  isProfileLockError,
  cleanStaleSingletonLock,
  launchInteractiveBrowser,
  launchInteractiveBrowserContext,
} from "../../src/browser-runtime.js";

function symlinkExists(path: string): boolean {
  try {
    lstatSync(path);
    return true;
  } catch {
    return false;
  }
}

function createBrowserStub(): Browser {
  return {} as Browser;
}

function createContextStub(): BrowserContext {
  return {} as BrowserContext;
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

describe("getDefaultBrowserProfileDir", () => {
  it("should return a path under the home directory", () => {
    const dir = getDefaultBrowserProfileDir();
    expect(dir).toContain(".teams-api");
    expect(dir).toContain("browser-profile");
  });
});

describe("launchInteractiveBrowserContext", () => {
  const userDataDir = "/tmp/test-profile";

  it("should use the first available installed browser channel with persistent context", async () => {
    const context = createContextStub();
    const chromium = {
      launch: vi.fn(),
      launchPersistentContext: vi.fn().mockResolvedValue(context),
    };

    const result = await launchInteractiveBrowserContext(
      chromium,
      vi.fn(),
      userDataDir,
      { platform: "win32", installBundledChromium: vi.fn() },
    );

    expect(result).toBe(context);
    expect(chromium.launchPersistentContext).toHaveBeenCalledTimes(1);
    expect(chromium.launchPersistentContext).toHaveBeenCalledWith(userDataDir, {
      headless: false,
      channel: "msedge",
    });
  });

  it("should fall back to the next channel with persistent context", async () => {
    const context = createContextStub();
    const chromium = {
      launch: vi.fn(),
      launchPersistentContext: vi
        .fn()
        .mockRejectedValueOnce(new Error("Edge not installed"))
        .mockResolvedValueOnce(context),
    };

    const result = await launchInteractiveBrowserContext(
      chromium,
      vi.fn(),
      userDataDir,
      { platform: "win32", installBundledChromium: vi.fn() },
    );

    expect(result).toBe(context);
    expect(chromium.launchPersistentContext).toHaveBeenNthCalledWith(
      1,
      userDataDir,
      { headless: false, channel: "msedge" },
    );
    expect(chromium.launchPersistentContext).toHaveBeenNthCalledWith(
      2,
      userDataDir,
      { headless: false, channel: "chrome" },
    );
  });

  it("should fall back to bundled Chromium with persistent context", async () => {
    const context = createContextStub();
    const chromium = {
      launch: vi.fn(),
      launchPersistentContext: vi
        .fn()
        .mockRejectedValueOnce(new Error("Edge not installed"))
        .mockRejectedValueOnce(new Error("Chrome not installed"))
        .mockResolvedValueOnce(context),
    };

    const result = await launchInteractiveBrowserContext(
      chromium,
      vi.fn(),
      userDataDir,
      { platform: "win32", installBundledChromium: vi.fn() },
    );

    expect(result).toBe(context);
    expect(chromium.launchPersistentContext).toHaveBeenNthCalledWith(
      3,
      userDataDir,
      { headless: false },
    );
  });

  it("should install bundled Chromium and retry with persistent context", async () => {
    const context = createContextStub();
    const installBundledChromium = vi.fn();
    const chromium = {
      launch: vi.fn(),
      launchPersistentContext: vi
        .fn()
        .mockRejectedValueOnce(new Error("Edge not installed"))
        .mockRejectedValueOnce(new Error("Chrome not installed"))
        .mockRejectedValueOnce(
          new Error("Executable doesn't exist at C:\\Users\\me\\browser.exe"),
        )
        .mockResolvedValueOnce(context),
    };

    const result = await launchInteractiveBrowserContext(
      chromium,
      vi.fn(),
      userDataDir,
      { platform: "win32", installBundledChromium },
    );

    expect(result).toBe(context);
    expect(installBundledChromium).toHaveBeenCalledTimes(1);
    expect(chromium.launchPersistentContext).toHaveBeenNthCalledWith(
      4,
      userDataDir,
      { headless: false },
    );
  });

  it("should retry on profile lock error after cleaning stale lock", async () => {
    const context = createContextStub();
    const chromium = {
      launch: vi.fn(),
      launchPersistentContext: vi
        .fn()
        .mockRejectedValueOnce(new Error("Browser exit code 21"))
        .mockResolvedValueOnce(context),
    };

    const result = await launchInteractiveBrowserContext(
      chromium,
      vi.fn(),
      userDataDir,
      { platform: "darwin", installBundledChromium: vi.fn() },
    );

    expect(result).toBe(context);
    // First attempt fails with lock error, second attempt succeeds
    expect(chromium.launchPersistentContext).toHaveBeenCalledTimes(2);
  });

  it("should throw after exhausting retries on persistent lock error", async () => {
    const chromium = {
      launch: vi.fn(),
      launchPersistentContext: vi
        .fn()
        .mockRejectedValue(new Error("Browser exit code 21")),
    };

    await expect(
      launchInteractiveBrowserContext(chromium, vi.fn(), userDataDir, {
        platform: "darwin",
        installBundledChromium: vi.fn(),
      }),
    ).rejects.toThrow("is locked and could not be cleaned");
  });
});

describe("isProfileLockError", () => {
  it("should detect exit code 21 errors", () => {
    expect(isProfileLockError(new Error("Browser exit code 21"))).toBe(true);
  });

  it("should detect SingletonLock errors", () => {
    expect(
      isProfileLockError(new Error("Failed: SingletonLock file exists")),
    ).toBe(true);
  });

  it("should detect process singleton errors", () => {
    expect(
      isProfileLockError(
        new Error("Failed to create a Process Singleton for your profile"),
      ),
    ).toBe(true);
  });

  it("should detect profile in use errors", () => {
    expect(
      isProfileLockError(
        new Error("profile directory is already in use by another process"),
      ),
    ).toBe(true);
  });

  it("should not match unrelated errors", () => {
    expect(isProfileLockError(new Error("Network timeout"))).toBe(false);
  });

  it("should not match non-Error values", () => {
    expect(isProfileLockError("exit code 21")).toBe(false);
    expect(isProfileLockError(null)).toBe(false);
  });
});

describe("cleanStaleSingletonLock", () => {
  let testProfileDir: string;

  beforeEach(() => {
    testProfileDir = join(
      tmpdir(),
      `teams-api-test-profile-${Date.now()}-${Math.random().toString(36).slice(2)}`,
    );
    mkdirSync(testProfileDir, { recursive: true });
  });

  afterEach(() => {
    rmSync(testProfileDir, { recursive: true, force: true });
  });

  it("should return false when no SingletonLock exists", () => {
    const result = cleanStaleSingletonLock(testProfileDir, vi.fn());
    expect(result).toBe(false);
  });

  it("should remove lock with dead PID", () => {
    const lockPath = join(testProfileDir, "SingletonLock");
    // Use a PID that almost certainly doesn't exist
    symlinkSync("hostname-999999999", lockPath);

    const result = cleanStaleSingletonLock(testProfileDir, vi.fn());
    expect(result).toBe(true);
    expect(symlinkExists(lockPath)).toBe(false);
  });

  it("should not remove lock held by a live process", () => {
    const lockPath = join(testProfileDir, "SingletonLock");
    // Use our own PID (which is definitely running)
    symlinkSync(`hostname-${process.pid}`, lockPath);

    const result = cleanStaleSingletonLock(testProfileDir, vi.fn());
    expect(result).toBe(false);
    expect(symlinkExists(lockPath)).toBe(true);
  });
});
