/**
 * Unit tests for the emoji-map module (src/emoji-map.ts).
 *
 * Verifies that initializeEmojiMap fetches the emoji catalog from the
 * Teams CDN and that resolveReactionKey correctly maps user-friendly
 * shortcuts to Teams emoji IDs.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  resolveReactionKey,
  initializeEmojiMap,
  resetEmojiMap,
} from "../../src/emoji-map.js";

/** Minimal CDN catalog fixture used across all tests. */
const testCatalog = {
  categories: [
    {
      emoticons: [
        { id: "like", shortcuts: ["(like)"] },
        { id: "heart", shortcuts: ["(heart)"] },
        { id: "laugh", shortcuts: ["(laugh)"] },
        { id: "surprised", shortcuts: ["(surprised)"] },
        { id: "sad", shortcuts: ["(sad)", ":("] },
        { id: "angry", shortcuts: ["(angry)"] },
        { id: "1f40e_horse", shortcuts: ["(horse)"] },
        { id: "00a9_copyrightsign", shortcuts: ["(copyright)"] },
        { id: "1f300_cyclone", shortcuts: ["(cyclone)"] },
        { id: "1f30b_volcano", shortcuts: ["(volcano)"] },
        { id: "grinningfacewithsmilingeyes", shortcuts: [":D"] },
        { id: "smile", shortcuts: [":)"] },
        { id: "wink", shortcuts: [";)"] },
        { id: "cash", shortcuts: ["(dollar)", "(cash)"] },
        { id: "bartlett", shortcuts: ["(soccer)", "(bartlett)"] },
        { id: "bunnyhug", shortcuts: ["(rabbit)", "(bunnyhug)"] },
        { id: "hug", shortcuts: ["(teddybear)", "(hug)"] },
      ],
    },
  ],
};

const TEST_ECS_VERSION = "aabbccdd11223344aabbccdd11223344";
const TEST_ECS_RESPONSE = `{"emoticonAssetVersion":"${TEST_ECS_VERSION}","otherConfig":"value"}`;

/**
 * Stub fetch to succeed: ECS returns a version, CDN returns the catalog.
 */
function mockFetchSuccess(): void {
  vi.stubGlobal(
    "fetch",
    vi.fn().mockImplementation((url: string) => {
      if (url.includes("config.teams.microsoft.com")) {
        return Promise.resolve({ ok: true, text: () => Promise.resolve(TEST_ECS_RESPONSE) });
      }
      return Promise.resolve({ ok: true, json: () => Promise.resolve(testCatalog) });
    }),
  );
}

/**
 * Stub fetch to fail for all requests.
 */
function mockFetchFailure(): void {
  vi.stubGlobal(
    "fetch",
    vi.fn().mockResolvedValue({ ok: false, status: 404 }),
  );
}

/**
 * Stub fetch so ECS fails but the CDN succeeds on the first fallback version.
 */
function mockFetchEcsFailCdnSuccess(): void {
  vi.stubGlobal(
    "fetch",
    vi.fn().mockImplementation((url: string) => {
      if (url.includes("config.teams.microsoft.com")) {
        return Promise.resolve({ ok: false, status: 503 });
      }
      return Promise.resolve({ ok: true, json: () => Promise.resolve(testCatalog) });
    }),
  );
}

beforeEach(() => {
  resetEmojiMap();
  vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("initializeEmojiMap", () => {
  it("resolves the current version from ECS then fetches the CDN catalog", async () => {
    mockFetchSuccess();
    await initializeEmojiMap();
    const calls = vi.mocked(fetch).mock.calls.map((c) => c[0] as string);
    expect(calls.some((u) => u.includes("config.teams.microsoft.com"))).toBe(true);
    expect(calls.some((u) => u.includes("statics.teams.cdn.office.net"))).toBe(true);
    // CDN URL uses the ECS-resolved version hash
    expect(calls.some((u) => u.includes(TEST_ECS_VERSION))).toBe(true);
  });

  it("is idempotent — only fetches once", async () => {
    mockFetchSuccess();
    await initializeEmojiMap();
    await initializeEmojiMap();
    // ECS call + 1 CDN call (not doubled on second initializeEmojiMap)
    expect(fetch).toHaveBeenCalledTimes(2);
  });

  it("falls back to hardcoded versions when ECS returns a version not on the CDN", async () => {
    // ECS gives a version hash, but CDN 404s for it; hardcoded fallback succeeds
    vi.stubGlobal(
      "fetch",
      vi.fn().mockImplementation((url: string) => {
        if (url.includes("config.teams.microsoft.com")) {
          return Promise.resolve({ ok: true, text: () => Promise.resolve(TEST_ECS_RESPONSE) });
        }
        if (url.includes(TEST_ECS_VERSION)) {
          return Promise.resolve({ ok: false, status: 404 });
        }
        // Fallback CDN versions succeed
        return Promise.resolve({ ok: true, json: () => Promise.resolve(testCatalog) });
      }),
    );
    await initializeEmojiMap();
    expect(resolveReactionKey("horse")).toBe("1f40e_horse");
    const calls = vi.mocked(fetch).mock.calls.map((c) => c[0] as string);
    // ECS call, ECS-version CDN call (404), then a fallback version CDN call
    expect(calls.filter((u) => u.includes("statics.teams.cdn.office.net")).length).toBeGreaterThanOrEqual(2);
  });

  it("validates the ECS version format and falls back on malformed values", async () => {
    const malformedEcsResponse = '{"emoticonAssetVersion":"not-a-valid-hash"}';
    vi.stubGlobal(
      "fetch",
      vi.fn().mockImplementation((url: string) => {
        if (url.includes("config.teams.microsoft.com")) {
          return Promise.resolve({ ok: true, text: () => Promise.resolve(malformedEcsResponse) });
        }
        return Promise.resolve({ ok: true, json: () => Promise.resolve(testCatalog) });
      }),
    );
    await initializeEmojiMap();
    // Should still succeed via fallback versions
    expect(resolveReactionKey("horse")).toBe("1f40e_horse");
    const warnings = vi.mocked(console.warn).mock.calls.map((c) => c[0] as string);
    expect(warnings.some((w) => w.includes("Unexpected emoticonAssetVersion format"))).toBe(true);
  });

  it("falls back to hardcoded versions when ECS is unreachable", async () => {
    mockFetchEcsFailCdnSuccess();
    await initializeEmojiMap();
    expect(resolveReactionKey("horse")).toBe("1f40e_horse");
    const calls = vi.mocked(fetch).mock.calls.map((c) => c[0] as string);
    // Should have tried ECS then a fallback CDN version
    expect(calls.some((u) => u.includes("config.teams.microsoft.com"))).toBe(true);
    expect(calls.some((u) => u.includes("statics.teams.cdn.office.net"))).toBe(true);
  });

  it("does not throw when all fetches fail", async () => {
    mockFetchFailure();
    await expect(initializeEmojiMap()).resolves.toBeUndefined();
  });

  it("logs warnings when all versions fail", async () => {
    mockFetchFailure();
    await initializeEmojiMap();
    const warnings = vi
      .mocked(console.warn)
      .mock.calls.map((call) => call[0] as string);
    expect(
      warnings.some((w) => w.includes("Failed to fetch emoji catalog")),
    ).toBe(true);
  });
});

describe("resolveReactionKey (map loaded)", () => {
  beforeEach(async () => {
    mockFetchSuccess();
    await initializeEmojiMap();
  });

  describe("standard reactions (shortcut === id)", () => {
    it("passes through standard reaction keys unchanged", () => {
      expect(resolveReactionKey("like")).toBe("like");
      expect(resolveReactionKey("heart")).toBe("heart");
      expect(resolveReactionKey("laugh")).toBe("laugh");
      expect(resolveReactionKey("surprised")).toBe("surprised");
      expect(resolveReactionKey("angry")).toBe("angry");
      expect(resolveReactionKey("sad")).toBe("sad");
    });

    it("lowercases standard reaction keys", () => {
      expect(resolveReactionKey("Like")).toBe("like");
      expect(resolveReactionKey("HEART")).toBe("heart");
      expect(resolveReactionKey("Laugh")).toBe("laugh");
    });
  });

  describe("non-standard emojis (shortcut !== id)", () => {
    it("resolves shortcut to emoji ID", () => {
      expect(resolveReactionKey("horse")).toBe("1f40e_horse");
      expect(resolveReactionKey("copyright")).toBe("00a9_copyrightsign");
      expect(resolveReactionKey("cyclone")).toBe("1f300_cyclone");
      expect(resolveReactionKey("volcano")).toBe("1f30b_volcano");
    });

    it("resolves case-insensitively", () => {
      expect(resolveReactionKey("Horse")).toBe("1f40e_horse");
      expect(resolveReactionKey("HORSE")).toBe("1f40e_horse");
      expect(resolveReactionKey("Volcano")).toBe("1f30b_volcano");
    });
  });

  describe("emoji IDs passed directly", () => {
    it("passes through emoji IDs unchanged", () => {
      expect(resolveReactionKey("1f40e_horse")).toBe("1f40e_horse");
      expect(resolveReactionKey("00a9_copyrightsign")).toBe(
        "00a9_copyrightsign",
      );
    });
  });

  describe("emoticon shortcuts", () => {
    it("resolves text emoticons to emoji IDs", () => {
      expect(resolveReactionKey(":D")).toBe("grinningfacewithsmilingeyes");
      expect(resolveReactionKey(":)")).toBe("smile");
      expect(resolveReactionKey(":(")).toBe("sad");
      expect(resolveReactionKey(";)")).toBe("wink");
    });
  });

  describe("unknown keys", () => {
    it("passes through unknown keys lowercased", () => {
      expect(resolveReactionKey("unknown_emoji_xyz")).toBe(
        "unknown_emoji_xyz",
      );
      expect(resolveReactionKey("CustomReaction")).toBe("customreaction");
    });
  });

  describe("aliases with non-hex IDs", () => {
    it("resolves shortcuts that map to non-hex IDs", () => {
      expect(resolveReactionKey("dollar")).toBe("cash");
      expect(resolveReactionKey("soccer")).toBe("bartlett");
      expect(resolveReactionKey("rabbit")).toBe("bunnyhug");
      expect(resolveReactionKey("teddybear")).toBe("hug");
    });
  });
});

describe("resolveReactionKey (map not loaded)", () => {
  it("falls back to lowercased input", () => {
    expect(resolveReactionKey("Horse")).toBe("horse");
    expect(resolveReactionKey("LIKE")).toBe("like");
    expect(resolveReactionKey("unknown")).toBe("unknown");
  });
});
