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

function mockFetchSuccess(): void {
  vi.stubGlobal(
    "fetch",
    vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve(testCatalog),
    }),
  );
}

function mockFetchFailure(): void {
  vi.stubGlobal(
    "fetch",
    vi.fn().mockResolvedValue({ ok: false, status: 404 }),
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
  it("fetches the emoji catalog from the CDN", async () => {
    mockFetchSuccess();
    await initializeEmojiMap();
    expect(fetch).toHaveBeenCalledWith(
      expect.stringContaining(
        "statics.teams.cdn.office.net/evergreen-assets/personal-expressions",
      ),
    );
  });

  it("is idempotent — only fetches once", async () => {
    mockFetchSuccess();
    await initializeEmojiMap();
    await initializeEmojiMap();
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it("does not throw when fetch fails", async () => {
    mockFetchFailure();
    await expect(initializeEmojiMap()).resolves.toBeUndefined();
  });

  it("tries multiple version hashes on failure", async () => {
    vi.stubGlobal(
      "fetch",
      vi
        .fn()
        .mockResolvedValueOnce({ ok: false, status: 404 })
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(testCatalog),
        }),
    );
    await initializeEmojiMap();
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(resolveReactionKey("horse")).toBe("1f40e_horse");
  });

  it("logs warnings when all versions fail", async () => {
    mockFetchFailure();
    await initializeEmojiMap();
    const warnings = vi
      .mocked(console.warn)
      .mock.calls.map((call) => call[0] as string);
    expect(warnings.some((warning) => warning.includes("CDN returned 404"))).toBe(true);
    expect(
      warnings.some((warning) => warning.includes("Failed to fetch emoji catalog")),
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
