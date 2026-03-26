import { describe, it, expect } from "vitest";
import {
  formatOutput,
  cleanContent,
  extractQuote,
  buildSenderLookup,
  formatTimestamp,
  groupBySpeaker,
} from "../../src/actions/formatters.js";
import type {
  ActionDefinition,
  OutputFormat,
} from "../../src/actions/formatters.js";
import type { Message, TranscriptEntry } from "../../src/types.js";

// ── Helpers ──────────────────────────────────────────────────────────

function makeMessage(overrides: Partial<Message> & { id: string }): Message {
  return {
    messageType: "RichText/Html",
    senderMri: "8:orgid:test",
    senderDisplayName: "Test User",
    content: "",
    originalArrivalTime: "2025-01-01T00:00:00.000Z",
    composeTime: "2025-01-01T00:00:00.000Z",
    editTime: null,
    subject: null,
    isDeleted: false,
    reactions: [],
    followers: [],
    mentions: [],
    quotedMessageId: null,
    images: [],
    files: [],
    ...overrides,
  };
}

function makeMockAction(
  overrides?: Partial<ActionDefinition>,
): ActionDefinition {
  return {
    name: "test-action",
    title: "Test Action",
    description: "A mock action for testing",
    parameters: [],
    execute: async () => ({}),
    formatConcise: (result) => `concise:${JSON.stringify(result)}`,
    ...overrides,
  };
}

// ── cleanContent ─────────────────────────────────────────────────────

describe("cleanContent", () => {
  it("strips HTML tags", () => {
    expect(cleanContent("<p>Hello <b>world</b></p>")).toBe("Hello world");
  });

  it("decodes &amp;", () => {
    expect(cleanContent("Tom &amp; Jerry")).toBe("Tom & Jerry");
  });

  it("decodes &lt; and &gt;", () => {
    expect(cleanContent("a &lt; b &gt; c")).toBe("a < b > c");
  });

  it("decodes &quot;", () => {
    expect(cleanContent("She said &quot;hi&quot;")).toBe('She said "hi"');
  });

  it("decodes &nbsp; to space", () => {
    expect(cleanContent("hello&nbsp;world")).toBe("hello world");
  });

  it("trims surrounding whitespace", () => {
    expect(cleanContent("  hello  ")).toBe("hello");
  });

  it("handles combined HTML and entities", () => {
    expect(cleanContent("<div>&lt;tag&gt; &amp; more</div>")).toBe(
      "<tag> & more",
    );
  });
});

// ── extractQuote ─────────────────────────────────────────────────────

describe("extractQuote", () => {
  it("extracts text from <blockquote> tags", () => {
    const content = "<blockquote>quoted text</blockquote>reply text";
    const result = extractQuote(content);
    expect(result.quote).toBe("quoted text");
    expect(result.body).toBe("reply text");
  });

  it("extracts text from <quote> tags", () => {
    const content = "<quote>quoted stuff</quote> rest of message";
    const result = extractQuote(content);
    expect(result.quote).toBe("quoted stuff");
    expect(result.body).toBe("rest of message");
  });

  it("returns null quote when no blockquote present", () => {
    const result = extractQuote("just a normal message");
    expect(result.quote).toBeNull();
    expect(result.body).toBe("just a normal message");
  });

  it("handles nested HTML inside blockquote", () => {
    const content =
      "<blockquote><b>bold</b> and <i>italic</i></blockquote>after";
    const result = extractQuote(content);
    expect(result.quote).toBe("bold and italic");
    expect(result.body).toBe("after");
  });

  it("returns null quote when blockquote content is empty after cleaning", () => {
    const content = "<blockquote>   </blockquote>body here";
    const result = extractQuote(content);
    // cleanContent("   ") => "" which is falsy, so quote = null
    expect(result.quote).toBeNull();
    expect(result.body).toBe("body here");
  });

  it("prefers blockquote over quote tag", () => {
    const content =
      "<blockquote>from blockquote</blockquote><quote>from quote</quote>body";
    const result = extractQuote(content);
    expect(result.quote).toBe("from blockquote");
  });
});

// ── buildSenderLookup ────────────────────────────────────────────────

describe("buildSenderLookup", () => {
  it("maps message IDs to sender display names", () => {
    const messages: Message[] = [
      makeMessage({ id: "1", senderDisplayName: "Alice" }),
      makeMessage({ id: "2", senderDisplayName: "Bob" }),
    ];
    const lookup = buildSenderLookup(messages);
    expect(lookup.get("1")).toBe("Alice");
    expect(lookup.get("2")).toBe("Bob");
  });

  it('uses "(system)" when senderDisplayName is empty', () => {
    const messages: Message[] = [
      makeMessage({ id: "3", senderDisplayName: "" }),
    ];
    const lookup = buildSenderLookup(messages);
    expect(lookup.get("3")).toBe("(system)");
  });

  it("handles empty array", () => {
    const lookup = buildSenderLookup([]);
    expect(lookup.size).toBe(0);
  });
});

// ── formatTimestamp ──────────────────────────────────────────────────

describe("formatTimestamp", () => {
  it("strips milliseconds", () => {
    expect(formatTimestamp("00:05:30.500")).toBe("05:30");
  });

  it('removes leading "00:" hours when hours are zero', () => {
    expect(formatTimestamp("00:12:45.000")).toBe("12:45");
  });

  it("keeps non-zero hours", () => {
    expect(formatTimestamp("01:23:45.789")).toBe("01:23:45");
  });

  it("handles timestamp without milliseconds", () => {
    expect(formatTimestamp("00:10:20")).toBe("10:20");
  });

  it("handles hours with no milliseconds", () => {
    expect(formatTimestamp("02:05:09")).toBe("02:05:09");
  });
});

// ── groupBySpeaker ──────────────────────────────────────────────────

describe("groupBySpeaker", () => {
  it("groups consecutive entries by the same speaker", () => {
    const entries: TranscriptEntry[] = [
      {
        speaker: "Alice",
        startTime: "00:00:01.000",
        endTime: "00:00:05.000",
        text: "Hello",
      },
      {
        speaker: "Alice",
        startTime: "00:00:06.000",
        endTime: "00:00:10.000",
        text: "How are you?",
      },
      {
        speaker: "Bob",
        startTime: "00:00:11.000",
        endTime: "00:00:15.000",
        text: "I'm good",
      },
    ];
    const groups = groupBySpeaker(entries);

    expect(groups).toHaveLength(2);
    expect(groups[0]).toEqual({
      speaker: "Alice",
      startTime: "00:00:01.000",
      segments: ["Hello", "How are you?"],
    });
    expect(groups[1]).toEqual({
      speaker: "Bob",
      startTime: "00:00:11.000",
      segments: ["I'm good"],
    });
  });

  it("starts a new group on speaker change", () => {
    const entries: TranscriptEntry[] = [
      {
        speaker: "Alice",
        startTime: "00:00:01.000",
        endTime: "00:00:03.000",
        text: "One",
      },
      {
        speaker: "Bob",
        startTime: "00:00:04.000",
        endTime: "00:00:06.000",
        text: "Two",
      },
      {
        speaker: "Alice",
        startTime: "00:00:07.000",
        endTime: "00:00:09.000",
        text: "Three",
      },
    ];
    const groups = groupBySpeaker(entries);

    expect(groups).toHaveLength(3);
    expect(groups[0].speaker).toBe("Alice");
    expect(groups[1].speaker).toBe("Bob");
    expect(groups[2].speaker).toBe("Alice");
  });

  it("handles a single entry", () => {
    const entries: TranscriptEntry[] = [
      {
        speaker: "Solo",
        startTime: "00:00:01.000",
        endTime: "00:00:05.000",
        text: "Just me",
      },
    ];
    const groups = groupBySpeaker(entries);

    expect(groups).toHaveLength(1);
    expect(groups[0]).toEqual({
      speaker: "Solo",
      startTime: "00:00:01.000",
      segments: ["Just me"],
    });
  });

  it("returns empty array for empty input", () => {
    expect(groupBySpeaker([])).toEqual([]);
  });
});

// ── formatOutput ─────────────────────────────────────────────────────

describe("formatOutput", () => {
  const action = makeMockAction();
  const testData = { key: "value" };

  it("defaults to concise format", () => {
    const result = formatOutput(action, testData);
    expect(result).toBe(`concise:${JSON.stringify(testData)}`);
  });

  it("dispatches to concise formatter", () => {
    const result = formatOutput(action, testData, "concise");
    expect(result).toBe(`concise:${JSON.stringify(testData)}`);
  });

  it("dispatches to detailed formatter (JSON)", () => {
    const result = formatOutput(action, testData, "detailed");
    expect(result).toBe(JSON.stringify(testData, null, 2));
  });
});
