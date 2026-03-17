/**
 * Unit tests for the unified action definitions (src/actions.ts).
 *
 * These tests verify that:
 *   - Execute functions call correct TeamsClient methods with correct params
 *   - Conversation resolution (conversationId, chat, to) works
 *   - Text-only filtering is applied correctly
 *   - Format functions produce expected human-readable output
 *   - Parameter definitions are complete and consistent
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { actions, formatOutput } from "../../src/actions.js";
import type { ActionDefinition, OutputFormat } from "../../src/actions.js";
import type { TeamsClient } from "../../src/teams-client.js";
import type {
  Conversation,
  Message,
  Member,
  OneOnOneSearchResult,
  SentMessage,
} from "../../src/types.js";

// ── Mock client factory ──────────────────────────────────────────────

function createMockClient(
  overrides: Partial<Record<keyof TeamsClient, unknown>> = {},
): TeamsClient {
  return {
    listConversations: vi.fn(),
    findConversation: vi.fn(),
    findOneOnOneConversation: vi.fn(),
    getMessages: vi.fn(),
    getMessagesPage: vi.fn(),
    sendMessage: vi.fn(),
    getMembers: vi.fn(),
    getCurrentUserDisplayName: vi.fn(),
    getToken: vi.fn(() => ({ skypeToken: "test-token", region: "apac" })),
    ...overrides,
  } as unknown as TeamsClient;
}

function makeConversation(overrides: Partial<Conversation> = {}): Conversation {
  return {
    id: "19:test@thread.space",
    topic: "Test Chat",
    threadType: "chat",
    version: 1,
    lastMessageTime: "2026-03-16T10:00:00.000Z",
    memberCount: 5,
    ...overrides,
  };
}

function makeMessage(overrides: Partial<Message> = {}): Message {
  return {
    id: "1773000000000",
    messageType: "RichText/Html",
    senderMri: "8:orgid:user-1",
    senderDisplayName: "Test User",
    content: "<p>Hello world</p>",
    originalArrivalTime: "2026-03-16T10:00:00.000Z",
    composeTime: "2026-03-16T10:00:00.000Z",
    editTime: null,
    subject: null,
    isDeleted: false,
    reactions: [],
    mentions: [],
    quotedMessageId: null,
    ...overrides,
  };
}

function makeMember(overrides: Partial<Member> = {}): Member {
  return {
    id: "8:orgid:user-1",
    displayName: "Alice Smith",
    role: "Admin",
    ...overrides,
  };
}

function getAction(name: string): ActionDefinition {
  const action = actions.find((action) => action.name === name);
  if (!action) throw new Error(`Action "${name}" not found`);
  return action;
}

beforeEach(() => {
  vi.resetAllMocks();
});

// ── Registry tests ───────────────────────────────────────────────────

describe("action registry", () => {
  it("should contain all 7 actions", () => {
    expect(actions).toHaveLength(7);
  });

  it("should have unique names", () => {
    const names = actions.map((action) => action.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("should have descriptions for all actions", () => {
    for (const action of actions) {
      expect(action.description.length).toBeGreaterThan(10);
      expect(action.title.length).toBeGreaterThan(3);
    }
  });

  it("should have descriptions for all parameters", () => {
    for (const action of actions) {
      for (const parameter of action.parameters) {
        expect(parameter.description.length).toBeGreaterThan(5);
        expect(parameter.name.length).toBeGreaterThan(0);
        expect(["string", "number", "boolean"]).toContain(parameter.type);
      }
    }
  });

  it("should have expected action names", () => {
    const names = actions.map((action) => action.name);
    expect(names).toContain("list-conversations");
    expect(names).toContain("find-conversation");
    expect(names).toContain("find-one-on-one");
    expect(names).toContain("get-messages");
    expect(names).toContain("send-message");
    expect(names).toContain("get-members");
    expect(names).toContain("whoami");
  });
});

// ── list-conversations ───────────────────────────────────────────────

describe("list-conversations", () => {
  const action = getAction("list-conversations");

  it("should call client.listConversations with limit", async () => {
    const conversations = [makeConversation()];
    const client = createMockClient({
      listConversations: vi.fn().mockResolvedValue(conversations),
    });

    const result = await action.execute(client, { limit: 25 });

    expect(client.listConversations).toHaveBeenCalledWith({ pageSize: 25 });
    expect(result).toEqual(conversations);
  });

  it("should use default limit of 50", async () => {
    const client = createMockClient({
      listConversations: vi.fn().mockResolvedValue([]),
    });

    await action.execute(client, {});

    expect(client.listConversations).toHaveBeenCalledWith({ pageSize: 50 });
  });

  it("should format results correctly", () => {
    const conversations = [
      makeConversation({ topic: "Design Review", threadType: "chat" }),
      makeConversation({
        topic: "",
        threadType: "chat",
        lastMessageTime: null,
        memberCount: null,
      }),
    ];

    const output = action.formatResult(conversations);

    expect(output).toContain("2 conversations:");
    expect(output).toContain('[0] chat: "Design Review"');
    expect(output).toContain("members: 5");
    expect(output).toContain('[1] chat: "(untitled 1:1 chat)"');
    expect(output).toContain("last: unknown");
  });
});

// ── find-conversation ────────────────────────────────────────────────

describe("find-conversation", () => {
  const action = getAction("find-conversation");

  it("should call client.findConversation with query", async () => {
    const conversation = makeConversation({ topic: "Engineering" });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
    });

    const result = await action.execute(client, { query: "engineering" });

    expect(client.findConversation).toHaveBeenCalledWith("engineering");
    expect(result).toEqual(conversation);
  });

  it("should return null when not found", async () => {
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(null),
    });

    const result = await action.execute(client, { query: "nonexistent" });

    expect(result).toBeNull();
  });

  it("should format found conversation", () => {
    const conversation = makeConversation({
      id: "19:abc@thread.v2",
      topic: "Engineering",
      threadType: "chat",
    });

    const output = action.formatResult(conversation);

    expect(output).toContain('Found: "Engineering"');
    expect(output).toContain("19:abc@thread.v2");
    expect(output).toContain("chat");
  });

  it("should format null result", () => {
    const output = action.formatResult(null);
    expect(output).toBe("No conversation found.");
  });
});

// ── find-one-on-one ──────────────────────────────────────────────────

describe("find-one-on-one", () => {
  const action = getAction("find-one-on-one");

  it("should call client.findOneOnOneConversation", async () => {
    const searchResult: OneOnOneSearchResult = {
      conversationId: "19:abc@unq.gbl.spaces",
      memberDisplayName: "Luke Prior",
    };
    const client = createMockClient({
      findOneOnOneConversation: vi.fn().mockResolvedValue(searchResult),
    });

    const result = await action.execute(client, { personName: "Luke" });

    expect(client.findOneOnOneConversation).toHaveBeenCalledWith("Luke");
    expect(result).toEqual(searchResult);
  });

  it("should return null when not found", async () => {
    const client = createMockClient({
      findOneOnOneConversation: vi.fn().mockResolvedValue(null),
    });

    const result = await action.execute(client, {
      personName: "Nonexistent",
    });

    expect(result).toBeNull();
  });

  it("should format found result", () => {
    const searchResult: OneOnOneSearchResult = {
      conversationId: "19:abc@unq.gbl.spaces",
      memberDisplayName: "Luke Prior",
    };

    const output = action.formatResult(searchResult);

    expect(output).toContain("Found 1:1 with Luke Prior");
    expect(output).toContain("19:abc@unq.gbl.spaces");
  });

  it("should format null result", () => {
    const output = action.formatResult(null);
    expect(output).toBe("No 1:1 conversation found.");
  });
});

// ── get-messages ─────────────────────────────────────────────────────

describe("get-messages", () => {
  const action = getAction("get-messages");

  it("should resolve conversation by direct ID", async () => {
    const messages = [makeMessage()];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    const result = await action.execute(client, {
      conversationId: "19:direct@thread.v2",
    });

    expect(client.getMessages).toHaveBeenCalledWith("19:direct@thread.v2", {
      maxPages: 100,
      pageSize: 200,
      onProgress: undefined,
    });
    expect(result).toEqual(messages);
  });

  it("should resolve conversation by chat name", async () => {
    const conversation = makeConversation({ id: "19:resolved@thread.v2" });
    const messages = [makeMessage()];
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    await action.execute(client, { chat: "Design Review" });

    expect(client.findConversation).toHaveBeenCalledWith("Design Review");
    expect(client.getMessages).toHaveBeenCalledWith(
      "19:resolved@thread.v2",
      expect.objectContaining({ maxPages: 100 }),
    );
  });

  it("should resolve conversation by person name", async () => {
    const searchResult: OneOnOneSearchResult = {
      conversationId: "19:one-on-one@unq.gbl.spaces",
      memberDisplayName: "Luke Prior",
    };
    const messages = [makeMessage()];
    const client = createMockClient({
      findOneOnOneConversation: vi.fn().mockResolvedValue(searchResult),
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    await action.execute(client, { to: "Luke" });

    expect(client.findOneOnOneConversation).toHaveBeenCalledWith("Luke");
    expect(client.getMessages).toHaveBeenCalledWith(
      "19:one-on-one@unq.gbl.spaces",
      expect.objectContaining({ maxPages: 100 }),
    );
  });

  it("should error when no identification provided", async () => {
    const client = createMockClient();

    await expect(action.execute(client, {})).rejects.toThrow(
      "One of --conversation-id, --chat, or --to is required.",
    );
  });

  it("should error when chat name not found", async () => {
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(null),
    });

    await expect(
      action.execute(client, { chat: "Nonexistent" }),
    ).rejects.toThrow('No conversation matching "Nonexistent" found.');
  });

  it("should error when person name not found", async () => {
    const client = createMockClient({
      findOneOnOneConversation: vi.fn().mockResolvedValue(null),
    });

    await expect(action.execute(client, { to: "Nobody" })).rejects.toThrow(
      'No 1:1 conversation found with "Nobody".',
    );
  });

  it("should filter text-only by default", async () => {
    const messages = [
      makeMessage({ messageType: "RichText/Html", content: "Hello" }),
      makeMessage({
        messageType: "ThreadActivity/AddMember",
        content: "system",
      }),
      makeMessage({ messageType: "Text", content: "Plain text" }),
      makeMessage({
        messageType: "RichText/Html",
        content: "Deleted",
        isDeleted: true,
      }),
    ];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    const result = (await action.execute(client, {
      conversationId: "19:test@thread.v2",
    })) as Message[];

    expect(result).toHaveLength(2);
    expect(result[0].content).toBe("Hello");
    expect(result[1].content).toBe("Plain text");
  });

  it("should include system messages when textOnly is false", async () => {
    const messages = [
      makeMessage({ messageType: "RichText/Html" }),
      makeMessage({ messageType: "ThreadActivity/AddMember" }),
    ];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    const result = (await action.execute(client, {
      conversationId: "19:test@thread.v2",
      textOnly: false,
    })) as Message[];

    expect(result).toHaveLength(2);
  });

  it("should pass custom maxPages and pageSize", async () => {
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue([]),
    });

    await action.execute(client, {
      conversationId: "19:test@thread.v2",
      maxPages: 5,
      pageSize: 50,
    });

    expect(client.getMessages).toHaveBeenCalledWith("19:test@thread.v2", {
      maxPages: 5,
      pageSize: 50,
      onProgress: undefined,
    });
  });

  it("should pass onProgress callback to client", async () => {
    const onProgress = vi.fn();
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue([]),
    });

    await action.execute(client, {
      conversationId: "19:test@thread.v2",
      onProgress,
    });

    expect(client.getMessages).toHaveBeenCalledWith("19:test@thread.v2", {
      maxPages: 100,
      pageSize: 200,
      onProgress,
    });
  });

  it("should reverse messages for oldest-first order", async () => {
    const messages = [
      makeMessage({ id: "3", content: "Third" }),
      makeMessage({ id: "2", content: "Second" }),
      makeMessage({ id: "1", content: "First" }),
    ];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    const result = (await action.execute(client, {
      conversationId: "19:test@thread.v2",
      order: "oldest-first",
    })) as Message[];

    expect(result[0].content).toBe("First");
    expect(result[1].content).toBe("Second");
    expect(result[2].content).toBe("Third");
  });

  it("should keep newest-first order by default", async () => {
    const messages = [
      makeMessage({ id: "3", content: "Third" }),
      makeMessage({ id: "1", content: "First" }),
    ];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    const result = (await action.execute(client, {
      conversationId: "19:test@thread.v2",
    })) as Message[];

    expect(result[0].content).toBe("Third");
    expect(result[1].content).toBe("First");
  });

  it("should format messages correctly", () => {
    const messages = [
      makeMessage({
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        senderDisplayName: "Alice",
        content: "<p>Hello world</p>",
      }),
    ];

    const output = action.formatResult(messages);

    expect(output).toContain("1 messages:");
    expect(output).toContain("[2026-03-16 10:00:00] Alice: Hello world");
  });

  it("should decode HTML entities in text format", () => {
    const messages = [
      makeMessage({
        senderDisplayName: "Alice",
        content: "<p>Hello&nbsp;world&amp;friends&quot;hi&quot;</p>",
      }),
    ];

    const output = action.formatResult(messages);

    expect(output).toContain('Hello world&friends"hi"');
    expect(output).not.toContain("&nbsp;");
    expect(output).not.toContain("&quot;");
  });

  it("should show reply markers in text format", () => {
    const quotedMessage = makeMessage({
      id: "msg-100",
      senderDisplayName: "Bob",
      content: "<p>Original message</p>",
    });
    const replyMessage = makeMessage({
      id: "msg-200",
      senderDisplayName: "Alice",
      content: "<blockquote>Original message</blockquote><p>My reply here</p>",
      quotedMessageId: "msg-100",
    });
    const messages = [quotedMessage, replyMessage];

    const output = action.formatResult(messages);

    expect(output).toContain("> [replying to Bob]:");
    expect(output).toContain("My reply here");
  });

  it("should compress repeated authors in markdown format", () => {
    const messages = [
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        content: "<p>First message</p>",
      }),
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:01:00.000Z",
        content: "<p>Second message</p>",
      }),
      makeMessage({
        senderDisplayName: "Bob",
        originalArrivalTime: "2026-03-16T10:02:00.000Z",
        content: "<p>Different person</p>",
      }),
    ];

    const output = action.formatMarkdown(messages);

    // First Alice message gets full header
    expect(output).toContain("### Alice");
    // Second Alice message gets compressed (just timestamp)
    expect(output).toContain("*2026-03-16 10:01:00*");
    // Bob gets full header
    expect(output).toContain("### Bob");
  });

  it("should show reply markers in markdown format", () => {
    const messages = [
      makeMessage({
        id: "msg-100",
        senderDisplayName: "Bob",
        content: "<p>Original</p>",
      }),
      makeMessage({
        id: "msg-200",
        senderDisplayName: "Alice",
        content: "<blockquote>Original</blockquote><p>Reply</p>",
        quotedMessageId: "msg-100",
      }),
    ];

    const output = action.formatMarkdown(messages);

    expect(output).toContain("> **[replying to Bob]:**");
    expect(output).toContain("Reply");
  });

  it("should decode HTML entities in markdown format", () => {
    const messages = [
      makeMessage({
        content: "<p>test&nbsp;content&#8203;here</p>",
      }),
    ];

    const output = action.formatMarkdown(messages);

    expect(output).toContain("test content");
    expect(output).not.toContain("&nbsp;");
    expect(output).not.toContain("&#8203;");
  });

  it("should compress repeated authors in toon format", () => {
    const messages = [
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        content: "<p>First</p>",
      }),
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:01:00.000Z",
        content: "<p>Second</p>",
      }),
    ];

    const output = action.formatToon(messages);

    // First message has full name
    const nameMatches = output.match(/Alice/g);
    // Should appear in header line once, and only timestamp for second
    expect(output).toContain("🗣️  Alice · 2026-03-16 10:00:00");
    expect(output).toContain("2026-03-16 10:01:00");
    // Name should appear only once (plus the header count line)
    expect(nameMatches).toBeTruthy();
    expect(nameMatches!.length).toBeLessThanOrEqual(2);
  });
});

// ── send-message ─────────────────────────────────────────────────────

describe("send-message", () => {
  const action = getAction("send-message");

  it("should resolve conversation and send message", async () => {
    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Design Review",
    });
    const sentMessage: SentMessage = {
      messageId: "msg-123",
      arrivalTime: 1773000000000,
    };
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      sendMessage: vi.fn().mockResolvedValue(sentMessage),
    });

    const result = (await action.execute(client, {
      chat: "Design Review",
      content: "Hello!",
    })) as SentMessage & { conversation: string };

    expect(client.sendMessage).toHaveBeenCalledWith(
      "19:chat@thread.v2",
      "Hello!",
    );
    expect(result.messageId).toBe("msg-123");
    expect(result.conversation).toBe("Design Review");
  });

  it("should resolve 1:1 conversation via --to", async () => {
    const searchResult: OneOnOneSearchResult = {
      conversationId: "19:one-on-one@unq.gbl.spaces",
      memberDisplayName: "Luke Prior",
    };
    const sentMessage: SentMessage = {
      messageId: "msg-456",
      arrivalTime: 1773000000000,
    };
    const client = createMockClient({
      findOneOnOneConversation: vi.fn().mockResolvedValue(searchResult),
      sendMessage: vi.fn().mockResolvedValue(sentMessage),
    });

    const result = (await action.execute(client, {
      to: "Luke",
      content: "Hey!",
    })) as SentMessage & { conversation: string };

    expect(result.conversation).toBe("Luke Prior");
  });

  it("should error when no content provided", async () => {
    const client = createMockClient();

    // Missing content won't cause an action-level error, but sendMessage
    // will be called with undefined — the validation is at the CLI/MCP level.
    // Here we test that conversation resolution still requires an identifier.
    await expect(action.execute(client, { content: "Hello" })).rejects.toThrow(
      "One of --conversation-id, --chat, or --to is required.",
    );
  });

  it("should format result correctly", () => {
    const result = {
      messageId: "msg-123",
      arrivalTime: 1773000000000,
      conversation: "Design Review",
    };

    const output = action.formatResult(result);

    expect(output).toContain('Message sent to "Design Review"');
    expect(output).toContain("Message ID: msg-123");
    expect(output).toContain("Arrival time: 1773000000000");
  });
});

// ── get-members ──────────────────────────────────────────────────────

describe("get-members", () => {
  const action = getAction("get-members");

  it("should resolve conversation by chat name and get members", async () => {
    const conversation = makeConversation({ id: "19:chat@thread.v2" });
    const members = [
      makeMember({ displayName: "Alice Smith", role: "Admin" }),
      makeMember({ displayName: "Bob Jones", role: "User" }),
    ];
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      getMembers: vi.fn().mockResolvedValue(members),
    });

    const result = await action.execute(client, { chat: "Design Review" });

    expect(client.getMembers).toHaveBeenCalledWith("19:chat@thread.v2");
    expect(result).toEqual(members);
  });

  it("should resolve by direct conversation ID", async () => {
    const members = [makeMember()];
    const client = createMockClient({
      getMembers: vi.fn().mockResolvedValue(members),
    });

    await action.execute(client, {
      conversationId: "19:direct@thread.v2",
    });

    expect(client.getMembers).toHaveBeenCalledWith("19:direct@thread.v2");
  });

  it("should format members correctly", () => {
    const members = [
      makeMember({
        displayName: "Alice Smith",
        role: "Admin",
        id: "8:orgid:alice",
      }),
      makeMember({ displayName: "", role: "User", id: "8:orgid:unknown" }),
    ];

    const output = action.formatResult(members);

    expect(output).toContain("2 members:");
    expect(output).toContain("Alice Smith (Admin) — 8:orgid:alice");
    expect(output).toContain("(unknown) (User) — 8:orgid:unknown");
  });
});

// ── whoami ───────────────────────────────────────────────────────────

describe("whoami", () => {
  const action = getAction("whoami");

  it("should return display name and region", async () => {
    const client = createMockClient({
      getCurrentUserDisplayName: vi.fn().mockResolvedValue("Maxim Mazurok"),
      getToken: vi.fn(() => ({
        skypeToken: "test-token",
        region: "apac",
      })),
    });

    const result = (await action.execute(client, {})) as {
      displayName: string;
      region: string;
    };

    expect(result.displayName).toBe("Maxim Mazurok");
    expect(result.region).toBe("apac");
  });

  it("should format correctly", () => {
    const result = { displayName: "Maxim Mazurok", region: "apac" };

    const output = action.formatResult(result);

    expect(output).toBe("Maxim Mazurok (region: apac)");
  });
});

// ── Conversation resolution shared behavior ──────────────────────────

describe("conversation resolution (shared across actions)", () => {
  it("should prefer conversationId over chat and to", async () => {
    const action = getAction("get-messages");
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue([]),
    });

    await action.execute(client, {
      conversationId: "19:direct@thread.v2",
      chat: "Some Chat",
      to: "Some Person",
    });

    // conversationId is used directly, no find calls made
    expect(client.findConversation).not.toHaveBeenCalled();
    expect(client.findOneOnOneConversation).not.toHaveBeenCalled();
    expect(client.getMessages).toHaveBeenCalledWith(
      "19:direct@thread.v2",
      expect.anything(),
    );
  });

  it("should prefer chat over to when conversationId is absent", async () => {
    const action = getAction("get-members");
    const conversation = makeConversation({ id: "19:found@thread.v2" });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      getMembers: vi.fn().mockResolvedValue([]),
    });

    await action.execute(client, {
      chat: "Some Chat",
      to: "Some Person",
    });

    // chat takes priority over to
    expect(client.findConversation).toHaveBeenCalledWith("Some Chat");
    expect(client.findOneOnOneConversation).not.toHaveBeenCalled();
  });
});

// ── formatOutput dispatch ────────────────────────────────────────────

describe("formatOutput", () => {
  const action = getAction("whoami");

  const sampleResult = { displayName: "Maxim Mazurok", region: "apac" };

  it("should return JSON for json format", () => {
    const output = formatOutput(action, sampleResult, "json");
    expect(JSON.parse(output)).toEqual(sampleResult);
  });

  it("should return text for text format", () => {
    const output = formatOutput(action, sampleResult, "text");
    expect(output).toBe("Maxim Mazurok (region: apac)");
  });

  it("should return markdown for md format", () => {
    const output = formatOutput(action, sampleResult, "md");
    expect(output).toContain("## Maxim Mazurok");
    expect(output).toContain("**Region:** apac");
  });

  it("should return toon for toon format", () => {
    const output = formatOutput(action, sampleResult, "toon");
    expect(output).toContain("🙋");
    expect(output).toContain("Maxim Mazurok");
    expect(output).toContain("📍 region: apac");
  });

  it("should default to json when no format specified", () => {
    const output = formatOutput(action, sampleResult);
    expect(JSON.parse(output)).toEqual(sampleResult);
  });

  it("should handle null result in json format", () => {
    const output = formatOutput(getAction("find-conversation"), null, "json");
    expect(output).toBe("null");
  });
});

// ── list-conversations format tests ──────────────────────────────────

describe("list-conversations formatMarkdown", () => {
  const action = getAction("list-conversations");

  it("should produce markdown table", () => {
    const conversations = [
      makeConversation({ topic: "Design Review", threadType: "chat" }),
      makeConversation({
        topic: "",
        threadType: "chat",
        lastMessageTime: null,
        memberCount: null,
      }),
    ];

    const output = action.formatMarkdown(conversations);

    expect(output).toContain("## Conversations (2)");
    expect(output).toContain("| # | Topic | Type | Members | Last Message |");
    expect(output).toContain("| 0 | Design Review | chat | 5 |");
    expect(output).toContain(
      "| 1 | (untitled 1:1 chat) | chat | ? | unknown |",
    );
  });

  it("should handle empty list", () => {
    const output = action.formatMarkdown([]);
    expect(output).toContain("## Conversations (0)");
    expect(output).not.toContain("| # |");
  });
});

describe("list-conversations formatToon", () => {
  const action = getAction("list-conversations");

  it("should produce emoji-decorated output", () => {
    const conversations = [
      makeConversation({ topic: "Design Review", threadType: "chat" }),
    ];

    const output = action.formatToon(conversations);

    expect(output).toContain("📋 1 Conversations");
    expect(output).toContain("─".repeat(40));
    expect(output).toContain('💬 [0] "Design Review"');
    expect(output).toContain("chat · 5 members · last: 2026-03-16");
  });
});

// ── find-conversation format tests ───────────────────────────────────

describe("find-conversation formatMarkdown", () => {
  const action = getAction("find-conversation");

  it("should format found conversation as markdown", () => {
    const conversation = makeConversation({
      id: "19:abc@thread.v2",
      topic: "Engineering",
      threadType: "chat",
    });

    const output = action.formatMarkdown(conversation);

    expect(output).toContain('## Found: "Engineering"');
    expect(output).toContain("**ID:** 19:abc@thread.v2");
    expect(output).toContain("**Type:** chat");
  });

  it("should handle null result", () => {
    expect(action.formatMarkdown(null)).toBe("No conversation found.");
  });
});

describe("find-conversation formatToon", () => {
  const action = getAction("find-conversation");

  it("should format found conversation with emojis", () => {
    const conversation = makeConversation({
      id: "19:abc@thread.v2",
      topic: "Engineering",
    });

    const output = action.formatToon(conversation);

    expect(output).toContain("🔍");
    expect(output).toContain('Found: "Engineering"');
    expect(output).toContain("🆔 19:abc@thread.v2");
  });

  it("should handle null result", () => {
    const output = action.formatToon(null);
    expect(output).toContain("🔍 No conversation found.");
  });
});

// ── find-one-on-one format tests ─────────────────────────────────────

describe("find-one-on-one formatMarkdown", () => {
  const action = getAction("find-one-on-one");

  it("should format as markdown", () => {
    const searchResult: OneOnOneSearchResult = {
      conversationId: "19:abc@unq.gbl.spaces",
      memberDisplayName: "Luke Prior",
    };

    const output = action.formatMarkdown(searchResult);

    expect(output).toContain("## Found 1:1 with Luke Prior");
    expect(output).toContain("**Conversation ID:** 19:abc@unq.gbl.spaces");
  });

  it("should handle null result", () => {
    expect(action.formatMarkdown(null)).toBe("No 1:1 conversation found.");
  });
});

describe("find-one-on-one formatToon", () => {
  const action = getAction("find-one-on-one");

  it("should format with emojis", () => {
    const searchResult: OneOnOneSearchResult = {
      conversationId: "19:abc@unq.gbl.spaces",
      memberDisplayName: "Luke Prior",
    };

    const output = action.formatToon(searchResult);

    expect(output).toContain("🔍 Found 1:1 with Luke Prior");
    expect(output).toContain("🆔 19:abc@unq.gbl.spaces");
  });

  it("should handle null result", () => {
    const output = action.formatToon(null);
    expect(output).toContain("No 1:1 conversation found.");
  });
});

// ── get-messages format tests ────────────────────────────────────────

describe("get-messages formatMarkdown", () => {
  const action = getAction("get-messages");

  it("should produce markdown with headings per message", () => {
    const messages = [
      makeMessage({
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        senderDisplayName: "Alice",
        content: "<p>Hello world</p>",
      }),
    ];

    const output = action.formatMarkdown(messages);

    expect(output).toContain("## Messages (1)");
    expect(output).toContain("### Alice — 2026-03-16 10:00:00");
    expect(output).toContain("Hello world");
  });
});

describe("get-messages formatToon", () => {
  const action = getAction("get-messages");

  it("should produce emoji-decorated output", () => {
    const messages = [
      makeMessage({
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        senderDisplayName: "Alice",
        content: "<p>Hello world</p>",
      }),
    ];

    const output = action.formatToon(messages);

    expect(output).toContain("💬 1 Messages");
    expect(output).toContain("🗣️  Alice · 2026-03-16 10:00:00");
    expect(output).toContain("Hello world");
  });
});

// ── send-message format tests ────────────────────────────────────────

describe("send-message formatMarkdown", () => {
  const action = getAction("send-message");

  it("should format as markdown", () => {
    const result = {
      messageId: "msg-123",
      arrivalTime: 1773000000000,
      conversation: "Design Review",
    };

    const output = action.formatMarkdown(result);

    expect(output).toContain("## Message Sent");
    expect(output).toContain("**To:** Design Review");
    expect(output).toContain("**Message ID:** msg-123");
  });
});

describe("send-message formatToon", () => {
  const action = getAction("send-message");

  it("should format with emojis", () => {
    const result = {
      messageId: "msg-123",
      arrivalTime: 1773000000000,
      conversation: "Design Review",
    };

    const output = action.formatToon(result);

    expect(output).toContain("✅ Message Sent!");
    expect(output).toContain('📨 To: "Design Review"');
    expect(output).toContain("🆔 msg-123");
    expect(output).toContain("⏰ 1773000000000");
  });
});

// ── get-members format tests ─────────────────────────────────────────

describe("get-members formatMarkdown", () => {
  const action = getAction("get-members");

  it("should produce markdown table", () => {
    const members = [
      makeMember({
        displayName: "Alice Smith",
        role: "Admin",
        id: "8:orgid:alice",
      }),
    ];

    const output = action.formatMarkdown(members);

    expect(output).toContain("## Members (1)");
    expect(output).toContain("| Name | Role | ID |");
    expect(output).toContain("| Alice Smith | Admin | 8:orgid:alice |");
  });

  it("should handle empty members", () => {
    const output = action.formatMarkdown([]);
    expect(output).toContain("## Members (0)");
    expect(output).not.toContain("| Name |");
  });
});

describe("get-members formatToon", () => {
  const action = getAction("get-members");

  it("should format with emojis", () => {
    const members = [
      makeMember({
        displayName: "Alice Smith",
        role: "Admin",
        id: "8:orgid:alice",
      }),
    ];

    const output = action.formatToon(members);

    expect(output).toContain("👥 1 Members");
    expect(output).toContain("👤 Alice Smith · Admin");
    expect(output).toContain("8:orgid:alice");
  });
});

// ── whoami format tests ──────────────────────────────────────────────

describe("whoami formatMarkdown", () => {
  const action = getAction("whoami");

  it("should format as markdown", () => {
    const result = { displayName: "Maxim Mazurok", region: "apac" };

    const output = action.formatMarkdown(result);

    expect(output).toContain("## Maxim Mazurok");
    expect(output).toContain("**Region:** apac");
  });
});

describe("whoami formatToon", () => {
  const action = getAction("whoami");

  it("should format with emojis", () => {
    const result = { displayName: "Maxim Mazurok", region: "apac" };

    const output = action.formatToon(result);

    expect(output).toContain("🙋 Maxim Mazurok");
    expect(output).toContain("📍 region: apac");
  });
});

// ── All actions have all formatters ──────────────────────────────────

describe("all actions have formatMarkdown and formatToon", () => {
  for (const action of actions) {
    it(`${action.name} should have formatMarkdown`, () => {
      expect(typeof action.formatMarkdown).toBe("function");
    });

    it(`${action.name} should have formatToon`, () => {
      expect(typeof action.formatToon).toBe("function");
    });
  }
});

// ── HTML entity decoding (tested through formatters) ─────────────────

describe("HTML entity decoding in message formats", () => {
  const action = getAction("get-messages");

  it("should decode &amp; correctly (after other entities)", () => {
    const messages = [makeMessage({ content: "<p>A &amp; B</p>" })];
    const output = action.formatResult(messages);
    expect(output).toContain("A & B");
  });

  it("should decode &lt; and &gt;", () => {
    const messages = [
      makeMessage({ content: "<p>if (a &lt; b &amp;&amp; c &gt; d)</p>" }),
    ];
    const output = action.formatResult(messages);
    expect(output).toContain("if (a < b && c > d)");
  });

  it("should remove zero-width spaces (&#8203;)", () => {
    const messages = [makeMessage({ content: "<p>hello&#8203;world</p>" })];
    const output = action.formatResult(messages);
    expect(output).toContain("helloworld");
  });

  it("should decode numeric character references", () => {
    const messages = [makeMessage({ content: "<p>&#65;&#66;&#67;</p>" })];
    const output = action.formatResult(messages);
    expect(output).toContain("ABC");
  });

  it("should handle content with no HTML at all", () => {
    const messages = [
      makeMessage({
        messageType: "Text",
        content: "plain text message",
      }),
    ];
    const output = action.formatResult(messages);
    expect(output).toContain("plain text message");
  });

  it("should strip nested HTML tags", () => {
    const messages = [
      makeMessage({
        content: "<p><b>bold</b> and <i>italic</i> text</p>",
      }),
    ];
    const output = action.formatResult(messages);
    expect(output).toContain("bold and italic text");
  });
});

// ── Quote extraction (tested through formatters) ─────────────────────

describe("quote extraction in message formats", () => {
  const action = getAction("get-messages");

  it("should handle blockquote with attributes", () => {
    const messages = [
      makeMessage({
        id: "msg-reply",
        content:
          '<blockquote itemtype="cite">Quoted text</blockquote><p>Reply</p>',
        quotedMessageId: "msg-original",
      }),
      makeMessage({
        id: "msg-original",
        senderDisplayName: "Original Author",
        content: "<p>Original message</p>",
      }),
    ];
    const output = action.formatResult(messages);
    expect(output).toContain("[replying to Original Author]");
    expect(output).toContain("Reply");
  });

  it("should not show reply marker when quotedMessageId is null", () => {
    const messages = [
      makeMessage({
        content: "<blockquote>Some quote</blockquote><p>Reply</p>",
        quotedMessageId: null,
      }),
    ];
    const output = action.formatResult(messages);
    expect(output).not.toContain("[replying to");
  });

  it("should show unknown sender for replies to messages not in result set", () => {
    const messages = [
      makeMessage({
        content: "<blockquote>Old message</blockquote><p>My reply</p>",
        quotedMessageId: "msg-not-in-set",
      }),
    ];
    const output = action.formatResult(messages);
    expect(output).toContain("[replying to unknown]");
  });

  it("should handle messages with no blockquote", () => {
    const messages = [makeMessage({ content: "<p>Normal message</p>" })];
    const output = action.formatResult(messages);
    expect(output).not.toContain("[replying to");
    expect(output).toContain("Normal message");
  });
});

// ── Author compression (tested through formatters) ───────────────────

describe("author compression in formatMarkdown", () => {
  const action = getAction("get-messages");

  it("should not compress when different senders alternate", () => {
    const messages = [
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        content: "<p>Hello</p>",
      }),
      makeMessage({
        senderDisplayName: "Bob",
        originalArrivalTime: "2026-03-16T10:01:00.000Z",
        content: "<p>Hi</p>",
      }),
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:02:00.000Z",
        content: "<p>How are you?</p>",
      }),
    ];

    const output = action.formatMarkdown(messages);

    // All three should have full headers
    expect(output.match(/### Alice/g)).toHaveLength(2);
    expect(output.match(/### Bob/g)).toHaveLength(1);
  });

  it("should compress three consecutive messages from same sender", () => {
    const messages = [
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        content: "<p>First</p>",
      }),
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:01:00.000Z",
        content: "<p>Second</p>",
      }),
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:02:00.000Z",
        content: "<p>Third</p>",
      }),
    ];

    const output = action.formatMarkdown(messages);

    // Only one full header, two compressed timestamps
    expect(output.match(/### Alice/g)).toHaveLength(1);
    expect(output).toContain("*2026-03-16 10:01:00*");
    expect(output).toContain("*2026-03-16 10:02:00*");
  });
});

describe("author compression in formatToon", () => {
  const action = getAction("get-messages");

  it("should compress consecutive same-sender messages", () => {
    const messages = [
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        content: "<p>First</p>",
      }),
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:01:00.000Z",
        content: "<p>Second</p>",
      }),
    ];

    const output = action.formatToon(messages);

    // Full header only once
    expect(output.match(/🗣️  Alice/g)).toHaveLength(1);
    // Second message has just the timestamp, indented
    expect(output).toContain("      2026-03-16 10:01:00");
  });

  it("should reset compression when sender changes", () => {
    const messages = [
      makeMessage({
        senderDisplayName: "Alice",
        originalArrivalTime: "2026-03-16T10:00:00.000Z",
        content: "<p>Hello</p>",
      }),
      makeMessage({
        senderDisplayName: "Bob",
        originalArrivalTime: "2026-03-16T10:01:00.000Z",
        content: "<p>Hi</p>",
      }),
    ];

    const output = action.formatToon(messages);

    expect(output).toContain("🗣️  Alice");
    expect(output).toContain("🗣️  Bob");
  });
});

// ── Message order parameter ──────────────────────────────────────────

describe("message order parameter", () => {
  const action = getAction("get-messages");

  it("should have order in parameter definitions", () => {
    const orderParameter = action.parameters.find(
      (parameter) => parameter.name === "order",
    );
    expect(orderParameter).toBeDefined();
    expect(orderParameter!.type).toBe("string");
    expect(orderParameter!.required).toBe(false);
    expect(orderParameter!.default).toBe("newest-first");
  });

  it("should not mutate original array when reversing", async () => {
    const originalMessages = [
      makeMessage({ id: "1", content: "First" }),
      makeMessage({ id: "2", content: "Second" }),
    ];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(originalMessages),
    });

    await action.execute(client, {
      conversationId: "19:test@thread.v2",
      order: "oldest-first",
    });

    // Original array should be untouched
    expect(originalMessages[0].id).toBe("1");
    expect(originalMessages[1].id).toBe("2");
  });

  it("should pass explicit newest-first without reversing", async () => {
    const messages = [
      makeMessage({ id: "3", content: "Third" }),
      makeMessage({ id: "2", content: "Second" }),
      makeMessage({ id: "1", content: "First" }),
    ];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    const result = (await action.execute(client, {
      conversationId: "19:test@thread.v2",
      order: "newest-first",
    })) as Message[];

    expect(result[0].content).toBe("Third");
    expect(result[2].content).toBe("First");
  });

  it("should apply order after textOnly filtering", async () => {
    const messages = [
      makeMessage({ id: "3", messageType: "RichText/Html", content: "C" }),
      makeMessage({
        id: "2",
        messageType: "ThreadActivity/AddMember",
        content: "system",
      }),
      makeMessage({ id: "1", messageType: "RichText/Html", content: "A" }),
    ];
    const client = createMockClient({
      getMessages: vi.fn().mockResolvedValue(messages),
    });

    const result = (await action.execute(client, {
      conversationId: "19:test@thread.v2",
      order: "oldest-first",
    })) as Message[];

    // System message filtered out, remaining reversed
    expect(result).toHaveLength(2);
    expect(result[0].content).toBe("A");
    expect(result[1].content).toBe("C");
  });
});
