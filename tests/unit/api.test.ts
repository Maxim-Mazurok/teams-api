/**
 * Unit tests for the REST API layer (src/api.ts).
 *
 * These tests mock global `fetch` and verify that the API layer
 * correctly transforms raw HTTP responses into typed objects.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  fetchConversations,
  fetchMessagesPage,
  fetchMembers,
  postMessage,
  fetchUserProperties,
  parseRawMessage,
  ApiAuthError,
} from "../../src/api.js";
import type { TeamsToken } from "../../src/types.js";

const testToken: TeamsToken = {
  skypeToken: "test-token-abc123",
  region: "apac",
};

const originalFetch = globalThis.fetch;

beforeEach(() => {
  globalThis.fetch = vi.fn();
});

afterEach(() => {
  globalThis.fetch = originalFetch;
});

function mockFetchResponse(body: unknown, status = 200): void {
  (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? "OK" : "Error",
    json: () => Promise.resolve(body),
    text: () => Promise.resolve(JSON.stringify(body)),
  });
}

describe("fetchConversations", () => {
  it("should parse conversations from API response", async () => {
    mockFetchResponse({
      conversations: [
        {
          id: "19:thread-id@thread.space",
          version: 12345,
          threadProperties: {
            topic: "Test Chat",
            threadType: "chat",
            memberCount: "5",
          },
          properties: {
            lastimreceivedtime: "2026-03-16T10:00:00.000Z",
          },
        },
        {
          id: "19:another-thread@thread.space",
          version: 67890,
          threadProperties: {
            threadType: "topic",
          },
          properties: {
            displayName: "General Channel",
          },
        },
      ],
    });

    const conversations = await fetchConversations(testToken, 50);

    expect(conversations).toHaveLength(2);
    expect(conversations[0]).toEqual({
      id: "19:thread-id@thread.space",
      topic: "Test Chat",
      threadType: "chat",
      version: 12345,
      lastMessageTime: "2026-03-16T10:00:00.000Z",
      memberCount: 5,
    });
    expect(conversations[1]).toEqual({
      id: "19:another-thread@thread.space",
      topic: "General Channel",
      threadType: "topic",
      version: 67890,
      lastMessageTime: null,
      memberCount: null,
    });
  });

  it("should build the correct URL with pageSize", async () => {
    mockFetchResponse({ conversations: [] });

    await fetchConversations(testToken, 100);

    expect(globalThis.fetch).toHaveBeenCalledWith(
      "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations?view=mychats&pageSize=100",
      { headers: { Authentication: "skypetoken=test-token-abc123" } },
    );
  });

  it("should throw on non-OK response", async () => {
    mockFetchResponse({}, 401);

    await expect(fetchConversations(testToken, 50)).rejects.toThrow(
      "Authentication failed: 401",
    );
  });

  it("should throw ApiAuthError on 401", async () => {
    mockFetchResponse({}, 401);

    await expect(fetchConversations(testToken, 50)).rejects.toBeInstanceOf(
      ApiAuthError,
    );
  });

  it("should throw regular Error on non-401 errors", async () => {
    mockFetchResponse({}, 500);

    const error = await fetchConversations(testToken, 50).catch(
      (caughtError: unknown) => caughtError,
    );
    expect(error).toBeInstanceOf(Error);
    expect(error).not.toBeInstanceOf(ApiAuthError);
  });

  it("should handle empty conversation list", async () => {
    mockFetchResponse({ conversations: [] });

    const conversations = await fetchConversations(testToken, 50);
    expect(conversations).toEqual([]);
  });

  it("should handle missing conversations field", async () => {
    mockFetchResponse({});

    const conversations = await fetchConversations(testToken, 50);
    expect(conversations).toEqual([]);
  });
});

describe("fetchMessagesPage", () => {
  it("should parse messages and metadata", async () => {
    mockFetchResponse({
      messages: [
        {
          id: "1773000000000",
          messagetype: "RichText/Html",
          from: "8:orgid:user-id",
          imdisplayname: "Test User",
          content: "<p>Hello world</p>",
          originalarrivaltime: "2026-03-16T10:00:00.000Z",
          composetime: "2026-03-16T10:00:00.000Z",
          properties: {},
        },
      ],
      _metadata: {
        backwardLink:
          "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/test/messages?startTime=123&pageSize=200",
        syncState: "sync-token-xyz",
      },
    });

    const page = await fetchMessagesPage(testToken, "test-conv", 200);

    expect(page.messages).toHaveLength(1);
    expect(page.messages[0].id).toBe("1773000000000");
    expect(page.messages[0].senderDisplayName).toBe("Test User");
    expect(page.messages[0].content).toBe("<p>Hello world</p>");
    expect(page.backwardLink).toContain("startTime=123");
    expect(page.syncState).toBe("sync-token-xyz");
  });

  it("should use backwardLink URL when provided", async () => {
    const customUrl = "https://apac.ng.msg.teams.microsoft.com/v1/custom-link";
    mockFetchResponse({ messages: [], _metadata: {} });

    await fetchMessagesPage(testToken, "any-id", 50, customUrl);

    expect(globalThis.fetch).toHaveBeenCalledWith(customUrl, expect.anything());
  });

  it("should return null backwardLink when not present", async () => {
    mockFetchResponse({ messages: [], _metadata: {} });

    const page = await fetchMessagesPage(testToken, "test", 50);
    expect(page.backwardLink).toBeNull();
    expect(page.syncState).toBeNull();
  });
});

describe("fetchMembers", () => {
  it("should parse members from API response", async () => {
    mockFetchResponse({
      members: [
        { id: "8:orgid:user1", userDisplayName: "Alice", role: "Admin" },
        { id: "8:orgid:user2", role: "User" },
      ],
    });

    const members = await fetchMembers(testToken, "conv-id");

    expect(members).toHaveLength(2);
    expect(members[0]).toEqual({
      id: "8:orgid:user1",
      displayName: "Alice",
      role: "Admin",
    });
    expect(members[1]).toEqual({
      id: "8:orgid:user2",
      displayName: "",
      role: "User",
    });
  });
});

describe("postMessage", () => {
  it("should send POST request with correct body", async () => {
    mockFetchResponse({ OriginalArrivalTime: 1773000000000, id: "msg-123" });

    const result = await postMessage(
      testToken,
      "conv-id",
      "Hello!",
      "Test User",
    );

    expect(result.messageId).toBe("msg-123");
    expect(result.arrivalTime).toBe(1773000000000);

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(fetchCall[0]).toContain("/users/ME/conversations/conv-id/messages");
    expect(fetchCall[1].method).toBe("POST");

    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    expect(sentBody.content).toBe("Hello!");
    expect(sentBody.messagetype).toBe("Text");
    expect(sentBody.imdisplayname).toBe("Test User");
  });

  it("should use clientMessageId as fallback when server returns no id", async () => {
    mockFetchResponse({ OriginalArrivalTime: 1773000000000 });

    const result = await postMessage(
      testToken,
      "conv-id",
      "Hello!",
      "Test User",
    );

    expect(result.messageId).toBeTruthy();
    expect(Number(result.messageId)).toBeGreaterThan(0);
  });

  it("should throw on failure with error body", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: false,
      status: 403,
      statusText: "Forbidden",
      text: () => Promise.resolve("Access denied"),
    });

    await expect(
      postMessage(testToken, "conv-id", "Hello!", "User"),
    ).rejects.toThrow("Failed to send message: 403 Forbidden — Access denied");
  });
  it("should throw ApiAuthError on 401", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
      text: () => Promise.resolve("Token expired"),
    });

    await expect(
      postMessage(testToken, "conv-id", "Hello!", "User"),
    ).rejects.toBeInstanceOf(ApiAuthError);
  });
});

describe("fetchUserProperties", () => {
  it("should return parsed properties", async () => {
    mockFetchResponse({ displayname: "Test User", someOtherProp: "value" });

    const properties = await fetchUserProperties(testToken);

    expect(properties.displayname).toBe("Test User");
  });
});

describe("parseRawMessage", () => {
  it("should parse a basic text message", () => {
    const message = parseRawMessage({
      id: "123",
      messagetype: "Text",
      from: "8:orgid:user1",
      imdisplayname: "Alice",
      content: "Hello world",
      originalarrivaltime: "2026-03-16T10:00:00.000Z",
      composetime: "2026-03-16T10:00:00.000Z",
      properties: {},
    });

    expect(message.id).toBe("123");
    expect(message.messageType).toBe("Text");
    expect(message.senderMri).toBe("8:orgid:user1");
    expect(message.senderDisplayName).toBe("Alice");
    expect(message.content).toBe("Hello world");
    expect(message.isDeleted).toBe(false);
    expect(message.reactions).toEqual([]);
    expect(message.mentions).toEqual([]);
    expect(message.quotedMessageId).toBeNull();
  });

  it("should parse emotions from JSON string", () => {
    const message = parseRawMessage({
      id: "123",
      messagetype: "Text",
      content: "msg",
      properties: {
        emotions: JSON.stringify([
          {
            key: "like",
            users: [{ mri: "8:orgid:user1", time: 1773000000 }],
          },
        ]),
      },
    });

    expect(message.reactions).toHaveLength(1);
    expect(message.reactions[0].key).toBe("like");
    expect(message.reactions[0].users[0].mri).toBe("8:orgid:user1");
  });

  it("should parse emotions from array", () => {
    const message = parseRawMessage({
      id: "123",
      messagetype: "Text",
      content: "msg",
      properties: {
        emotions: [
          { key: "heart", users: [{ mri: "8:orgid:user2", time: 1773000000 }] },
        ],
      },
    });

    expect(message.reactions).toHaveLength(1);
    expect(message.reactions[0].key).toBe("heart");
  });

  it("should parse mentions from JSON string", () => {
    const message = parseRawMessage({
      id: "123",
      messagetype: "RichText/Html",
      content: "<at>@Alice</at>",
      properties: {
        mentions: JSON.stringify([
          { id: "8:orgid:user1", displayName: "Alice" },
        ]),
      },
    });

    expect(message.mentions).toHaveLength(1);
    expect(message.mentions[0].displayName).toBe("Alice");
  });

  it("should detect quoted message ID from HTML content", () => {
    const content =
      '<div itemscope itemtype="http://schema.skype.com/Reply" itemid="1773000000"><p>Quoted text</p></div>';
    const message = parseRawMessage({
      id: "456",
      messagetype: "RichText/Html",
      content,
      properties: {},
    });

    expect(message.quotedMessageId).toBe("1773000000");
  });

  it("should detect deleted messages", () => {
    const deletedViaType = parseRawMessage({
      id: "1",
      messagetype: "MessageDelete",
      content: "",
      properties: {},
    });
    expect(deletedViaType.isDeleted).toBe(true);

    const deletedViaProperty = parseRawMessage({
      id: "2",
      messagetype: "Text",
      content: "deleted",
      properties: { deletetime: "2026-03-16T10:00:00.000Z" },
    });
    expect(deletedViaProperty.isDeleted).toBe(true);
  });

  it("should handle malformed emotions gracefully", () => {
    const message = parseRawMessage({
      id: "123",
      messagetype: "Text",
      content: "msg",
      properties: { emotions: "not-valid-json{" },
    });

    expect(message.reactions).toEqual([]);
  });

  it("should handle missing fields gracefully", () => {
    const message = parseRawMessage({});

    expect(message.id).toBe("");
    expect(message.messageType).toBe("");
    expect(message.senderMri).toBe("");
    expect(message.senderDisplayName).toBe("");
    expect(message.content).toBe("");
  });
});
