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
  fetchProfiles,
  postMessage,
  editMessage,
  deleteMessage,
  fetchUserProperties,
  parseRawMessage,
  ApiAuthError,
  ApiRateLimitError,
  extractTranscriptUrl,
  extractMeetingTitle,
  isSuccessfulRecording,
  parseVtt,
  fetchTranscriptVtt,
  fetchTranscript,
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

function mockFetchResponse(
  body: unknown,
  status = 200,
  responseHeaders: Record<string, string> = {},
): void {
  (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? "OK" : "Error",
    json: () => Promise.resolve(body),
    text: () => Promise.resolve(JSON.stringify(body)),
    headers: new Headers(responseHeaders),
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
      memberType: "person",
    });
    expect(members[1]).toEqual({
      id: "8:orgid:user2",
      displayName: "",
      role: "User",
      memberType: "person",
    });
  });

  it("should detect bot members by 28: prefix", async () => {
    mockFetchResponse({
      members: [
        { id: "8:orgid:user1", role: "Admin" },
        { id: "28:bot-guid-here", role: "User" },
      ],
    });

    const members = await fetchMembers(testToken, "conv-id");

    expect(members[0].memberType).toBe("person");
    expect(members[1].memberType).toBe("bot");
  });
});

const tokenWithBearer: TeamsToken = {
  skypeToken: "test-token-abc123",
  region: "apac",
  bearerToken: "test-bearer-token-xyz",
};

describe("fetchProfiles", () => {
  it("should resolve profiles from MRIs via middle-tier API", async () => {
    mockFetchResponse({
      value: [
        {
          mri: "8:orgid:user1",
          displayName: "Alice Smith",
          email: "alice@example.com",
          jobTitle: "Engineer",
          userType: "Member",
        },
        {
          mri: "8:orgid:user2",
          displayName: "Bob Jones",
          email: "bob@example.com",
          jobTitle: "Manager",
          userType: "Member",
        },
      ],
    });

    const profiles = await fetchProfiles(tokenWithBearer, [
      "8:orgid:user1",
      "8:orgid:user2",
    ]);

    expect(profiles).toHaveLength(2);
    expect(profiles[0]).toEqual({
      mri: "8:orgid:user1",
      displayName: "Alice Smith",
      email: "alice@example.com",
      jobTitle: "Engineer",
      userType: "Member",
    });
    expect(profiles[1].displayName).toBe("Bob Jones");

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(fetchCall[0]).toContain("/users/fetchShortProfile");
    expect(fetchCall[1].method).toBe("POST");
    expect(fetchCall[1].headers.Authorization).toBe(
      "Bearer test-bearer-token-xyz",
    );
  });

  it("should throw ApiAuthError when no bearer token is available", async () => {
    await expect(
      fetchProfiles(testToken, ["8:orgid:user1"]),
    ).rejects.toBeInstanceOf(ApiAuthError);
    expect(globalThis.fetch).not.toHaveBeenCalled();
  });

  it("should return empty when MRI list is empty", async () => {
    const profiles = await fetchProfiles(tokenWithBearer, []);
    expect(profiles).toEqual([]);
    expect(globalThis.fetch).not.toHaveBeenCalled();
  });

  it("should return empty on non-auth failure", async () => {
    mockFetchResponse({}, 500);
    const profiles = await fetchProfiles(tokenWithBearer, ["8:orgid:user1"]);
    expect(profiles).toEqual([]);
  });

  it("should throw ApiAuthError on 401", async () => {
    mockFetchResponse({}, 401);
    await expect(
      fetchProfiles(tokenWithBearer, ["8:orgid:user1"]),
    ).rejects.toThrow(ApiAuthError);
  });
});

describe("postMessage", () => {
  it("should default to markdown format and convert to HTML", async () => {
    mockFetchResponse({ OriginalArrivalTime: 1773000000000 });

    const result = await postMessage(
      testToken,
      "conv-id",
      "**Hello!**",
      "Test User",
    );

    expect(result.messageId).toBe("1773000000000");
    expect(result.arrivalTime).toBe(1773000000000);

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(fetchCall[0]).toContain("/users/ME/conversations/conv-id/messages");
    expect(fetchCall[1].method).toBe("POST");

    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    expect(sentBody.content).toContain("<strong>Hello!</strong>");
    expect(sentBody.messagetype).toBe("RichText/Html");
    expect(sentBody.imdisplayname).toBe("Test User");
  });

  it("should send plain text when format is text", async () => {
    mockFetchResponse({ OriginalArrivalTime: 1773000000000 });

    await postMessage(testToken, "conv-id", "Hello!", "Test User", "text");

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    expect(sentBody.content).toBe("Hello!");
    expect(sentBody.messagetype).toBe("Text");
  });

  it("should pass through raw HTML when format is html", async () => {
    mockFetchResponse({ OriginalArrivalTime: 1773000000000 });

    const htmlContent = "<b>Bold</b> and <i>italic</i>";
    await postMessage(testToken, "conv-id", htmlContent, "Test User", "html");

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    expect(sentBody.content).toBe(htmlContent);
    expect(sentBody.messagetype).toBe("RichText/Html");
  });

  it("should convert markdown features to HTML", async () => {
    mockFetchResponse({ OriginalArrivalTime: 1773000000000 });

    const markdownContent = [
      "# Heading",
      "",
      "**bold** and *italic*",
      "",
      "- item one",
      "- item two",
    ].join("\n");
    await postMessage(
      testToken,
      "conv-id",
      markdownContent,
      "Test User",
      "markdown",
    );

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    const content = sentBody.content as string;
    expect(content).toContain("<h1>Heading</h1>");
    expect(content).toContain("<strong>bold</strong>");
    expect(content).toContain("<em>italic</em>");
    expect(content).toContain("<li>item one</li>");
    expect(content).toContain("<li>item two</li>");
    expect(sentBody.messagetype).toBe("RichText/Html");
  });

  it("should use OriginalArrivalTime as messageId", async () => {
    mockFetchResponse({ OriginalArrivalTime: 1773000000000 });

    const result = await postMessage(
      testToken,
      "conv-id",
      "Hello!",
      "Test User",
    );

    expect(result.messageId).toBe("1773000000000");
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

describe("editMessage", () => {
  it("should send PUT request with skypeeditedid and return edit time", async () => {
    mockFetchResponse({ edittime: "2026-03-24T10:00:00.000Z" });

    const result = await editMessage(
      testToken,
      "conv-id",
      "msg-123",
      "**Updated!**",
      "Test User",
    );

    expect(result.messageId).toBe("msg-123");
    expect(result.editTime).toBe("2026-03-24T10:00:00.000Z");

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(fetchCall[0]).toContain(
      "/users/ME/conversations/conv-id/messages/msg-123",
    );
    expect(fetchCall[1].method).toBe("PUT");

    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    expect(sentBody.content).toContain("<strong>Updated!</strong>");
    expect(sentBody.messagetype).toBe("RichText/Html");
    expect(sentBody.skypeeditedid).toBe("msg-123");
    expect(sentBody.imdisplayname).toBe("Test User");
  });

  it("should send plain text when format is text", async () => {
    mockFetchResponse({ edittime: "2026-03-24T10:00:00.000Z" });

    await editMessage(
      testToken,
      "conv-id",
      "msg-456",
      "plain update",
      "Test User",
      "text",
    );

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    expect(sentBody.content).toBe("plain update");
    expect(sentBody.messagetype).toBe("Text");
  });

  it("should pass through raw HTML when format is html", async () => {
    mockFetchResponse({ edittime: "2026-03-24T10:00:00.000Z" });

    const htmlContent = "<b>Bold edit</b>";
    await editMessage(
      testToken,
      "conv-id",
      "msg-789",
      htmlContent,
      "Test User",
      "html",
    );

    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    const sentBody = JSON.parse(fetchCall[1].body as string) as Record<
      string,
      unknown
    >;
    expect(sentBody.content).toBe(htmlContent);
    expect(sentBody.messagetype).toBe("RichText/Html");
  });

  it("should throw on failure with error body", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: false,
      status: 403,
      statusText: "Forbidden",
      text: () => Promise.resolve("Access denied"),
    });

    await expect(
      editMessage(testToken, "conv-id", "msg-123", "Hello!", "User"),
    ).rejects.toThrow("Failed to edit message: 403 Forbidden — Access denied");
  });

  it("should throw ApiAuthError on 401", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
      text: () => Promise.resolve("Token expired"),
    });

    await expect(
      editMessage(testToken, "conv-id", "msg-123", "Hello!", "User"),
    ).rejects.toBeInstanceOf(ApiAuthError);
  });

  it("should handle empty response body", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      text: () => Promise.resolve(""),
      headers: new Headers(),
    });

    const result = await editMessage(
      testToken,
      "conv-id",
      "msg-empty",
      "Updated!",
      "Test User",
    );

    expect(result.messageId).toBe("msg-empty");
    expect(result.editTime).toBeTruthy();
  });
});

describe("deleteMessage", () => {
  it("should send DELETE request to correct URL", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      text: () => Promise.resolve(""),
      headers: new Headers(),
    });

    const result = await deleteMessage(testToken, "conv-id", "msg-123");

    expect(globalThis.fetch).toHaveBeenCalledWith(
      expect.stringContaining(
        "/users/ME/conversations/conv-id/messages/msg-123",
      ),
      expect.objectContaining({ method: "DELETE" }),
    );
    expect(result.messageId).toBe("msg-123");
  });

  it("should throw on failure with error body", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: false,
      status: 403,
      statusText: "Forbidden",
      text: () => Promise.resolve("Access denied"),
    });

    await expect(
      deleteMessage(testToken, "conv-id", "msg-123"),
    ).rejects.toThrow(
      "Failed to delete message: 403 Forbidden — Access denied",
    );
  });

  it("should throw ApiAuthError on 401", async () => {
    (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
      text: () => Promise.resolve("Token expired"),
    });

    await expect(
      deleteMessage(testToken, "conv-id", "msg-123"),
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
    expect(message.followers).toEqual([]);
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
    expect(message.followers).toEqual([]);
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
    expect(message.followers).toEqual([]);
  });

  it("should separate follow entries from reactions", () => {
    const message = parseRawMessage({
      id: "123",
      messagetype: "Text",
      content: "msg",
      properties: {
        emotions: [
          { key: "like", users: [{ mri: "8:orgid:user1", time: 1773000000 }] },
          {
            key: "follow",
            users: [
              { mri: "8:orgid:follower1", time: 1773000001, value: "0" },
              { mri: "8:orgid:follower2", time: 1773000002, value: "0" },
              { mri: "8:orgid:unfollowed", time: 1773000003, value: "1" },
            ],
          },
          { key: "heart", users: [{ mri: "8:orgid:user2", time: 1773000004 }] },
        ],
      },
    });

    expect(message.reactions).toHaveLength(2);
    expect(message.reactions[0].key).toBe("like");
    expect(message.reactions[1].key).toBe("heart");

    expect(message.followers).toHaveLength(2);
    expect(message.followers[0].mri).toBe("8:orgid:follower1");
    expect(message.followers[1].mri).toBe("8:orgid:follower2");
  });

  it("should exclude unfollowed users from followers", () => {
    const message = parseRawMessage({
      id: "123",
      messagetype: "Text",
      content: "msg",
      properties: {
        emotions: [
          {
            key: "follow",
            users: [
              { mri: "8:orgid:active", time: 1773000001, value: "0" },
              { mri: "8:orgid:removed", time: 1773000002, value: "1" },
            ],
          },
        ],
      },
    });

    expect(message.reactions).toEqual([]);
    expect(message.followers).toHaveLength(1);
    expect(message.followers[0].mri).toBe("8:orgid:active");
    expect(message.followers[0].time).toBe(1773000001);
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

describe("rate limit retry (429 handling)", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("should retry on 429 and succeed when next attempt returns 200", async () => {
    mockFetchResponse({ errorCode: 429, message: "Rate limited" }, 429, {
      "Retry-After": "1",
    });
    mockFetchResponse({ conversations: [] });

    const promise = fetchConversations(testToken, 50);
    await vi.advanceTimersByTimeAsync(1_000);
    const result = await promise;

    expect(result).toEqual([]);
    expect(globalThis.fetch).toHaveBeenCalledTimes(2);
  });

  it("should use exponential backoff on repeated 429 responses", async () => {
    // 3 consecutive 429s, then success
    mockFetchResponse({ errorCode: 429 }, 429, { "Retry-After": "1" });
    mockFetchResponse({ errorCode: 429 }, 429, { "Retry-After": "1" });
    mockFetchResponse({ errorCode: 429 }, 429, { "Retry-After": "1" });
    mockFetchResponse({ conversations: [] });

    const promise = fetchConversations(testToken, 50);

    // Attempt 0: backoff = 1s * 2^0 = 1s
    await vi.advanceTimersByTimeAsync(1_000);
    // Attempt 1: backoff = 1s * 2^1 = 2s
    await vi.advanceTimersByTimeAsync(2_000);
    // Attempt 2: backoff = 1s * 2^2 = 4s
    await vi.advanceTimersByTimeAsync(4_000);

    const result = await promise;
    expect(result).toEqual([]);
    expect(globalThis.fetch).toHaveBeenCalledTimes(4);
  });

  it("should throw ApiRateLimitError after exhausting all retries", async () => {
    for (let i = 0; i <= 5; i++) {
      mockFetchResponse({ errorCode: 429, message: "Rate limited" }, 429, {
        "Retry-After": "1",
      });
    }

    const promise = fetchConversations(testToken, 50);
    // Prevent unhandled rejection warning while advancing timers
    const errorPromise = promise.then(
      () => {
        throw new Error("Expected promise to reject");
      },
      (error: unknown) => error,
    );

    // Advance past all retry delays: 1s, 2s, 4s, 8s, 16s
    for (const seconds of [1, 2, 4, 8, 16]) {
      await vi.advanceTimersByTimeAsync(seconds * 1_000);
    }

    const caughtError = await errorPromise;
    expect(caughtError).toBeInstanceOf(ApiRateLimitError);
    expect((caughtError as ApiRateLimitError).message).toMatch(
      /Rate limit exceeded after 6 attempts/,
    );
    expect(globalThis.fetch).toHaveBeenCalledTimes(6);
  });

  it("should default to 2s backoff when Retry-After header is missing", async () => {
    mockFetchResponse({ errorCode: 429 }, 429);
    mockFetchResponse({ conversations: [] });

    const promise = fetchConversations(testToken, 50);

    // Default: 2s * 2^0 = 2s
    await vi.advanceTimersByTimeAsync(2_000);

    const result = await promise;
    expect(result).toEqual([]);
    expect(globalThis.fetch).toHaveBeenCalledTimes(2);
  });

  it("should retry 429 for fetchMessagesPage", async () => {
    mockFetchResponse({ errorCode: 429 }, 429, { "Retry-After": "1" });
    mockFetchResponse({ messages: [], _metadata: {} });

    const promise = fetchMessagesPage(testToken, "conv-id", 50);
    await vi.advanceTimersByTimeAsync(1_000);
    const result = await promise;

    expect(result.messages).toEqual([]);
    expect(globalThis.fetch).toHaveBeenCalledTimes(2);
  });
});

// ── Transcript helpers ────────────────────────────────────────────────

function mockFetchTextResponse(
  text: string,
  status = 200,
  responseHeaders: Record<string, string> = {},
): void {
  (globalThis.fetch as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
    ok: status >= 200 && status < 300,
    status,
    statusText: status === 200 ? "OK" : "Error",
    json: () => Promise.resolve({}),
    text: () => Promise.resolve(text),
    headers: new Headers(responseHeaders),
  });
}

const sampleRecordingXml = `
<URIObject type="Video.1/Message.1">
  <OriginalName v="Test Meeting Title"/>
  <RecordingStatus status="Success"/>
  <RecordingContent>
    <item type="amsTranscript" uri="https://as-prod.asyncgw.teams.microsoft.com/v1/objects/abc123/views/transcript"/>
    <item type="onedriveForBusinessTranscript" uri="https://example.sharepoint.com/_api/v2.1/drives/driveId/items/itemId" driveId="driveId" driveItemId="itemId" id="transcriptId"/>
    <item type="onedriveForBusinessVideo" uri="https://example.sharepoint.com/..." driveId="driveId" driveItemId="videoItemId"/>
  </RecordingContent>
</URIObject>`;

const sampleVtt = `WEBVTT

00:00:00.000 --> 00:00:05.240
<v Alice Smith>Hello everyone, let&#39;s get started.</v>

00:00:05.240 --> 00:00:10.500
<v Alice Smith>Today we&#39;re discussing the project status.</v>

00:00:10.500 --> 00:00:15.800
<v Bob Jones>Sounds good. I have some updates.</v>

00:00:15.800 --> 00:00:22.100
<v Alice Smith>Great, go ahead Bob.</v>
`;

describe("extractTranscriptUrl", () => {
  it("should extract AMS transcript URL from recording XML", () => {
    const url = extractTranscriptUrl(sampleRecordingXml);
    expect(url).toBe(
      "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/abc123/views/transcript",
    );
  });

  it("should return null when no amsTranscript item exists", () => {
    const xml =
      '<URIObject><RecordingContent><item type="someOther" uri="http://example.com"/></RecordingContent></URIObject>';
    expect(extractTranscriptUrl(xml)).toBeNull();
  });

  it("should return null for empty content", () => {
    expect(extractTranscriptUrl("")).toBeNull();
  });
});

describe("extractMeetingTitle", () => {
  it("should extract meeting title from OriginalName element", () => {
    expect(extractMeetingTitle(sampleRecordingXml)).toBe("Test Meeting Title");
  });

  it("should return 'Unknown Meeting' when OriginalName is missing", () => {
    expect(extractMeetingTitle("<URIObject/>")).toBe("Unknown Meeting");
  });
});

describe("isSuccessfulRecording", () => {
  it('should return true when status is "Success"', () => {
    expect(isSuccessfulRecording(sampleRecordingXml)).toBe(true);
  });

  it('should return false when status is not "Success"', () => {
    const startedXml = '<RecordingStatus status="Started"/>';
    expect(isSuccessfulRecording(startedXml)).toBe(false);
  });

  it("should return false for empty content", () => {
    expect(isSuccessfulRecording("")).toBe(false);
  });
});

describe("parseVtt", () => {
  it("should parse VTT content into transcript entries", () => {
    const entries = parseVtt(sampleVtt);

    expect(entries).toHaveLength(4);
    expect(entries[0]).toEqual({
      speaker: "Alice Smith",
      startTime: "00:00:00.000",
      endTime: "00:00:05.240",
      text: "Hello everyone, let's get started.",
    });
    expect(entries[2]).toEqual({
      speaker: "Bob Jones",
      startTime: "00:00:10.500",
      endTime: "00:00:15.800",
      text: "Sounds good. I have some updates.",
    });
  });

  it("should decode HTML entities in speaker names and text", () => {
    const vtt = `WEBVTT

00:00:00.000 --> 00:00:05.000
<v O&#39;Brien>This &amp; that are &lt;important&gt;.</v>
`;
    const entries = parseVtt(vtt);
    expect(entries).toHaveLength(1);
    expect(entries[0].speaker).toBe("O'Brien");
    expect(entries[0].text).toBe("This & that are <important>.");
  });

  it("should return empty array for empty or header-only VTT", () => {
    expect(parseVtt("")).toEqual([]);
    expect(parseVtt("WEBVTT\n\n")).toEqual([]);
  });
});

describe("fetchTranscriptVtt", () => {
  it("should fetch VTT content using skype_token auth header", async () => {
    mockFetchTextResponse(sampleVtt);

    const result = await fetchTranscriptVtt(
      testToken,
      "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/abc123/views/transcript",
    );

    expect(result).toBe(sampleVtt);
    const fetchCall = (globalThis.fetch as ReturnType<typeof vi.fn>).mock
      .calls[0];
    expect(fetchCall[1].headers.Authorization).toBe(
      "skype_token test-token-abc123",
    );
  });

  it("should throw ApiAuthError on 401", async () => {
    mockFetchTextResponse("", 401);

    await expect(
      fetchTranscriptVtt(testToken, "https://example.com/transcript"),
    ).rejects.toBeInstanceOf(ApiAuthError);
  });

  it("should throw Error on non-auth failure", async () => {
    mockFetchTextResponse("", 500);

    await expect(
      fetchTranscriptVtt(testToken, "https://example.com/transcript"),
    ).rejects.toThrow("Failed to fetch transcript: 500");
  });
});

describe("fetchTranscript", () => {
  it("should find recording message, fetch VTT, and parse entries", async () => {
    // First call: fetchMessagesPage for finding the recording message
    mockFetchResponse({
      messages: [
        {
          id: "msg-1",
          messagetype: "Text",
          content: "Normal message",
          properties: {},
        },
        {
          id: "msg-2",
          messagetype: "RichText/Media_CallRecording",
          content: sampleRecordingXml,
          properties: {},
        },
      ],
      _metadata: {},
    });

    // Second call: fetchTranscriptVtt
    mockFetchTextResponse(sampleVtt);

    const result = await fetchTranscript(testToken, "conv-id");

    expect(result.meetingTitle).toBe("Test Meeting Title");
    expect(result.rawVtt).toBe(sampleVtt);
    expect(result.entries).toHaveLength(4);
    expect(result.entries[0].speaker).toBe("Alice Smith");
  });

  it("should throw when no recording message is found", async () => {
    mockFetchResponse({
      messages: [
        {
          id: "msg-1",
          messagetype: "Text",
          content: "Just a normal message",
          properties: {},
        },
      ],
      _metadata: {},
    });

    await expect(fetchTranscript(testToken, "conv-id")).rejects.toThrow(
      "No meeting transcript found",
    );
  });

  it("should skip recording messages without Success status", async () => {
    const startedXml =
      '<URIObject><RecordingStatus status="Started"/><RecordingContent><item type="amsTranscript" uri="https://example.com/t"/></RecordingContent></URIObject>';

    mockFetchResponse({
      messages: [
        {
          id: "msg-1",
          messagetype: "RichText/Media_CallRecording",
          content: startedXml,
          properties: {},
        },
      ],
      _metadata: {},
    });

    await expect(fetchTranscript(testToken, "conv-id")).rejects.toThrow(
      "No meeting transcript found",
    );
  });

  it("should paginate through messages to find a recording", async () => {
    // First page: no recording, but has backwardLink
    mockFetchResponse({
      messages: [
        {
          id: "msg-1",
          messagetype: "Text",
          content: "no transcript here",
          properties: {},
        },
      ],
      _metadata: {
        backwardLink: "https://apac.ng.msg.teams.microsoft.com/v1/page2",
      },
    });

    // Second page: has the recording message
    mockFetchResponse({
      messages: [
        {
          id: "msg-2",
          messagetype: "RichText/Media_CallRecording",
          content: sampleRecordingXml,
          properties: {},
        },
      ],
      _metadata: {},
    });

    // Third call: fetch the VTT
    mockFetchTextResponse(sampleVtt);

    const result = await fetchTranscript(testToken, "conv-id");

    expect(result.meetingTitle).toBe("Test Meeting Title");
    expect(result.entries).toHaveLength(4);
    expect(globalThis.fetch).toHaveBeenCalledTimes(3);
  });
});
