/**
 * Unit tests for TeamsClient (src/teams-client.ts).
 *
 * These tests mock the API layer to verify that the client class
 * correctly orchestrates operations like pagination, filtering,
 * and conversation search.
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { TeamsClient } from "../../src/teams-client.js";
import * as api from "../../src/api.js";
import type {
  Conversation,
  Message,
  Member,
  MessagesPage,
  SentMessage,
} from "../../src/types.js";

vi.mock("../../src/api.js");

const mockedApi = vi.mocked(api);

function makeConversation(overrides: Partial<Conversation> = {}): Conversation {
  return {
    id: "19:test@thread.space",
    topic: "Test Chat",
    threadType: "chat",
    version: 1,
    lastMessageTime: null,
    memberCount: null,
    ...overrides,
  };
}

function makeMessage(overrides: Partial<Message> = {}): Message {
  return {
    id: "1773000000000",
    messageType: "Text",
    senderMri: "8:orgid:user-1",
    senderDisplayName: "Test User",
    content: "Hello",
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

function makeMessagesPage(
  messages: Message[],
  backwardLink: string | null = null,
): MessagesPage {
  return { messages, backwardLink, syncState: null };
}

beforeEach(() => {
  vi.resetAllMocks();
});

describe("TeamsClient.fromToken", () => {
  it("should create a client with the provided token", () => {
    const client = TeamsClient.fromToken("my-token", "emea");
    const token = client.getToken();

    expect(token.skypeToken).toBe("my-token");
    expect(token.region).toBe("emea");
  });

  it("should default region to apac", () => {
    const client = TeamsClient.fromToken("my-token");
    expect(client.getToken().region).toBe("apac");
  });

  it("should return a copy of the token (not a reference)", () => {
    const client = TeamsClient.fromToken("original");
    const token1 = client.getToken();
    const token2 = client.getToken();

    expect(token1).toEqual(token2);
    expect(token1).not.toBe(token2);
  });
});

describe("listConversations", () => {
  it("should filter out system streams by default", async () => {
    mockedApi.fetchConversations.mockResolvedValueOnce([
      makeConversation({ topic: "Real Chat", threadType: "chat" }),
      makeConversation({
        topic: "",
        threadType: "streamofannotations",
      }),
      makeConversation({
        topic: "",
        threadType: "streamofnotifications",
      }),
    ]);

    const client = TeamsClient.fromToken("token");
    const conversations = await client.listConversations();

    expect(conversations).toHaveLength(1);
    expect(conversations[0].topic).toBe("Real Chat");
  });

  it("should include system streams when excludeSystemStreams is false", async () => {
    mockedApi.fetchConversations.mockResolvedValueOnce([
      makeConversation({ threadType: "chat" }),
      makeConversation({ threadType: "streamofannotations" }),
    ]);

    const client = TeamsClient.fromToken("token");
    const conversations = await client.listConversations({
      excludeSystemStreams: false,
    });

    expect(conversations).toHaveLength(2);
  });

  it("should pass pageSize to the API", async () => {
    mockedApi.fetchConversations.mockResolvedValueOnce([]);

    const client = TeamsClient.fromToken("token");
    await client.listConversations({ pageSize: 100 });

    expect(mockedApi.fetchConversations).toHaveBeenCalledWith(
      expect.objectContaining({ skypeToken: "token" }),
      100,
    );
  });
});

describe("findConversation", () => {
  it("should find by partial topic match (case-insensitive)", async () => {
    mockedApi.fetchConversations.mockResolvedValueOnce([
      makeConversation({ topic: "Engineering Updates" }),
      makeConversation({ topic: "Random" }),
    ]);

    const client = TeamsClient.fromToken("token");
    const result = await client.findConversation("engineering");

    expect(result).not.toBeNull();
    expect(result!.topic).toBe("Engineering Updates");
  });

  it("should return null when no match", async () => {
    mockedApi.fetchConversations.mockResolvedValueOnce([
      makeConversation({ topic: "Unrelated Chat" }),
    ]);

    const client = TeamsClient.fromToken("token");
    const result = await client.findConversation("nonexistent");

    expect(result).toBeNull();
  });

  it("should skip conversations without topics", async () => {
    mockedApi.fetchConversations.mockResolvedValueOnce([
      makeConversation({ topic: "" }),
      makeConversation({ topic: "Match This" }),
    ]);

    const client = TeamsClient.fromToken("token");
    const result = await client.findConversation("match");

    expect(result!.topic).toBe("Match This");
  });
});

describe("findOneOnOneConversation", () => {
  it("should match self-chat when searching for own name", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes", threadType: "streamofnotes" }),
      makeConversation({ id: "19:chat1", threadType: "chat", topic: "" }),
    ]);

    // getCurrentUserDisplayName calls fetchConversations + fetchMessagesPage
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Maxim Mazurok" })]),
    );

    const client = TeamsClient.fromToken("token");
    const result = await client.findOneOnOneConversation("Maxim");

    expect(result).not.toBeNull();
    expect(result!.conversationId).toBe("48:notes");
    expect(result!.memberDisplayName).toContain("Maxim Mazurok");
  });

  it("should find 1:1 chat by scanning message senders", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "19:chat1", threadType: "chat", topic: "" }),
    ]);

    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([
        makeMessage({ senderDisplayName: "Alice Smith" }),
        makeMessage({ senderDisplayName: "Bob Jones" }),
      ]),
    );

    // Also need to mock for getCurrentUserDisplayName (called for self-chat check
    // when there's no 48:notes conversation)

    const client = TeamsClient.fromToken("token");
    const result = await client.findOneOnOneConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.memberDisplayName).toBe("Alice Smith");
  });

  it("should return null when no match found", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "19:chat1", threadType: "chat", topic: "" }),
    ]);

    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Someone Else" })]),
    );

    const client = TeamsClient.fromToken("token");
    const result = await client.findOneOnOneConversation("Nonexistent Person");

    expect(result).toBeNull();
  });
});

describe("getMessages", () => {
  it("should paginate through multiple pages", async () => {
    mockedApi.fetchMessagesPage
      .mockResolvedValueOnce(
        makeMessagesPage(
          [makeMessage({ id: "1" }), makeMessage({ id: "2" })],
          "https://next-page",
        ),
      )
      .mockResolvedValueOnce(
        makeMessagesPage([makeMessage({ id: "3" })], null),
      );

    const client = TeamsClient.fromToken("token");
    const messages = await client.getMessages("conv-id");

    expect(messages).toHaveLength(3);
    expect(mockedApi.fetchMessagesPage).toHaveBeenCalledTimes(2);
  });

  it("should stop at maxPages", async () => {
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage()], "https://always-more"),
    );

    const client = TeamsClient.fromToken("token");
    const messages = await client.getMessages("conv-id", { maxPages: 3 });

    expect(messages).toHaveLength(3);
    expect(mockedApi.fetchMessagesPage).toHaveBeenCalledTimes(3);
  });

  it("should invoke onProgress callback", async () => {
    mockedApi.fetchMessagesPage
      .mockResolvedValueOnce(
        makeMessagesPage([makeMessage(), makeMessage()], "https://next"),
      )
      .mockResolvedValueOnce(makeMessagesPage([makeMessage()], null));

    const progressCounts: number[] = [];
    const client = TeamsClient.fromToken("token");

    await client.getMessages("conv-id", {
      onProgress: (count) => progressCounts.push(count),
    });

    expect(progressCounts).toEqual([2, 3]);
  });

  it("should stop when backwardLink is null", async () => {
    mockedApi.fetchMessagesPage.mockResolvedValueOnce(
      makeMessagesPage([makeMessage()], null),
    );

    const client = TeamsClient.fromToken("token");
    await client.getMessages("conv-id");

    expect(mockedApi.fetchMessagesPage).toHaveBeenCalledTimes(1);
  });
});

describe("sendMessage", () => {
  it("should resolve display name and send", async () => {
    // getCurrentUserDisplayName will call fetchConversations + fetchMessagesPage
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );

    const expectedResult: SentMessage = {
      messageId: "msg-1",
      arrivalTime: 1773000000000,
    };
    mockedApi.postMessage.mockResolvedValue(expectedResult);

    const client = TeamsClient.fromToken("token");
    const result = await client.sendMessage("conv-id", "Hello!");

    expect(result).toEqual(expectedResult);
    expect(mockedApi.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({ skypeToken: "token" }),
      "conv-id",
      "Hello!",
      "Test User",
    );
  });
});

describe("getMembers", () => {
  it("should delegate to fetchMembers", async () => {
    const expectedMembers: Member[] = [
      { id: "8:orgid:user1", displayName: "Alice", role: "Admin" },
    ];
    mockedApi.fetchMembers.mockResolvedValue(expectedMembers);

    const client = TeamsClient.fromToken("token");
    const members = await client.getMembers("conv-id");

    expect(members).toEqual(expectedMembers);
  });
});

describe("getCurrentUserDisplayName", () => {
  it("should get name from self-chat messages", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Maxim Mazurok" })]),
    );

    const client = TeamsClient.fromToken("token");
    const name = await client.getCurrentUserDisplayName();

    expect(name).toBe("Maxim Mazurok");
  });

  it("should cache the result", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Cached Name" })]),
    );

    const client = TeamsClient.fromToken("token");
    await client.getCurrentUserDisplayName();
    await client.getCurrentUserDisplayName();

    // fetchConversations should only be called once due to caching
    expect(mockedApi.fetchConversations).toHaveBeenCalledTimes(1);
  });

  it("should fallback to user properties endpoint", async () => {
    mockedApi.fetchConversations.mockResolvedValue([]);
    mockedApi.fetchUserProperties.mockResolvedValue({
      displayname: "From Properties",
    });

    const client = TeamsClient.fromToken("token");
    const name = await client.getCurrentUserDisplayName();

    expect(name).toBe("From Properties");
  });

  it("should return 'Unknown User' when all methods fail", async () => {
    mockedApi.fetchConversations.mockResolvedValue([]);
    mockedApi.fetchUserProperties.mockRejectedValue(new Error("Network error"));

    const client = TeamsClient.fromToken("token");
    const name = await client.getCurrentUserDisplayName();

    expect(name).toBe("Unknown User");
  });
});
