/**
 * Unit tests for TeamsClient (src/teams-client.ts).
 *
 * These tests mock the API layer to verify that the client class
 * correctly orchestrates operations like pagination, filtering,
 * and conversation search.
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { TeamsClient } from "../../src/teams-client.js";
import * as chatService from "../../src/api/chat-service.js";
import * as middleTier from "../../src/api/middle-tier.js";
import * as substrate from "../../src/api/substrate.js";
import * as attachments from "../../src/api/attachments.js";
import { ApiAuthError } from "../../src/api/common.js";
import * as tokenStore from "../../src/token-store.js";
import * as autoLogin from "../../src/auth/auto-login.js";
import type {
  Conversation,
  Message,
  Member,
  MessagesPage,
  SentMessage,
  ScheduledMessage,
} from "../../src/types.js";

vi.mock("../../src/api/chat-service.js", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../src/api/chat-service.js")>();
  return {
    ...actual,
    fetchConversations: vi.fn(),
    fetchMessagesPage: vi.fn(),
    fetchMembers: vi.fn(),
    postMessage: vi.fn(),
    postScheduledMessage: vi.fn(),
    fetchUserProperties: vi.fn(),
  };
});
vi.mock("../../src/api/middle-tier.js", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../src/api/middle-tier.js")>();
  return {
    ...actual,
    fetchProfiles: vi.fn(),
  };
});
vi.mock("../../src/api/substrate.js", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../src/api/substrate.js")>();
  return {
    ...actual,
    searchPeople: vi.fn(),
    searchChats: vi.fn(),
  };
});
vi.mock("../../src/token-store.js");
vi.mock("../../src/auth/auto-login.js");
vi.mock("../../src/api/attachments.js", async (importOriginal) => {
  const actual =
    await importOriginal<typeof import("../../src/api/attachments.js")>();
  return {
    ...actual,
    uploadSharePointFile: vi.fn(),
    buildFilesPropertyJson: vi.fn(),
    uploadAmsImage: vi.fn(),
  };
});

const mockedApi = {
  ...vi.mocked(chatService),
  ...vi.mocked(middleTier),
  ...vi.mocked(substrate),
};
const mockedAttachments = vi.mocked(attachments);
const mockedTokenStore = vi.mocked(tokenStore);
const mockedAuth = {
  ...vi.mocked(autoLogin),
};

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
    followers: [],
    mentions: [],
    quotedMessageId: null,
    images: [],
    files: [],
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

  it("should preserve optional bearer and substrate tokens", () => {
    const client = TeamsClient.fromToken(
      "my-token",
      "emea",
      "bearer-token",
      "substrate-token",
    );

    expect(client.getToken()).toEqual({
      skypeToken: "my-token",
      region: "emea",
      bearerToken: "bearer-token",
      substrateToken: "substrate-token",
    });
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

  it("should fall back to Substrate chat search when no topic match", async () => {
    const untitledChat = makeConversation({
      id: "19:alice-bob@unq.gbl.spaces",
      topic: "",
      threadType: "chat",
    });
    mockedApi.fetchConversations.mockResolvedValueOnce([untitledChat]);

    mockedApi.searchChats.mockResolvedValue([
      {
        name: "",
        threadId: "19:alice-bob@unq.gbl.spaces",
        threadType: "Chat",
        matchingMembers: [
          { displayName: "Alice Smith", mri: "8:orgid:alice-uuid" },
        ],
        chatMembers: [],
        totalMemberCount: 2,
      },
    ]);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.id).toBe("19:alice-bob@unq.gbl.spaces");
    expect(mockedApi.searchChats).toHaveBeenCalledWith(
      expect.objectContaining({ substrateToken: "substrate-token" }),
      "Alice",
      5,
    );
  });

  it("should not call Substrate search when topic match exists", async () => {
    mockedApi.fetchConversations.mockResolvedValueOnce([
      makeConversation({ topic: "Alice Design Review" }),
    ]);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.topic).toBe("Alice Design Review");
    expect(mockedApi.searchChats).not.toHaveBeenCalled();
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
      makeMessagesPage([makeMessage({ senderDisplayName: "Alice Smith" })]),
    );

    const client = TeamsClient.fromToken("token");
    const result = await client.findOneOnOneConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.conversationId).toBe("48:notes");
    expect(result!.memberDisplayName).toContain("Alice Smith");
  });

  it("should find 1:1 chat via Substrate search when token is available", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "19:chat1", threadType: "chat", topic: "" }),
    ]);

    mockedApi.searchPeople.mockResolvedValue([
      {
        displayName: "Alice Smith",
        mri: "8:orgid:alice-uuid",
        email: "alice@example.com",
        jobTitle: "Engineer",
        department: "Dev",
        objectId: "alice-uuid",
      },
    ]);

    mockedApi.searchChats.mockResolvedValue([
      {
        name: "",
        threadId: "19:alice-chat-thread@unq.gbl.spaces",
        threadType: "Chat",
        matchingMembers: [
          { displayName: "Alice Smith", mri: "8:orgid:alice-uuid" },
        ],
        chatMembers: [],
        totalMemberCount: 2,
      },
    ]);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findOneOnOneConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.conversationId).toBe("19:alice-chat-thread@unq.gbl.spaces");
    expect(result!.memberDisplayName).toBe("Alice Smith");
  });

  it("should fall back to conversation ID matching when chat search returns no 1:1", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({
        id: "19:my-uuid_alice-uuid@unq.gbl.spaces",
        threadType: "chat",
        topic: "",
      }),
    ]);

    mockedApi.searchPeople.mockResolvedValue([
      {
        displayName: "Alice Smith",
        mri: "8:orgid:alice-uuid",
        email: "alice@example.com",
        jobTitle: "Engineer",
        department: "Dev",
        objectId: "alice-uuid",
      },
    ]);

    // No chat results from search
    mockedApi.searchChats.mockResolvedValue([]);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findOneOnOneConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.conversationId).toBe("19:my-uuid_alice-uuid@unq.gbl.spaces");
    expect(result!.memberDisplayName).toBe("Alice Smith");
  });

  it("should fall back to message scanning when no substrate token", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "19:chat1", threadType: "chat", topic: "" }),
    ]);

    // searchPeople throws when substrate token is missing
    mockedApi.searchPeople.mockRejectedValue(
      new ApiAuthError("Substrate token is missing"),
    );

    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([
        makeMessage({
          senderDisplayName: "Alice Smith",
          senderMri: "8:orgid:alice-uuid",
        }),
      ]),
    );

    // No substrate token — falls back to message scanning
    const client = TeamsClient.fromToken("token");
    const result = await client.findOneOnOneConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.memberDisplayName).toBe("Alice Smith");
  });

  it("should return null when no match found via search", async () => {
    mockedApi.fetchConversations.mockResolvedValue([]);

    mockedApi.searchPeople.mockResolvedValue([]);
    mockedApi.searchChats.mockResolvedValue([]);
    mockedApi.fetchProfiles.mockResolvedValue([]);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findOneOnOneConversation("Nonexistent Person");

    expect(result).toBeNull();
  });

  it("should fall back to profile-based matching when substrate search returns empty", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({
        id: "19:aaa11111-1111-1111-1111-111111111111_bbb22222-2222-2222-2222-222222222222@unq.gbl.spaces",
        threadType: "chat",
        topic: "",
      }),
    ]);

    // Substrate returns nothing
    mockedApi.searchPeople.mockResolvedValue([]);
    mockedApi.searchChats.mockResolvedValue([]);

    // Profile resolution resolves the UUIDs from conversation IDs
    mockedApi.fetchProfiles.mockResolvedValue([
      {
        mri: "8:orgid:aaa11111-1111-1111-1111-111111111111",
        displayName: "Current User",
        email: "me@example.com",
        jobTitle: "",
        userType: "Member",
      },
      {
        mri: "8:orgid:bbb22222-2222-2222-2222-222222222222",
        displayName: "Alice Smith",
        email: "alice@example.com",
        jobTitle: "Engineer",
        userType: "Member",
      },
    ]);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findOneOnOneConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.conversationId).toBe(
      "19:aaa11111-1111-1111-1111-111111111111_bbb22222-2222-2222-2222-222222222222@unq.gbl.spaces",
    );
    expect(result!.memberDisplayName).toBe("Alice Smith");
  });

  it("should fall back to profile-based matching when no substrate token", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({
        id: "19:aaa11111-1111-1111-1111-111111111111_bbb22222-2222-2222-2222-222222222222@unq.gbl.spaces",
        threadType: "chat",
        topic: "",
      }),
    ]);

    // searchPeople will throw when no substrate token
    mockedApi.searchPeople.mockRejectedValue(
      new ApiAuthError("Substrate token is missing"),
    );

    mockedApi.fetchProfiles.mockResolvedValue([
      {
        mri: "8:orgid:bbb22222-2222-2222-2222-222222222222",
        displayName: "Alice Smith",
        email: "alice@example.com",
        jobTitle: "Engineer",
        userType: "Member",
      },
    ]);

    // Bearer token but no substrate token
    const client = TeamsClient.fromToken("token", "apac", "bearer");
    const result = await client.findOneOnOneConversation("Alice");

    expect(result).not.toBeNull();
    expect(result!.conversationId).toBe(
      "19:aaa11111-1111-1111-1111-111111111111_bbb22222-2222-2222-2222-222222222222@unq.gbl.spaces",
    );
    expect(result!.memberDisplayName).toBe("Alice Smith");
  });

  it("should throw ApiAuthError when no match found via fallback and auth failed", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "19:chat1", threadType: "chat", topic: "" }),
    ]);

    // No substrate token — searchPeople throws
    mockedApi.searchPeople.mockRejectedValue(
      new ApiAuthError("Substrate token is missing"),
    );

    // fetchProfiles also throws (no bearer token)
    mockedApi.fetchProfiles.mockRejectedValue(
      new ApiAuthError("Bearer token is missing"),
    );

    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Someone Else" })]),
    );

    mockedApi.fetchUserProperties.mockResolvedValue({});

    const client = TeamsClient.fromToken("token");

    await expect(
      client.findOneOnOneConversation("Nonexistent Person"),
    ).rejects.toBeInstanceOf(ApiAuthError);
  });
});

describe("findPeople", () => {
  it("should return people from Substrate search", async () => {
    const people = [
      {
        displayName: "Alice Smith",
        mri: "8:orgid:alice-uuid",
        email: "alice@example.com",
        jobTitle: "Engineer",
        department: "Dev",
        objectId: "alice-uuid",
      },
    ];
    mockedApi.searchPeople.mockResolvedValue(people);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findPeople("Alice");

    expect(result).toEqual(people);
    expect(mockedApi.searchPeople).toHaveBeenCalledWith(
      expect.objectContaining({ substrateToken: "substrate-token" }),
      "Alice",
      10,
    );
  });

  it("should throw ApiAuthError when no substrate token and no bearer token", async () => {
    mockedApi.searchPeople.mockRejectedValue(
      new ApiAuthError("Substrate token is missing"),
    );
    mockedApi.fetchConversations.mockResolvedValue([]);

    const client = TeamsClient.fromToken("token");

    await expect(client.findPeople("Alice")).rejects.toBeInstanceOf(
      ApiAuthError,
    );
  });

  it("should fall back to conversation members when substrate returns empty", async () => {
    mockedApi.searchPeople.mockRejectedValue(
      new ApiAuthError("Substrate token is missing"),
    );

    // 1:1 chat → UUIDs extracted from conversation ID directly
    // Group chat → members fetched via fetchMembers
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({
        id: "19:aaa11111-1111-1111-1111-111111111111_bbb22222-2222-2222-2222-222222222222@unq.gbl.spaces",
        threadType: "chat",
        topic: "",
      }),
      makeConversation({ id: "19:group@thread.v2", threadType: "chat" }),
    ]);

    mockedApi.fetchMembers.mockResolvedValue([
      {
        id: "8:orgid:ccc33333-3333-3333-3333-333333333333",
        displayName: "",
        role: "User",
        memberType: "person" as const,
      },
    ]);

    mockedApi.fetchProfiles.mockResolvedValue([
      {
        mri: "8:orgid:aaa11111-1111-1111-1111-111111111111",
        displayName: "Current User",
        email: "me@example.com",
        jobTitle: "",
        userType: "Member",
      },
      {
        mri: "8:orgid:bbb22222-2222-2222-2222-222222222222",
        displayName: "Alice Smith",
        email: "alice@example.com",
        jobTitle: "Engineer",
        userType: "Member",
      },
      {
        mri: "8:orgid:ccc33333-3333-3333-3333-333333333333",
        displayName: "Bob Jones",
        email: "bob@example.com",
        jobTitle: "Designer",
        userType: "Member",
      },
    ]);

    const client = TeamsClient.fromToken("token", "apac", "bearer");
    const result = await client.findPeople("Alice");

    expect(result).toHaveLength(1);
    expect(result[0].displayName).toBe("Alice Smith");
    expect(result[0].email).toBe("alice@example.com");
  });
});

describe("findChats", () => {
  it("should return chats from Substrate search", async () => {
    const chats = [
      {
        name: "Design Team",
        threadId: "19:design@thread.v2",
        threadType: "Chat",
        matchingMembers: [
          { displayName: "Alice Smith", mri: "8:orgid:alice-uuid" },
        ],
        chatMembers: [],
        totalMemberCount: 4,
      },
    ];
    mockedApi.searchChats.mockResolvedValue(chats);

    const client = TeamsClient.fromToken(
      "token",
      "apac",
      "bearer",
      "substrate-token",
    );
    const result = await client.findChats("Design");

    expect(result).toEqual(chats);
    expect(mockedApi.searchChats).toHaveBeenCalledWith(
      expect.objectContaining({ substrateToken: "substrate-token" }),
      "Design",
      10,
    );
  });

  it("should fall back to local topic matching when substrate returns empty", async () => {
    mockedApi.searchChats.mockRejectedValue(
      new ApiAuthError("Substrate token is missing"),
    );

    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({
        id: "19:design@thread.v2",
        topic: "Design Team Chat",
        threadType: "topic",
        memberCount: 5,
      }),
      makeConversation({
        id: "19:other@thread.v2",
        topic: "Marketing Team",
        threadType: "topic",
        memberCount: 3,
      }),
    ]);

    const client = TeamsClient.fromToken("token");
    const result = await client.findChats("Design");

    expect(result).toHaveLength(1);
    expect(result[0].name).toBe("Design Team Chat");
    expect(result[0].threadId).toBe("19:design@thread.v2");
    expect(result[0].threadType).toBe("topic");
    expect(result[0].totalMemberCount).toBe(5);
  });

  it("should throw ApiAuthError when no substrate token and no topic matches", async () => {
    mockedApi.searchChats.mockRejectedValue(
      new ApiAuthError("Substrate token is missing"),
    );

    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({
        id: "19:other@thread.v2",
        topic: "Marketing Team",
        threadType: "topic",
      }),
    ]);

    const client = TeamsClient.fromToken("token");

    await expect(client.findChats("Design")).rejects.toBeInstanceOf(
      ApiAuthError,
    );
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

  it("should stop fetching once limit is reached", async () => {
    mockedApi.fetchMessagesPage
      .mockResolvedValueOnce(
        makeMessagesPage(
          [makeMessage({ id: "1" }), makeMessage({ id: "2" })],
          "https://next-page",
        ),
      )
      .mockResolvedValueOnce(
        makeMessagesPage(
          [makeMessage({ id: "3" }), makeMessage({ id: "4" })],
          "https://another-page",
        ),
      );

    const client = TeamsClient.fromToken("token");
    const messages = await client.getMessages("conv-id", { limit: 3 });

    expect(messages).toHaveLength(3);
    expect(messages.map((message) => message.id)).toEqual(["1", "2", "3"]);
    // Should not fetch a third page since limit was reached after page 2
    expect(mockedApi.fetchMessagesPage).toHaveBeenCalledTimes(2);
  });

  it("should return all messages when limit exceeds available", async () => {
    mockedApi.fetchMessagesPage.mockResolvedValueOnce(
      makeMessagesPage([makeMessage({ id: "1" })], null),
    );

    const client = TeamsClient.fromToken("token");
    const messages = await client.getMessages("conv-id", { limit: 100 });

    expect(messages).toHaveLength(1);
  });
});

describe("sendMessage", () => {
  it("should resolve display name and send with default markdown format", async () => {
    // getCurrentUserDisplayName will call fetchConversations + fetchMessagesPage
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );

    const expectedResult: SentMessage = {
      messageId: "1773000000000",
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
      "markdown",
      [],
    );
  });

  it("should pass explicit format to postMessage", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );
    mockedApi.postMessage.mockResolvedValue({
      messageId: "msg-2",
      arrivalTime: 1773000000000,
    });

    const client = TeamsClient.fromToken("token");
    await client.sendMessage("conv-id", "<b>Bold</b>", "html");

    expect(mockedApi.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({ skypeToken: "token" }),
      "conv-id",
      "<b>Bold</b>",
      "Test User",
      "html",
      [],
    );
  });
});

describe("scheduleMessage", () => {
  it("should resolve display name and schedule with default markdown format", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );

    const expectedResult: ScheduledMessage = {
      messageId: "1753021800000",
      arrivalTime: 1753021800000,
      scheduledTime: "2025-07-20T14:30:00.000Z",
    };
    mockedApi.postScheduledMessage.mockResolvedValue(expectedResult);

    const client = TeamsClient.fromToken("token");
    const scheduleAt = new Date("2025-07-20T14:30:00.000Z");
    const result = await client.scheduleMessage("conv-id", "Hello later!", scheduleAt);

    expect(result).toEqual(expectedResult);
    expect(mockedApi.postScheduledMessage).toHaveBeenCalledWith(
      expect.objectContaining({ skypeToken: "token" }),
      "conv-id",
      "Hello later!",
      "Test User",
      scheduleAt,
      "markdown",
      [],
    );
  });

  it("should pass explicit format to postScheduledMessage", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );
    mockedApi.postScheduledMessage.mockResolvedValue({
      messageId: "msg-2",
      arrivalTime: 1753021800000,
      scheduledTime: "2025-07-20T14:30:00.000Z",
    });

    const client = TeamsClient.fromToken("token");
    await client.scheduleMessage(
      "conv-id",
      "<b>Bold</b>",
      new Date("2025-07-20T14:30:00.000Z"),
      "html",
    );

    expect(mockedApi.postScheduledMessage).toHaveBeenCalledWith(
      expect.objectContaining({ skypeToken: "token" }),
      "conv-id",
      "<b>Bold</b>",
      "Test User",
      expect.any(Date),
      "html",
      [],
    );
  });
});

describe("getMembers", () => {
  describe("with bearerToken (profile API)", () => {
    it("should resolve display names via fetchProfiles", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person" as const,
        },
        {
          id: "8:orgid:user2",
          displayName: "",
          role: "User",
          memberType: "person" as const,
        },
      ]);
      mockedApi.fetchProfiles.mockResolvedValue([
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
      ]);

      const client = TeamsClient.fromToken("token", "apac", "bearer-token");
      const members = await client.getMembers("conv-id");

      expect(members).toEqual([
        {
          id: "8:orgid:user1",
          displayName: "Alice Smith",
          role: "Admin",
          memberType: "person",
        },
        {
          id: "8:orgid:user2",
          displayName: "Bob Jones",
          role: "User",
          memberType: "person",
        },
      ]);
      expect(mockedApi.fetchProfiles).toHaveBeenCalledOnce();
      expect(mockedApi.fetchMessagesPage).not.toHaveBeenCalled();
    });

    it("should fall back to message history when fetchProfiles fails", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person" as const,
        },
      ]);
      mockedApi.fetchProfiles.mockRejectedValue(new Error("profile API down"));
      mockedApi.fetchMessagesPage.mockResolvedValue(
        makeMessagesPage([
          makeMessage({
            senderMri:
              "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:user1",
            senderDisplayName: "Alice",
          }),
        ]),
      );

      const client = TeamsClient.fromToken("token", "apac", "bearer-token");
      const members = await client.getMembers("conv-id");

      expect(members[0].displayName).toBe("Alice");
      expect(mockedApi.fetchProfiles).toHaveBeenCalledOnce();
      expect(mockedApi.fetchMessagesPage).toHaveBeenCalled();
    });

    it("should only send unresolved person MRIs to fetchProfiles", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "Already Named",
          role: "Admin",
          memberType: "person" as const,
        },
        {
          id: "8:orgid:user2",
          displayName: "",
          role: "User",
          memberType: "person" as const,
        },
        {
          id: "28:bot-id",
          displayName: "",
          role: "User",
          memberType: "bot" as const,
        },
      ]);
      mockedApi.fetchProfiles.mockResolvedValue([
        {
          mri: "8:orgid:user2",
          displayName: "Bob",
          email: "",
          jobTitle: "",
          userType: "",
        },
      ]);

      const client = TeamsClient.fromToken("token", "apac", "bearer-token");
      const members = await client.getMembers("conv-id");

      expect(members[0].displayName).toBe("Already Named");
      expect(members[1].displayName).toBe("Bob");
      expect(members[2].displayName).toBe("");
      // Only user2 should have been sent to fetchProfiles
      expect(mockedApi.fetchProfiles).toHaveBeenCalledWith(expect.anything(), [
        "8:orgid:user2",
      ]);
    });
  });

  describe("without bearerToken (message history fallback)", () => {
    it("should resolve display names from message history", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person" as const,
        },
        {
          id: "8:orgid:user2",
          displayName: "",
          role: "User",
          memberType: "person" as const,
        },
      ]);
      mockedApi.fetchMessagesPage.mockResolvedValue(
        makeMessagesPage([
          makeMessage({
            senderMri:
              "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:user1",
            senderDisplayName: "Alice",
          }),
          makeMessage({
            senderMri:
              "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:user2",
            senderDisplayName: "Bob",
          }),
        ]),
      );

      const client = TeamsClient.fromToken("token");
      const members = await client.getMembers("conv-id");

      expect(members).toEqual([
        {
          id: "8:orgid:user1",
          displayName: "Alice",
          role: "Admin",
          memberType: "person",
        },
        {
          id: "8:orgid:user2",
          displayName: "Bob",
          role: "User",
          memberType: "person",
        },
      ]);
    });

    it("should paginate through messages to resolve all names", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person" as const,
        },
        {
          id: "8:orgid:user2",
          displayName: "",
          role: "User",
          memberType: "person" as const,
        },
      ]);
      mockedApi.fetchMessagesPage.mockResolvedValueOnce(
        makeMessagesPage(
          [
            makeMessage({
              senderMri:
                "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:user1",
              senderDisplayName: "Alice",
            }),
          ],
          "https://backward-link-page2",
        ),
      );
      mockedApi.fetchMessagesPage.mockResolvedValueOnce(
        makeMessagesPage([
          makeMessage({
            senderMri:
              "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:user2",
            senderDisplayName: "Bob",
          }),
        ]),
      );

      const client = TeamsClient.fromToken("token");
      const members = await client.getMembers("conv-id");

      expect(members[0].displayName).toBe("Alice");
      expect(members[1].displayName).toBe("Bob");
      expect(mockedApi.fetchMessagesPage).toHaveBeenCalledTimes(2);
    });

    it("should stop paginating once all people are resolved", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person" as const,
        },
      ]);
      mockedApi.fetchMessagesPage.mockResolvedValueOnce(
        makeMessagesPage(
          [
            makeMessage({
              senderMri: "8:orgid:user1",
              senderDisplayName: "Alice",
            }),
          ],
          "https://backward-link-page2",
        ),
      );

      const client = TeamsClient.fromToken("token");
      const members = await client.getMembers("conv-id");

      expect(members[0].displayName).toBe("Alice");
      expect(mockedApi.fetchMessagesPage).toHaveBeenCalledTimes(1);
    });

    it("should handle bare MRI senderMri values", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person" as const,
        },
      ]);
      mockedApi.fetchMessagesPage.mockResolvedValue(
        makeMessagesPage([
          makeMessage({
            senderMri: "8:orgid:user1",
            senderDisplayName: "Alice",
          }),
        ]),
      );

      const client = TeamsClient.fromToken("token");
      const members = await client.getMembers("conv-id");

      expect(members[0].displayName).toBe("Alice");
    });

    it("should still return members when message fetch fails", async () => {
      mockedApi.fetchMembers.mockResolvedValue([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person" as const,
        },
      ]);
      mockedApi.fetchMessagesPage.mockRejectedValue(new Error("forbidden"));

      const client = TeamsClient.fromToken("token");
      const members = await client.getMembers("conv-id");

      expect(members).toEqual([
        {
          id: "8:orgid:user1",
          displayName: "",
          role: "Admin",
          memberType: "person",
        },
      ]);
    });
  });

  it("should skip name resolution when all members already have names", async () => {
    mockedApi.fetchMembers.mockResolvedValue([
      {
        id: "8:orgid:user1",
        displayName: "Already Named",
        role: "Admin",
        memberType: "person" as const,
      },
    ]);

    const client = TeamsClient.fromToken("token");
    const members = await client.getMembers("conv-id");

    expect(members[0].displayName).toBe("Already Named");
    expect(mockedApi.fetchProfiles).not.toHaveBeenCalled();
    expect(mockedApi.fetchMessagesPage).not.toHaveBeenCalled();
  });

  it("should skip name resolution for bot-only unresolved members", async () => {
    mockedApi.fetchMembers.mockResolvedValue([
      {
        id: "8:orgid:user1",
        displayName: "Alice",
        role: "Admin",
        memberType: "person" as const,
      },
      {
        id: "28:bot-id",
        displayName: "",
        role: "User",
        memberType: "bot" as const,
      },
    ]);

    const client = TeamsClient.fromToken("token");
    const members = await client.getMembers("conv-id");

    expect(members[1].displayName).toBe("");
    expect(mockedApi.fetchProfiles).not.toHaveBeenCalled();
    expect(mockedApi.fetchMessagesPage).not.toHaveBeenCalled();
  });

  it("should leave display name empty when no messages match", async () => {
    mockedApi.fetchMembers.mockResolvedValue([
      {
        id: "8:orgid:user1",
        displayName: "",
        role: "Admin",
        memberType: "person" as const,
      },
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(makeMessagesPage([]));

    const client = TeamsClient.fromToken("token");
    const members = await client.getMembers("conv-id");

    expect(members[0].displayName).toBe("");
  });
});

describe("getCurrentUserDisplayName", () => {
  it("should get name from self-chat messages", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Alice Smith" })]),
    );

    const client = TeamsClient.fromToken("token");
    const name = await client.getCurrentUserDisplayName();

    expect(name).toBe("Alice Smith");
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

  it("should fallback to userDetails JSON in user properties", async () => {
    mockedApi.fetchConversations.mockResolvedValue([]);
    mockedApi.fetchUserProperties.mockResolvedValue({
      userDetails: JSON.stringify({ name: "Alice Smith" }),
    });

    const client = TeamsClient.fromToken("token");
    const name = await client.getCurrentUserDisplayName();

    expect(name).toBe("Alice Smith");
  });

  it("should prefer displayname over userDetails", async () => {
    mockedApi.fetchConversations.mockResolvedValue([]);
    mockedApi.fetchUserProperties.mockResolvedValue({
      displayname: "From Properties",
      userDetails: JSON.stringify({ name: "From UserDetails" }),
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

describe("TeamsClient.create", () => {
  it("should use cached token when available", async () => {
    const cachedToken = {
      skypeToken: "cached-token",
      region: "apac",
      substrateToken: "substrate-token",
      bearerToken: "bearer-token",
      amsToken: "ams-token",
      sharePointHost: "company-my.sharepoint.com",
    };
    mockedTokenStore.loadToken.mockReturnValue(cachedToken);

    const client = await TeamsClient.create({
      email: "user@company.com",
    });

    expect(mockedTokenStore.loadToken).toHaveBeenCalledWith("user@company.com");
    expect(mockedAuth.acquireTokenViaAutoLogin).not.toHaveBeenCalled();
    expect(client.getToken()).toEqual(cachedToken);
  });

  it("should apply an explicit region override to a cached token", async () => {
    const cachedToken = {
      skypeToken: "cached-token",
      region: "apac",
      substrateToken: "substrate-token",
      bearerToken: "bearer-token",
      amsToken: "ams-token",
      sharePointHost: "company-my.sharepoint.com",
    };
    mockedTokenStore.loadToken.mockReturnValue(cachedToken);

    const client = await TeamsClient.create({
      email: "user@company.com",
      region: "emea",
    });

    expect(mockedAuth.acquireTokenViaAutoLogin).not.toHaveBeenCalled();
    expect(mockedTokenStore.saveToken).toHaveBeenCalledWith(
      "user@company.com",
      { ...cachedToken, region: "emea" },
    );
    expect(client.getToken()).toEqual({
      ...cachedToken,
      region: "emea",
    });
  });

  it("should auto-login and save token when no cache exists", async () => {
    const freshToken = { skypeToken: "fresh-token", region: "apac" };
    mockedTokenStore.loadToken.mockReturnValue(null);
    mockedAuth.acquireTokenViaAutoLogin.mockResolvedValue(freshToken);

    const client = await TeamsClient.create({
      email: "user@company.com",
    });

    expect(mockedAuth.acquireTokenViaAutoLogin).toHaveBeenCalledWith(
      expect.objectContaining({ email: "user@company.com" }),
    );
    expect(mockedTokenStore.saveToken).toHaveBeenCalledWith(
      "user@company.com",
      freshToken,
    );
    expect(client.getToken()).toEqual(freshToken);
  });

  it("should re-authenticate when cached token is missing substrate token", async () => {
    const incompleteToken = { skypeToken: "cached-token", region: "apac" };
    const freshToken = {
      skypeToken: "fresh-token",
      region: "apac",
      substrateToken: "substrate-token",
      bearerToken: "bearer-token",
      amsToken: "ams-token",
    };
    mockedTokenStore.loadToken.mockReturnValue(incompleteToken);
    mockedAuth.acquireTokenViaAutoLogin.mockResolvedValue(freshToken);

    const client = await TeamsClient.create({
      email: "user@company.com",
    });

    expect(mockedTokenStore.clearToken).toHaveBeenCalledWith(
      "user@company.com",
    );
    expect(mockedAuth.acquireTokenViaAutoLogin).toHaveBeenCalledWith(
      expect.objectContaining({ email: "user@company.com" }),
    );
    expect(mockedTokenStore.saveToken).toHaveBeenCalledWith(
      "user@company.com",
      freshToken,
    );
    expect(client.getToken()).toEqual(freshToken);
  });

  it("should re-authenticate when cached token is missing bearer token", async () => {
    const incompleteToken = {
      skypeToken: "cached-token",
      region: "apac",
      substrateToken: "substrate-token",
    };
    const freshToken = {
      skypeToken: "fresh-token",
      region: "apac",
      substrateToken: "substrate-token",
      bearerToken: "bearer-token",
      amsToken: "ams-token",
    };
    mockedTokenStore.loadToken.mockReturnValue(incompleteToken);
    mockedAuth.acquireTokenViaAutoLogin.mockResolvedValue(freshToken);

    const client = await TeamsClient.create({
      email: "user@company.com",
    });

    expect(mockedTokenStore.clearToken).toHaveBeenCalledWith(
      "user@company.com",
    );
    expect(mockedAuth.acquireTokenViaAutoLogin).toHaveBeenCalled();
    expect(client.getToken()).toEqual(freshToken);
  });
});

describe("TeamsClient.clearCachedToken", () => {
  it("should delegate to token store", () => {
    TeamsClient.clearCachedToken("user@company.com");

    expect(mockedTokenStore.clearToken).toHaveBeenCalledWith(
      "user@company.com",
    );
  });
});

describe("withTokenRefresh (automatic 401 retry)", () => {
  it("should refresh token and retry on ApiAuthError", async () => {
    const initialToken = {
      skypeToken: "old-token",
      region: "apac",
      substrateToken: "substrate-token",
      bearerToken: "bearer-token",
      amsToken: "ams-token",
    };
    const refreshedToken = { skypeToken: "new-token", region: "apac" };
    mockedTokenStore.loadToken.mockReturnValue(initialToken);

    const client = await TeamsClient.create({
      email: "user@company.com",
    });

    // First call to fetchConversations fails with 401
    mockedApi.fetchConversations
      .mockRejectedValueOnce(new ApiAuthError("Authentication failed: 401"))
      .mockResolvedValueOnce([makeConversation({ topic: "After Refresh" })]);

    // The refresh will call acquireTokenViaAutoLogin
    mockedAuth.acquireTokenViaAutoLogin.mockResolvedValue(refreshedToken);

    const conversations = await client.listConversations();

    expect(conversations).toHaveLength(1);
    expect(conversations[0].topic).toBe("After Refresh");

    // Verify refresh happened: clear old token, acquire new, save new
    expect(mockedTokenStore.clearToken).toHaveBeenCalledWith(
      "user@company.com",
    );
    expect(mockedTokenStore.saveToken).toHaveBeenCalledWith(
      "user@company.com",
      refreshedToken,
    );
  });

  it("should not retry on non-auth errors", async () => {
    mockedTokenStore.loadToken.mockReturnValue({
      skypeToken: "token",
      region: "apac",
      substrateToken: "substrate-token",
      bearerToken: "bearer-token",
      amsToken: "ams-token",
      sharePointHost: "company-my.sharepoint.com",
    });

    const client = await TeamsClient.create({
      email: "user@company.com",
    });

    mockedApi.fetchConversations.mockRejectedValueOnce(
      new Error("Network error"),
    );

    await expect(client.listConversations()).rejects.toThrow("Network error");

    // Should not have tried to refresh
    expect(mockedTokenStore.clearToken).not.toHaveBeenCalled();
  });

  it("should not retry when created via fromToken (no auto-login options)", async () => {
    const client = TeamsClient.fromToken("manual-token");

    mockedApi.fetchConversations.mockRejectedValueOnce(
      new ApiAuthError("Authentication failed: 401"),
    );

    await expect(client.listConversations()).rejects.toThrow(
      "Authentication failed: 401",
    );

    // No refresh attempt since fromToken doesn't have auto-login options
    expect(mockedAuth.acquireTokenViaAutoLogin).not.toHaveBeenCalled();
  });

  it("should propagate error if retry also fails", async () => {
    const initialToken = {
      skypeToken: "token",
      region: "apac",
      substrateToken: "substrate-token",
      bearerToken: "bearer-token",
    };
    const refreshedToken = { skypeToken: "new-token", region: "apac" };
    mockedTokenStore.loadToken.mockReturnValue(initialToken);

    const client = await TeamsClient.create({
      email: "user@company.com",
    });

    // Both attempts fail with 401
    mockedApi.fetchConversations
      .mockRejectedValueOnce(new ApiAuthError("Authentication failed: 401"))
      .mockRejectedValueOnce(new ApiAuthError("Authentication failed: 401"));

    mockedAuth.acquireTokenViaAutoLogin.mockResolvedValue(refreshedToken);

    await expect(client.listConversations()).rejects.toThrow(
      "Authentication failed: 401",
    );
  });
});

describe("sendMessageWithFiles", () => {
  it("should throw when email is not set", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );

    const client = TeamsClient.fromToken("token");
    // Do NOT call setEmail

    await expect(
      client.sendMessageWithFiles("conv-id", [
        { type: "file", data: Buffer.from("hello"), fileName: "test.md" },
      ]),
    ).rejects.toThrow("User email is required for file upload");
  });

  it("should upload files to SharePoint and send message with properties.files", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );

    const uploadResult: attachments.SharePointUploadResult = {
      itemId: "sp-item-123",
      siteId: "site-456",
      fileName: "report.md",
      fileType: "md",
      fileUrl: "https://sp.com/report.md",
      webDavUrl: "https://sp.com/dav/report.md",
      siteBaseUrl: "https://company-my.sharepoint.com",
      personalPath: "/personal/user_company_com",
    };
    mockedAttachments.uploadSharePointFile.mockResolvedValue(uploadResult);
    mockedAttachments.buildFilesPropertyJson.mockReturnValue(
      '[{"@type":"http://schema.skype.com/File"}]',
    );

    const expectedResult: SentMessage = {
      messageId: "1773000000000",
      arrivalTime: 1773000000000,
    };
    mockedApi.postMessage.mockResolvedValue(expectedResult);

    const client = TeamsClient.fromToken("token");
    client.setEmail("user@company.com");

    const result = await client.sendMessageWithFiles("conv-id", [
      { type: "text", text: "Here is the file:" },
      { type: "file", data: Buffer.from("# Report"), fileName: "report.md" },
    ]);

    expect(result).toEqual(expectedResult);

    // Verify SharePoint upload was called
    expect(mockedAttachments.uploadSharePointFile).toHaveBeenCalledWith(
      expect.objectContaining({ skypeToken: "token" }),
      Buffer.from("# Report"),
      "report.md",
      "user@company.com",
    );

    // Verify buildFilesPropertyJson was called with upload results
    expect(mockedAttachments.buildFilesPropertyJson).toHaveBeenCalledWith([
      uploadResult,
    ]);

    // Verify postMessage was called with files JSON
    expect(mockedApi.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({ skypeToken: "token" }),
      "conv-id",
      "<div><p>Here is the file:</p></div>",
      "Test User",
      "html",
      [],
      '[{"@type":"http://schema.skype.com/File"}]',
    );
  });

  it("should handle file-only messages with no text content", async () => {
    mockedApi.fetchConversations.mockResolvedValue([
      makeConversation({ id: "48:notes" }),
    ]);
    mockedApi.fetchMessagesPage.mockResolvedValue(
      makeMessagesPage([makeMessage({ senderDisplayName: "Test User" })]),
    );

    mockedAttachments.uploadSharePointFile.mockResolvedValue({
      itemId: "item-1",
      siteId: "site-1",
      fileName: "data.csv",
      fileType: "csv",
      fileUrl: "https://sp.com/data.csv",
      webDavUrl: "https://sp.com/dav/data.csv",
      siteBaseUrl: "https://sp.com",
      personalPath: "/personal/user",
    });
    mockedAttachments.buildFilesPropertyJson.mockReturnValue("[]");
    mockedApi.postMessage.mockResolvedValue({
      messageId: "msg-1",
      arrivalTime: 1773000000000,
    });

    const client = TeamsClient.fromToken("token");
    client.setEmail("user@company.com");

    await client.sendMessageWithFiles("conv-id", [
      { type: "file", data: Buffer.from("a,b,c"), fileName: "data.csv" },
    ]);

    // No text content: should send empty string
    expect(mockedApi.postMessage).toHaveBeenCalledWith(
      expect.anything(),
      "conv-id",
      "",
      "Test User",
      "html",
      [],
      expect.any(String),
    );
  });
});
