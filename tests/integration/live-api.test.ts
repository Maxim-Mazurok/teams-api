/**
 * Integration tests — hit the real Teams API.
 *
 * These tests are skipped by default.
 * Run them with: TEAMS_TOKEN=<token> TEAMS_REGION=<apac|emea|amer> npx vitest run tests/integration
 *
 * Optional env vars for richer tests:
 *   - TEAMS_BEARER_TOKEN  — enables profile resolution and search tests
 *   - TEAMS_SUBSTRATE_TOKEN — enables people/chat search tests
 *
 * Prerequisites:
 *   - A valid skype token (acquire via `npx tsx src/cli.ts auth --auto --email you@company.com`)
 *   - A Teams account with at least one conversation
 */

import { describe, it, expect } from "vitest";
import { TeamsClient } from "../../src/teams-client.js";

const skypeToken = process.env["TEAMS_TOKEN"];
const region = process.env["TEAMS_REGION"];
const bearerToken = process.env["TEAMS_BEARER_TOKEN"];
const substrateToken = process.env["TEAMS_SUBSTRATE_TOKEN"];

const shouldRun = Boolean(skypeToken && region);
const hasSearchTokens = Boolean(bearerToken && substrateToken);

describe.skipIf(!shouldRun)("Live Teams API", () => {
  function createClient(): TeamsClient {
    if (!skypeToken) throw new Error("TEAMS_TOKEN is required");
    if (!region) throw new Error("TEAMS_REGION is required");
    return TeamsClient.fromToken(
      skypeToken,
      region,
      bearerToken,
      substrateToken,
    );
  }

  // ── Conversation operations ──────────────────────────────────────

  it("should list conversations", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 10 });

    expect(conversations.length).toBeGreaterThan(0);
    for (const conversation of conversations) {
      expect(typeof conversation.id).toBe("string");
      expect(typeof conversation.threadType).toBe("string");
    }
  });

  it("should list conversations with excludeSystemStreams disabled", async () => {
    const client = createClient();
    const conversations = await client.listConversations({
      pageSize: 50,
      excludeSystemStreams: false,
    });

    expect(conversations.length).toBeGreaterThan(0);
  });

  it("should find a conversation by topic", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 10 });
    const topicConversation = conversations.find(
      (conversation) => conversation.topic,
    );

    if (!topicConversation) return; // skip if no named conversations

    const found = await client.findConversation(topicConversation.topic);
    expect(found).not.toBeNull();
    expect(found!.id).toBe(topicConversation.id);
  });

  // ── Message operations ───────────────────────────────────────────

  it("should get messages from the first conversation", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 5 });
    expect(conversations.length).toBeGreaterThan(0);

    const messages = await client.getMessages(conversations[0].id, {
      maxPages: 1,
      pageSize: 10,
    });

    expect(messages.length).toBeGreaterThan(0);
    for (const message of messages) {
      expect(typeof message.id).toBe("string");
      expect(typeof message.originalArrivalTime).toBe("string");
    }
  });

  it("should get messages with a page limit", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 5 });
    expect(conversations.length).toBeGreaterThan(0);

    const messages = await client.getMessages(conversations[0].id, {
      limit: 5,
    });

    expect(messages.length).toBeLessThanOrEqual(5);
  });

  it("should send, edit, and delete a message in self-chat", async () => {
    const client = createClient();
    const displayName = await client.getCurrentUserDisplayName();
    const selfChat = await client.findOneOnOneConversation(displayName);

    if (!selfChat) return; // skip if self-chat not found

    // Send
    const uniqueContent = `Integration test message ${Date.now()}`;
    const sentResult = await client.sendMessage(
      selfChat.conversationId,
      uniqueContent,
    );
    expect(typeof sentResult.messageId).toBe("string");
    expect(typeof sentResult.arrivalTime).toBe("number");

    // Verify message appears
    const messagesAfterSend = await client.getMessages(
      selfChat.conversationId,
      { maxPages: 1, pageSize: 20 },
    );
    const sentMessage = messagesAfterSend.find((message) =>
      message.content.includes(uniqueContent),
    );
    expect(sentMessage).toBeDefined();

    // Edit
    const editedContent = `${uniqueContent} (edited)`;
    const editResult = await client.editMessage(
      selfChat.conversationId,
      sentResult.messageId,
      editedContent,
    );
    expect(typeof editResult.editTime).toBe("string");

    // Delete
    const deleteResult = await client.deleteMessage(
      selfChat.conversationId,
      sentResult.messageId,
    );
    expect(typeof deleteResult.messageId).toBe("string");
  });

  // ── Member operations ────────────────────────────────────────────

  it("should resolve current user display name", async () => {
    const client = createClient();
    const displayName = await client.getCurrentUserDisplayName();

    expect(typeof displayName).toBe("string");
    expect(displayName).not.toBe("Unknown User");
  });

  it("should get members of a conversation", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 5 });
    expect(conversations.length).toBeGreaterThan(0);

    const members = await client.getMembers(conversations[0].id);

    expect(members.length).toBeGreaterThan(0);
    for (const member of members) {
      expect(typeof member.id).toBe("string");
    }
  });

  // ── 1:1 conversation lookup ──────────────────────────────────────

  it("should find self-chat by display name", async () => {
    const client = createClient();
    const displayName = await client.getCurrentUserDisplayName();
    const selfChat = await client.findOneOnOneConversation(displayName);

    expect(selfChat).not.toBeNull();
    expect(typeof selfChat!.conversationId).toBe("string");
    expect(selfChat!.memberDisplayName).toBe(displayName);
  });

  // ── Search operations (require bearer + substrate tokens) ────────

  describe.skipIf(!hasSearchTokens)("Search API", () => {
    it("should find people by name", async () => {
      const client = createClient();
      const displayName = await client.getCurrentUserDisplayName();
      // Search for the current user — guaranteed to exist
      const people = await client.findPeople(displayName, 5);

      expect(people.length).toBeGreaterThan(0);
      expect(typeof people[0].displayName).toBe("string");
      expect(typeof people[0].email).toBe("string");
      expect(typeof people[0].mri).toBe("string");
    });

    it("should find chats by name", async () => {
      const client = createClient();
      const conversations = await client.listConversations({ pageSize: 10 });
      const topicConversation = conversations.find(
        (conversation) => conversation.topic,
      );

      if (!topicConversation) return; // skip if no named conversations

      const chats = await client.findChats(topicConversation.topic, 5);
      // Substrate search may return results depending on indexing
      expect(Array.isArray(chats)).toBe(true);
    });
  });
});
