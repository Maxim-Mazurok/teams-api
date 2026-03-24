import { describe, it, expect, vi } from "vitest";
import {
  conversationParameters,
  resolveConversationId,
} from "../../src/actions/conversation-resolution.js";
import type { TeamsClient } from "../../src/teams-client.js";

// ── Mock helpers ─────────────────────────────────────────────────────

function createMockClient(overrides?: {
  findConversation?: TeamsClient["findConversation"];
  findOneOnOneConversation?: TeamsClient["findOneOnOneConversation"];
}): TeamsClient {
  return {
    findConversation:
      overrides?.findConversation ?? vi.fn().mockResolvedValue(null),
    findOneOnOneConversation:
      overrides?.findOneOnOneConversation ?? vi.fn().mockResolvedValue(null),
  } as unknown as TeamsClient;
}

// ── conversationParameters ──────────────────────────────────────────

describe("conversationParameters", () => {
  it("contains exactly 3 parameter definitions", () => {
    expect(conversationParameters).toHaveLength(3);
  });

  it("defines chat, to, and conversationId parameters", () => {
    const names = conversationParameters.map((parameter) => parameter.name);
    expect(names).toContain("chat");
    expect(names).toContain("to");
    expect(names).toContain("conversationId");
  });

  it("marks all parameters as not required", () => {
    for (const parameter of conversationParameters) {
      expect(parameter.required).toBe(false);
    }
  });
});

// ── resolveConversationId ───────────────────────────────────────────

describe("resolveConversationId", () => {
  it("uses conversationId directly when provided", async () => {
    const client = createMockClient();
    const result = await resolveConversationId(client, {
      conversationId: "19:abc123@thread.v2",
    });

    expect(result.conversationId).toBe("19:abc123@thread.v2");
    expect(result.label).toBe("19:abc123@thread.v2");
  });

  it("uses chat as raw ID when it starts with 19: and contains @", async () => {
    const client = createMockClient();
    const result = await resolveConversationId(client, {
      chat: "19:meeting_abc@thread.v2",
    });

    expect(result.conversationId).toBe("19:meeting_abc@thread.v2");
    expect(result.label).toBe("19:meeting_abc@thread.v2");
    expect(client.findConversation).not.toHaveBeenCalled();
  });

  it("resolves chat via findConversation when topic matches", async () => {
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue({
        id: "19:resolved@thread.v2",
        topic: "Engineering Chat",
      }),
    });

    const result = await resolveConversationId(client, { chat: "Engineering" });

    expect(result.conversationId).toBe("19:resolved@thread.v2");
    expect(result.label).toBe("Engineering Chat");
  });

  it("falls back to findOneOnOneConversation when findConversation returns null for chat", async () => {
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(null),
      findOneOnOneConversation: vi.fn().mockResolvedValue({
        conversationId: "19:oneononeid@thread.v2",
        memberDisplayName: "John Smith",
      }),
    });

    const result = await resolveConversationId(client, { chat: "John" });

    expect(client.findConversation).toHaveBeenCalledWith("John");
    expect(client.findOneOnOneConversation).toHaveBeenCalledWith("John");
    expect(result.conversationId).toBe("19:oneononeid@thread.v2");
    expect(result.label).toBe("John Smith");
  });

  it("throws when chat matches no conversation and no 1:1", async () => {
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(null),
      findOneOnOneConversation: vi.fn().mockResolvedValue(null),
    });

    await expect(
      resolveConversationId(client, { chat: "nonexistent" }),
    ).rejects.toThrow('No conversation matching "nonexistent" found.');
  });

  it("resolves to param via findOneOnOneConversation", async () => {
    const client = createMockClient({
      findOneOnOneConversation: vi.fn().mockResolvedValue({
        conversationId: "19:dm@thread.v2",
        memberDisplayName: "Jane Doe",
      }),
    });

    const result = await resolveConversationId(client, { to: "Jane" });

    expect(client.findOneOnOneConversation).toHaveBeenCalledWith("Jane");
    expect(result.conversationId).toBe("19:dm@thread.v2");
    expect(result.label).toBe("Jane Doe");
  });

  it("throws when to param finds no 1:1 conversation", async () => {
    const client = createMockClient({
      findOneOnOneConversation: vi.fn().mockResolvedValue(null),
    });

    await expect(
      resolveConversationId(client, { to: "Nobody" }),
    ).rejects.toThrow('No 1:1 conversation found with "Nobody".');
  });

  it("throws when no identification parameters are provided", async () => {
    const client = createMockClient();

    await expect(resolveConversationId(client, {})).rejects.toThrow(
      "One of --conversation-id, --chat, or --to is required.",
    );
  });

  it("prefers conversationId over chat and to", async () => {
    const client = createMockClient();
    const result = await resolveConversationId(client, {
      conversationId: "19:direct@thread.v2",
      chat: "SomeTopic",
      to: "SomePerson",
    });

    expect(result.conversationId).toBe("19:direct@thread.v2");
    expect(client.findConversation).not.toHaveBeenCalled();
    expect(client.findOneOnOneConversation).not.toHaveBeenCalled();
  });

  it("prefers chat over to when conversationId is absent", async () => {
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue({
        id: "19:chatfound@thread.v2",
        topic: "Team Chat",
      }),
    });

    const result = await resolveConversationId(client, {
      chat: "Team",
      to: "SomePerson",
    });

    expect(result.conversationId).toBe("19:chatfound@thread.v2");
    expect(client.findOneOnOneConversation).not.toHaveBeenCalled();
  });
});
