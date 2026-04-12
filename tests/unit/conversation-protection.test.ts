/**
 * Unit tests for conversation protection (src/conversation-protection.ts).
 *
 * Tests pattern resolution, glob matching, and integration with
 * edit/delete actions.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import {
  resolveProtectedPatterns,
  matchProtectedConversation,
} from "../../src/conversation-protection.js";
import { actions } from "../../src/actions/definitions.js";
import type { TeamsClient } from "../../src/teams-client.js";
import type {
  Conversation,
  EditedMessage,
  DeletedMessage,
} from "../../src/types.js";

// ── Helpers ──────────────────────────────────────────────────────────

function getAction(name: string) {
  const action = actions.find((a) => a.name === name);
  if (!action) throw new Error(`Action "${name}" not found`);
  return action;
}

function makeConversation(overrides: Partial<Conversation> = {}): Conversation {
  return {
    id: "19:test@thread.space",
    topic: "Architecture Decisions",
    threadType: "chat",
    version: 1,
    lastMessageTime: "2026-03-16T10:00:00.000Z",
    memberCount: 5,
    ...overrides,
  };
}

function createMockClient(
  overrides: Partial<Record<keyof TeamsClient, unknown>> = {},
): TeamsClient {
  return {
    listConversations: vi.fn(),
    findConversation: vi.fn(),
    findOneOnOneConversation: vi.fn(),
    findPeople: vi.fn(),
    findChats: vi.fn(),
    getMessages: vi.fn(),
    getMessagesPage: vi.fn(),
    sendMessage: vi.fn(),
    sendMessageWithImages: vi.fn(),
    sendMessageWithFiles: vi.fn(),
    editMessage: vi.fn(),
    checkForReplies: vi
      .fn()
      .mockResolvedValue({ hasReplies: false, replyCount: 0 }),
    deleteMessage: vi.fn(),
    addReaction: vi.fn(),
    removeReaction: vi.fn(),
    scheduleMessage: vi.fn(),
    getMembers: vi.fn(),
    getCurrentUserDisplayName: vi.fn(),
    getToken: vi.fn(() => ({ skypeToken: "test-token", region: "apac" })),
    setEmail: vi.fn(),
    ...overrides,
  } as unknown as TeamsClient;
}

// ── resolveProtectedPatterns ─────────────────────────────────────────

describe("resolveProtectedPatterns", () => {
  beforeEach(() => {
    delete process.env.TEAMS_PROTECTED_CONVERSATIONS;
  });

  afterEach(() => {
    delete process.env.TEAMS_PROTECTED_CONVERSATIONS;
  });

  it("returns empty array when no env var and no parameter", () => {
    expect(resolveProtectedPatterns()).toEqual([]);
  });

  it("returns empty array for empty string", () => {
    expect(resolveProtectedPatterns("")).toEqual([]);
  });

  it("parses a single pattern", () => {
    expect(resolveProtectedPatterns("Architecture *")).toEqual([
      "Architecture *",
    ]);
  });

  it("parses multiple comma-separated patterns", () => {
    expect(
      resolveProtectedPatterns("Architecture *,*compliance*,Incident *"),
    ).toEqual(["Architecture *", "*compliance*", "Incident *"]);
  });

  it("trims whitespace from patterns", () => {
    expect(
      resolveProtectedPatterns("  Architecture * , *compliance* , Incident * "),
    ).toEqual(["Architecture *", "*compliance*", "Incident *"]);
  });

  it("filters out empty segments", () => {
    expect(resolveProtectedPatterns("Pattern A,,Pattern B,")).toEqual([
      "Pattern A",
      "Pattern B",
    ]);
  });

  it("reads from TEAMS_PROTECTED_CONVERSATIONS env var when no parameter", () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Incident *,*audit*";
    expect(resolveProtectedPatterns()).toEqual(["Incident *", "*audit*"]);
  });

  it("parameter overrides env var", () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "from-env";
    expect(resolveProtectedPatterns("from-param")).toEqual(["from-param"]);
  });

  it("null parameter falls back to env var", () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "from-env";
    expect(resolveProtectedPatterns(null)).toEqual(["from-env"]);
  });
});

// ── matchProtectedConversation ───────────────────────────────────────

describe("matchProtectedConversation", () => {
  it("returns undefined when patterns array is empty", () => {
    expect(matchProtectedConversation("Any Chat", [])).toBeUndefined();
  });

  it("matches exact conversation name", () => {
    expect(
      matchProtectedConversation("Architecture Decisions", [
        "Architecture Decisions",
      ]),
    ).toBe("Architecture Decisions");
  });

  it("matches with leading wildcard", () => {
    expect(
      matchProtectedConversation("Q4 Compliance Review", ["*compliance*"]),
    ).toBe("*compliance*");
  });

  it("matches with trailing wildcard", () => {
    expect(
      matchProtectedConversation("Incident Response Team", ["Incident *"]),
    ).toBe("Incident *");
  });

  it("matches with wildcards on both sides", () => {
    expect(matchProtectedConversation("2026 Audit Results", ["*audit*"])).toBe(
      "*audit*",
    );
  });

  it("matching is case-insensitive", () => {
    expect(
      matchProtectedConversation("architecture decisions", ["Architecture *"]),
    ).toBe("Architecture *");
  });

  it("returns the first matching pattern", () => {
    expect(
      matchProtectedConversation("Architecture Audit", [
        "Architecture *",
        "*Audit*",
      ]),
    ).toBe("Architecture *");
  });

  it("returns undefined when no patterns match", () => {
    expect(
      matchProtectedConversation("General Chat", [
        "Architecture *",
        "*compliance*",
      ]),
    ).toBeUndefined();
  });

  it("handles special regex characters in conversation name", () => {
    expect(
      matchProtectedConversation("Team (2026) Compliance", ["Team (2026) *"]),
    ).toBe("Team (2026) *");
  });

  it("handles special regex characters in pattern", () => {
    expect(matchProtectedConversation("test.topic", ["test.topic"])).toBe(
      "test.topic",
    );
  });

  it("does not partially match", () => {
    expect(
      matchProtectedConversation("Architecture Decisions Extra", [
        "Architecture Decisions",
      ]),
    ).toBeUndefined();
  });
});

// ── edit-message integration ─────────────────────────────────────────

describe("edit-message conversation protection", () => {
  const editAction = getAction("edit-message");

  beforeEach(() => {
    delete process.env.TEAMS_PROTECTED_CONVERSATIONS;
    delete process.env.TEAMS_AUDIT_LOG;
    delete process.env.TEAMS_AGENT_MARKER;
    delete process.env.TEAMS_EDIT_REPLY_GUARD;
  });

  afterEach(() => {
    delete process.env.TEAMS_PROTECTED_CONVERSATIONS;
    delete process.env.TEAMS_AUDIT_LOG;
    delete process.env.TEAMS_AGENT_MARKER;
    delete process.env.TEAMS_EDIT_REPLY_GUARD;
  });

  it("blocks edit in a protected conversation (env var)", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Architecture *";
    const conversation = makeConversation({
      topic: "Architecture Decisions",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn<() => Promise<EditedMessage>>().mockResolvedValue({
        messageId: "msg-1",
        editTime: "2026-04-12T10:00:00Z",
      }),
    });

    await expect(
      editAction.execute(client, {
        chat: "Architecture Decisions",
        messageId: "msg-1",
        content: "Updated content",
      }),
    ).rejects.toThrow(/conversation is protected/);

    expect(client.editMessage).not.toHaveBeenCalled();
  });

  it("blocks edit in a protected conversation (parameter override)", async () => {
    const conversation = makeConversation({
      topic: "Architecture Decisions",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn<() => Promise<EditedMessage>>().mockResolvedValue({
        messageId: "msg-1",
        editTime: "2026-04-12T10:00:00Z",
      }),
    });

    await expect(
      editAction.execute(client, {
        chat: "Architecture Decisions",
        messageId: "msg-1",
        content: "Updated content",
        protectedConversations: "Architecture *",
      }),
    ).rejects.toThrow(/conversation is protected/);

    expect(client.editMessage).not.toHaveBeenCalled();
  });

  it("allows edit when conversation does not match protected patterns", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Architecture *";
    const conversation = makeConversation({
      topic: "General Chat",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn<() => Promise<EditedMessage>>().mockResolvedValue({
        messageId: "msg-1",
        editTime: "2026-04-12T10:00:00Z",
      }),
    });

    const result = await editAction.execute(client, {
      chat: "General Chat",
      messageId: "msg-1",
      content: "Updated content",
    });

    expect(client.editMessage).toHaveBeenCalled();
    expect(result).toHaveProperty("messageId", "msg-1");
  });

  it("allows edit when no protected patterns are configured", async () => {
    const conversation = makeConversation({
      topic: "Architecture Decisions",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn<() => Promise<EditedMessage>>().mockResolvedValue({
        messageId: "msg-1",
        editTime: "2026-04-12T10:00:00Z",
      }),
    });

    const result = await editAction.execute(client, {
      chat: "Architecture Decisions",
      messageId: "msg-1",
      content: "Updated content",
    });

    expect(client.editMessage).toHaveBeenCalled();
    expect(result).toHaveProperty("messageId", "msg-1");
  });

  it("error message includes conversation name and matched pattern", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "*compliance*";
    const conversation = makeConversation({
      topic: "Q4 Compliance Review",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
    });

    await expect(
      editAction.execute(client, {
        chat: "Q4 Compliance Review",
        messageId: "msg-1",
        content: "new",
      }),
    ).rejects.toThrow(
      /Cannot edit message in "Q4 Compliance Review".*matched pattern "\*compliance\*"/,
    );
  });
});

// ── delete-message integration ───────────────────────────────────────

describe("delete-message conversation protection", () => {
  const deleteAction = getAction("delete-message");

  beforeEach(() => {
    delete process.env.TEAMS_PROTECTED_CONVERSATIONS;
    delete process.env.TEAMS_AUDIT_LOG;
    delete process.env.TEAMS_DELETE_MODE;
    delete process.env.TEAMS_DELETE_TOMBSTONE;
  });

  afterEach(() => {
    delete process.env.TEAMS_PROTECTED_CONVERSATIONS;
    delete process.env.TEAMS_AUDIT_LOG;
    delete process.env.TEAMS_DELETE_MODE;
    delete process.env.TEAMS_DELETE_TOMBSTONE;
  });

  it("blocks hard delete in a protected conversation", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Incident *";
    const conversation = makeConversation({
      topic: "Incident Response 2026-04-12",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      deleteMessage: vi.fn<() => Promise<DeletedMessage>>().mockResolvedValue({
        messageId: "msg-1",
      }),
    });

    await expect(
      deleteAction.execute(client, {
        chat: "Incident Response 2026-04-12",
        messageId: "msg-1",
      }),
    ).rejects.toThrow(/conversation is protected/);

    expect(client.deleteMessage).not.toHaveBeenCalled();
  });

  it("blocks soft delete in a protected conversation", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Incident *";
    const conversation = makeConversation({
      topic: "Incident Response 2026-04-12",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn<() => Promise<EditedMessage>>().mockResolvedValue({
        messageId: "msg-1",
        editTime: "2026-04-12T10:00:00Z",
      }),
    });

    await expect(
      deleteAction.execute(client, {
        chat: "Incident Response 2026-04-12",
        messageId: "msg-1",
        deleteMode: "soft",
      }),
    ).rejects.toThrow(/conversation is protected/);

    expect(client.editMessage).not.toHaveBeenCalled();
  });

  it("allows delete when conversation does not match protected patterns", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Incident *";
    const conversation = makeConversation({
      topic: "General Chat",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      deleteMessage: vi.fn<() => Promise<DeletedMessage>>().mockResolvedValue({
        messageId: "msg-1",
      }),
    });

    const result = await deleteAction.execute(client, {
      chat: "General Chat",
      messageId: "msg-1",
    });

    expect(client.deleteMessage).toHaveBeenCalled();
    expect(result).toHaveProperty("messageId", "msg-1");
  });

  it("allows delete when no protected patterns are configured", async () => {
    const conversation = makeConversation({
      topic: "Incident Response 2026-04-12",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      deleteMessage: vi.fn<() => Promise<DeletedMessage>>().mockResolvedValue({
        messageId: "msg-1",
      }),
    });

    const result = await deleteAction.execute(client, {
      chat: "Incident Response 2026-04-12",
      messageId: "msg-1",
    });

    expect(client.deleteMessage).toHaveBeenCalled();
    expect(result).toHaveProperty("messageId", "msg-1");
  });

  it("error message includes conversation name and matched pattern", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Incident *";
    const conversation = makeConversation({
      topic: "Incident Response 2026-04-12",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
    });

    await expect(
      deleteAction.execute(client, {
        chat: "Incident Response 2026-04-12",
        messageId: "msg-1",
      }),
    ).rejects.toThrow(
      /Cannot delete message in "Incident Response 2026-04-12".*matched pattern "Incident \*"/,
    );
  });

  it("protection check runs before deleteMode check", async () => {
    process.env.TEAMS_PROTECTED_CONVERSATIONS = "Incident *";
    const conversation = makeConversation({
      topic: "Incident Response",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
    });

    // Even with deleteMode "block", the protection error should fire first
    await expect(
      deleteAction.execute(client, {
        chat: "Incident Response",
        messageId: "msg-1",
        deleteMode: "block",
      }),
    ).rejects.toThrow(/conversation is protected/);
  });
});
