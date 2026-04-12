/**
 * Unit tests for audit logging (src/audit.ts).
 *
 * Tests the audit event emission to stderr and file destinations,
 * and integration with edit/delete actions.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { mkdtempSync, readFileSync, rmSync, existsSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { resolveAuditDestination, emitAuditEvent } from "../../src/audit.js";
import { actions } from "../../src/actions/definitions.js";
import type { TeamsClient } from "../../src/teams-client.js";
import type {
  Conversation,
  EditedMessage,
  DeletedMessage,
  AuditEvent,
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
    topic: "Test Chat",
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

// ── resolveAuditDestination ──────────────────────────────────────────

describe("resolveAuditDestination", () => {
  beforeEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
  });

  afterEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
  });

  it("should return 'off' when TEAMS_AUDIT_LOG is not set", () => {
    expect(resolveAuditDestination()).toBe("off");
  });

  it("should return 'off' when TEAMS_AUDIT_LOG is empty", () => {
    process.env.TEAMS_AUDIT_LOG = "";
    expect(resolveAuditDestination()).toBe("off");
  });

  it("should return 'off' when TEAMS_AUDIT_LOG is 'off'", () => {
    process.env.TEAMS_AUDIT_LOG = "off";
    expect(resolveAuditDestination()).toBe("off");
  });

  it("should return 'stderr' when TEAMS_AUDIT_LOG is 'stderr'", () => {
    process.env.TEAMS_AUDIT_LOG = "stderr";
    expect(resolveAuditDestination()).toBe("stderr");
  });

  it("should return file destination when TEAMS_AUDIT_LOG starts with 'file:'", () => {
    process.env.TEAMS_AUDIT_LOG = "file:/tmp/audit.jsonl";
    expect(resolveAuditDestination()).toBe("file:/tmp/audit.jsonl");
  });

  it("should trim whitespace from TEAMS_AUDIT_LOG", () => {
    process.env.TEAMS_AUDIT_LOG = "  stderr  ";
    expect(resolveAuditDestination()).toBe("stderr");
  });

  it("should return 'off' for unrecognized values", () => {
    process.env.TEAMS_AUDIT_LOG = "webhook:https://example.com";
    expect(resolveAuditDestination()).toBe("off");
  });
});

// ── emitAuditEvent ───────────────────────────────────────────────────

describe("emitAuditEvent", () => {
  let temporaryDirectory: string;

  const sampleEvent: AuditEvent = {
    timestamp: "2026-04-12T09:00:00.000Z",
    action: "edit",
    conversationId: "19:chat@thread.v2",
    conversationLabel: "Design Review",
    messageId: "msg-123",
    content: "Updated content",
  };

  beforeEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
    temporaryDirectory = mkdtempSync(join(tmpdir(), "audit-test-"));
  });

  afterEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
    if (existsSync(temporaryDirectory)) {
      rmSync(temporaryDirectory, { recursive: true });
    }
  });

  it("should not write anything when audit is off", () => {
    const stderrSpy = vi.spyOn(process.stderr, "write").mockReturnValue(true);
    emitAuditEvent(sampleEvent);
    expect(stderrSpy).not.toHaveBeenCalled();
    stderrSpy.mockRestore();
  });

  it("should write JSON line to stderr when destination is stderr", () => {
    process.env.TEAMS_AUDIT_LOG = "stderr";
    const stderrSpy = vi.spyOn(process.stderr, "write").mockReturnValue(true);

    emitAuditEvent(sampleEvent);

    expect(stderrSpy).toHaveBeenCalledOnce();
    const written = stderrSpy.mock.calls[0][0] as string;
    expect(written).toMatch(/\n$/);
    const parsed = JSON.parse(written.trimEnd());
    expect(parsed.action).toBe("edit");
    expect(parsed.conversationId).toBe("19:chat@thread.v2");
    expect(parsed.messageId).toBe("msg-123");
    expect(parsed.content).toBe("Updated content");

    stderrSpy.mockRestore();
  });

  it("should append JSON line to file when destination is file:<path>", () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    emitAuditEvent(sampleEvent);

    const contents = readFileSync(filePath, "utf-8");
    const lines = contents.trimEnd().split("\n");
    expect(lines).toHaveLength(1);
    const parsed = JSON.parse(lines[0]);
    expect(parsed.action).toBe("edit");
    expect(parsed.messageId).toBe("msg-123");
  });

  it("should append multiple events to the same file", () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    emitAuditEvent(sampleEvent);
    emitAuditEvent({ ...sampleEvent, action: "delete", content: null });

    const contents = readFileSync(filePath, "utf-8");
    const lines = contents.trimEnd().split("\n");
    expect(lines).toHaveLength(2);
    expect(JSON.parse(lines[0]).action).toBe("edit");
    expect(JSON.parse(lines[1]).action).toBe("delete");
  });

  it("should create parent directories for file destination", () => {
    const filePath = join(temporaryDirectory, "nested", "deep", "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    emitAuditEvent(sampleEvent);

    expect(existsSync(filePath)).toBe(true);
    const parsed = JSON.parse(readFileSync(filePath, "utf-8").trimEnd());
    expect(parsed.action).toBe("edit");
  });

  it("should silently handle write errors", () => {
    process.env.TEAMS_AUDIT_LOG = "file:/nonexistent-root-path/audit.jsonl";

    // Should not throw
    expect(() => emitAuditEvent(sampleEvent)).not.toThrow();
  });
});

// ── Integration: edit-message audit ──────────────────────────────────

describe("edit-message audit logging", () => {
  const editAction = getAction("edit-message");
  let temporaryDirectory: string;

  beforeEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
    temporaryDirectory = mkdtempSync(join(tmpdir(), "audit-edit-test-"));
  });

  afterEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
    delete process.env.TEAMS_AGENT_MARKER;
    if (existsSync(temporaryDirectory)) {
      rmSync(temporaryDirectory, { recursive: true });
    }
  });

  it("should emit audit event after successful edit", async () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Design Review",
    });
    const editedMessage: EditedMessage = {
      messageId: "msg-123",
      editTime: "2026-04-12T09:00:00Z",
    };
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn().mockResolvedValue(editedMessage),
    });

    await editAction.execute(client, {
      chat: "Design Review",
      messageId: "msg-123",
      content: "Updated content",
    });

    const contents = readFileSync(filePath, "utf-8");
    const event = JSON.parse(contents.trimEnd()) as AuditEvent;
    expect(event.action).toBe("edit");
    expect(event.conversationId).toBe("19:chat@thread.v2");
    expect(event.conversationLabel).toBe("Design Review");
    expect(event.messageId).toBe("msg-123");
    expect(event.content).toBe("Updated content");
    expect(event.timestamp).toBeTruthy();
  });

  it("should include agent marker in audited content", async () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;
    process.env.TEAMS_AGENT_MARKER = "Ⓜ";

    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Test Chat",
    });
    const editedMessage: EditedMessage = {
      messageId: "msg-456",
      editTime: "2026-04-12T09:00:00Z",
    };
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn().mockResolvedValue(editedMessage),
    });

    await editAction.execute(client, {
      chat: "Test Chat",
      messageId: "msg-456",
      content: "Hello",
    });

    const event = JSON.parse(
      readFileSync(filePath, "utf-8").trimEnd(),
    ) as AuditEvent;
    expect(event.content).toBe("Ⓜ Hello");
  });

  it("should not emit audit event when audit is off", async () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    // TEAMS_AUDIT_LOG not set — defaults to off

    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Test Chat",
    });
    const editedMessage: EditedMessage = {
      messageId: "msg-789",
      editTime: "2026-04-12T09:00:00Z",
    };
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn().mockResolvedValue(editedMessage),
    });

    await editAction.execute(client, {
      chat: "Test Chat",
      messageId: "msg-789",
      content: "No audit",
    });

    expect(existsSync(filePath)).toBe(false);
  });
});

// ── Integration: delete-message audit ────────────────────────────────

describe("delete-message audit logging", () => {
  const deleteAction = getAction("delete-message");
  let temporaryDirectory: string;

  beforeEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
    delete process.env.TEAMS_DELETE_MODE;
    delete process.env.TEAMS_DELETE_TOMBSTONE;
    temporaryDirectory = mkdtempSync(join(tmpdir(), "audit-delete-test-"));
  });

  afterEach(() => {
    delete process.env.TEAMS_AUDIT_LOG;
    delete process.env.TEAMS_DELETE_MODE;
    delete process.env.TEAMS_DELETE_TOMBSTONE;
    if (existsSync(temporaryDirectory)) {
      rmSync(temporaryDirectory, { recursive: true });
    }
  });

  it("should emit audit event for hard delete", async () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Design Review",
    });
    const deletedMessage: DeletedMessage = { messageId: "msg-123" };
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      deleteMessage: vi.fn().mockResolvedValue(deletedMessage),
    });

    await deleteAction.execute(client, {
      chat: "Design Review",
      messageId: "msg-123",
    });

    const contents = readFileSync(filePath, "utf-8");
    const event = JSON.parse(contents.trimEnd()) as AuditEvent;
    expect(event.action).toBe("delete");
    expect(event.conversationId).toBe("19:chat@thread.v2");
    expect(event.conversationLabel).toBe("Design Review");
    expect(event.messageId).toBe("msg-123");
    expect(event.content).toBeNull();
  });

  it("should emit audit event for soft delete with default tombstone", async () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Design Review",
    });
    const editedMessage: EditedMessage = {
      messageId: "msg-123",
      editTime: "2026-04-12T09:00:00Z",
    };
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn().mockResolvedValue(editedMessage),
    });

    await deleteAction.execute(client, {
      chat: "Design Review",
      messageId: "msg-123",
      deleteMode: "soft",
    });

    const event = JSON.parse(
      readFileSync(filePath, "utf-8").trimEnd(),
    ) as AuditEvent;
    expect(event.action).toBe("soft-delete");
    expect(event.content).toBe("~~This message was removed by an agent~~");
  });

  it("should emit audit event for soft delete with custom tombstone", async () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Test Chat",
    });
    const editedMessage: EditedMessage = {
      messageId: "msg-456",
      editTime: "2026-04-12T09:00:00Z",
    };
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
      editMessage: vi.fn().mockResolvedValue(editedMessage),
    });

    await deleteAction.execute(client, {
      chat: "Test Chat",
      messageId: "msg-456",
      deleteMode: "soft",
      deleteTombstone: "🗑️ [removed]",
    });

    const event = JSON.parse(
      readFileSync(filePath, "utf-8").trimEnd(),
    ) as AuditEvent;
    expect(event.action).toBe("soft-delete");
    expect(event.content).toBe("🗑️ [removed]");
  });

  it("should not emit audit event when delete is blocked", async () => {
    const filePath = join(temporaryDirectory, "audit.jsonl");
    process.env.TEAMS_AUDIT_LOG = `file:${filePath}`;

    const conversation = makeConversation({
      id: "19:chat@thread.v2",
      topic: "Test Chat",
    });
    const client = createMockClient({
      findConversation: vi.fn().mockResolvedValue(conversation),
    });

    await expect(
      deleteAction.execute(client, {
        chat: "Test Chat",
        messageId: "msg-789",
        deleteMode: "block",
      }),
    ).rejects.toThrow('delete mode is set to "block"');

    expect(existsSync(filePath)).toBe(false);
  });
});
