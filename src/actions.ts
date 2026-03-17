/**
 * Unified action definitions for the Teams API.
 *
 * This is the single source of truth for all operations. CLI commands,
 * MCP tools, and programmatic usage all derive from these definitions.
 *
 * Each action declares:
 *   - name, title, description — shared help text and documentation
 *   - parameters — typed parameter definitions with descriptions and defaults
 *   - execute — the implementation, calling TeamsClient methods
 *   - formatResult — human-readable output formatter (CLI without --json)
 */

import type { TeamsClient } from "./teams-client.js";
import type {
  Conversation,
  Message,
  Member,
  OneOnOneSearchResult,
} from "./types.js";

// ── Parameter & Action types ─────────────────────────────────────────

export interface ActionParameter {
  /** Parameter name in camelCase (CLI flags auto-converted to kebab-case). */
  name: string;
  /** Parameter type. Determines CLI flag syntax and MCP Zod schema. */
  type: "string" | "number" | "boolean";
  /** Description for CLI help, MCP tool description, and documentation. */
  description: string;
  /** Whether the parameter must be provided. */
  required: boolean;
  /** Default value when parameter is omitted. */
  default?: string | number | boolean;
}

export interface ActionDefinition {
  /** Kebab-case name. CLI command name; MCP tool name is `teams_` + snake_case. */
  name: string;
  /** Human-readable title (MCP tool title). */
  title: string;
  /** Full description shared across CLI help, MCP, and documentation. */
  description: string;
  /** Typed parameter definitions. */
  parameters: ActionParameter[];
  /** Execute the action against a TeamsClient. */
  execute: (
    client: TeamsClient,
    parameters: Record<string, unknown>,
  ) => Promise<unknown>;
  /** Format result as human-readable text (CLI output without --json). */
  formatResult: (result: unknown) => string;
}

// ── Shared conversation resolution ───────────────────────────────────

/**
 * Resolve a conversation ID from the standard identification parameters.
 *
 * Supports three ways to identify a conversation:
 *   1. conversationId — direct thread ID
 *   2. chat — topic name (partial match via findConversation)
 *   3. to — person name (1:1 lookup via findOneOnOneConversation)
 *
 * Returns both the resolved ID and a human-readable label.
 */
async function resolveConversationId(
  client: TeamsClient,
  parameters: Record<string, unknown>,
): Promise<{ conversationId: string; label: string }> {
  const conversationId = parameters.conversationId as string | undefined;
  const chat = parameters.chat as string | undefined;
  const to = parameters.to as string | undefined;

  if (conversationId) {
    return { conversationId, label: conversationId };
  }

  if (chat) {
    const conversation = await client.findConversation(chat);
    if (!conversation) {
      throw new Error(`No conversation matching "${chat}" found.`);
    }
    return { conversationId: conversation.id, label: conversation.topic };
  }

  if (to) {
    const result = await client.findOneOnOneConversation(to);
    if (!result) {
      throw new Error(`No 1:1 conversation found with "${to}".`);
    }
    return {
      conversationId: result.conversationId,
      label: result.memberDisplayName,
    };
  }

  throw new Error("One of --conversation-id, --chat, or --to is required.");
}

/** Shared parameter definitions for conversation identification. */
const conversationParameters: ActionParameter[] = [
  {
    name: "chat",
    type: "string",
    description: "Find conversation by topic name (partial match)",
    required: false,
  },
  {
    name: "to",
    type: "string",
    description: "Find 1:1 conversation by person name",
    required: false,
  },
  {
    name: "conversationId",
    type: "string",
    description: "Direct conversation thread ID",
    required: false,
  },
];

// ── Action definitions ───────────────────────────────────────────────

const listConversations: ActionDefinition = {
  name: "list-conversations",
  title: "List Teams Conversations",
  description:
    "List conversations (chats, group chats, meetings, channels). " +
    "Returns conversation ID, topic, type, member count, and last message time.",
  parameters: [
    {
      name: "limit",
      type: "number",
      description: "Maximum number of conversations to return",
      required: false,
      default: 50,
    },
  ],
  execute: async (client, parameters) => {
    const limit = (parameters.limit as number | undefined) ?? 50;
    return client.listConversations({ pageSize: limit });
  },
  formatResult: (result) => {
    const conversations = result as Conversation[];
    const lines = [`\n${conversations.length} conversations:\n`];
    for (let i = 0; i < conversations.length; i++) {
      const conversation = conversations[i];
      const lastMessage =
        conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
      const topic = conversation.topic || "(untitled 1:1 chat)";
      lines.push(
        `  [${i}] ${conversation.threadType}: "${topic}" ` +
          `(members: ${conversation.memberCount ?? "?"}, last: ${lastMessage})`,
      );
    }
    return lines.join("\n");
  },
};

const findConversation: ActionDefinition = {
  name: "find-conversation",
  title: "Find Teams Conversation",
  description:
    "Find a conversation by topic name (case-insensitive partial match). " +
    "For 1:1 chats (which have no topic), use find-one-on-one instead.",
  parameters: [
    {
      name: "query",
      type: "string",
      description: "Partial topic name to search for",
      required: true,
    },
  ],
  execute: async (client, parameters) => {
    const query = parameters.query as string;
    return client.findConversation(query);
  },
  formatResult: (result) => {
    if (!result) return "No conversation found.";
    const conversation = result as Conversation;
    const lastMessage = conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
    return (
      `Found: "${conversation.topic}" ` +
      `(${conversation.id}, ${conversation.threadType}, ` +
      `members: ${conversation.memberCount ?? "?"}, last: ${lastMessage})`
    );
  },
};

const findOneOnOne: ActionDefinition = {
  name: "find-one-on-one",
  title: "Find 1:1 Conversation",
  description:
    "Find a 1:1 conversation with a person by name. " +
    "Searches untitled chats by scanning recent message sender names. " +
    "Also finds the self-chat if the name matches the current user.",
  parameters: [
    {
      name: "personName",
      type: "string",
      description:
        "Name of the person to find (case-insensitive partial match)",
      required: true,
    },
  ],
  execute: async (client, parameters) => {
    const personName = parameters.personName as string;
    return client.findOneOnOneConversation(personName);
  },
  formatResult: (result) => {
    if (!result) return "No 1:1 conversation found.";
    const searchResult = result as OneOnOneSearchResult;
    return `Found 1:1 with ${searchResult.memberDisplayName} (${searchResult.conversationId})`;
  },
};

const getMessages: ActionDefinition = {
  name: "get-messages",
  title: "Get Messages",
  description:
    "Get messages from a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required.",
  parameters: [
    ...conversationParameters,
    {
      name: "maxPages",
      type: "number",
      description: "Maximum pagination pages to fetch",
      required: false,
      default: 100,
    },
    {
      name: "pageSize",
      type: "number",
      description: "Messages per page",
      required: false,
      default: 200,
    },
    {
      name: "textOnly",
      type: "boolean",
      description:
        "Only return text messages, excluding system events (default: true)",
      required: false,
      default: true,
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    const maxPages = (parameters.maxPages as number | undefined) ?? 100;
    const pageSize = (parameters.pageSize as number | undefined) ?? 200;
    const textOnly = (parameters.textOnly as boolean | undefined) ?? true;
    const onProgress = parameters.onProgress as
      | ((count: number) => void)
      | undefined;

    let messages = await client.getMessages(conversationId, {
      maxPages,
      pageSize,
      onProgress,
    });

    if (textOnly) {
      messages = messages.filter(
        (message) =>
          (message.messageType === "RichText/Html" ||
            message.messageType === "Text") &&
          !message.isDeleted,
      );
    }

    return messages;
  },
  formatResult: (result) => {
    const messages = result as Message[];
    const lines = [`\n${messages.length} messages:\n`];
    for (const message of messages) {
      const time = message.originalArrivalTime.slice(0, 19).replace("T", " ");
      const sender = message.senderDisplayName || "(system)";
      const preview = message.content.replace(/<[^>]*>/g, "").slice(0, 120);
      lines.push(`  [${time}] ${sender}: ${preview}`);
    }
    return lines.join("\n");
  },
};

const sendMessage: ActionDefinition = {
  name: "send-message",
  title: "Send Message",
  description:
    "Send a plain-text message to a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required.",
  parameters: [
    ...conversationParameters,
    {
      name: "content",
      type: "string",
      description: "Message text to send",
      required: true,
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId, label } = await resolveConversationId(
      client,
      parameters,
    );
    const content = parameters.content as string;
    const result = await client.sendMessage(conversationId, content);
    return { ...result, conversation: label };
  },
  formatResult: (result) => {
    const { messageId, arrivalTime, conversation } = result as {
      messageId: string;
      arrivalTime: number;
      conversation: string;
    };
    return [
      `Message sent to "${conversation}"`,
      `  Message ID: ${messageId}`,
      `  Arrival time: ${arrivalTime}`,
    ].join("\n");
  },
};

const getMembers: ActionDefinition = {
  name: "get-members",
  title: "Get Conversation Members",
  description:
    "List members of a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required. " +
    "Note: 1:1 chat members may have empty display names.",
  parameters: [...conversationParameters],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    return client.getMembers(conversationId);
  },
  formatResult: (result) => {
    const members = result as Member[];
    const lines = [`\n${members.length} members:\n`];
    for (const member of members) {
      const name = member.displayName || "(unknown)";
      lines.push(`  ${name} (${member.role}) — ${member.id}`);
    }
    return lines.join("\n");
  },
};

const whoami: ActionDefinition = {
  name: "whoami",
  title: "Current User Info",
  description:
    "Get the display name and region of the currently authenticated user.",
  parameters: [],
  execute: async (client) => {
    const displayName = await client.getCurrentUserDisplayName();
    const token = client.getToken();
    return { displayName, region: token.region };
  },
  formatResult: (result) => {
    const { displayName, region } = result as {
      displayName: string;
      region: string;
    };
    return `${displayName} (region: ${region})`;
  },
};

// ── Registry ─────────────────────────────────────────────────────────

export const actions: ActionDefinition[] = [
  listConversations,
  findConversation,
  findOneOnOne,
  getMessages,
  sendMessage,
  getMembers,
  whoami,
];
