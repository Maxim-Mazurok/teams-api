/**
 * Conversation-related action definitions.
 *
 * Actions: list-conversations, find-conversation, find-one-on-one.
 */

import type { Conversation, OneOnOneSearchResult } from "../types.js";
import { type ActionDefinition, toonHeader } from "./formatters.js";

export const listConversations: ActionDefinition = {
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
  formatMarkdown: (result) => {
    const conversations = result as Conversation[];
    const lines = [`## Conversations (${conversations.length})`, ""];
    if (conversations.length === 0) return lines.join("\n");
    lines.push("| # | Topic | Type | Members | Last Message |");
    lines.push("|---|-------|------|---------|--------------|");
    for (let i = 0; i < conversations.length; i++) {
      const conversation = conversations[i];
      const lastMessage =
        conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
      const topic = conversation.topic || "(untitled 1:1 chat)";
      lines.push(
        `| ${i} | ${topic} | ${conversation.threadType} | ${conversation.memberCount ?? "?"} | ${lastMessage} |`,
      );
    }
    return lines.join("\n");
  },
  formatToon: (result) => {
    const conversations = result as Conversation[];
    const lines = [toonHeader("📋", `${conversations.length} Conversations`)];
    for (let i = 0; i < conversations.length; i++) {
      const conversation = conversations[i];
      const lastMessage =
        conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
      const topic = conversation.topic || "(untitled 1:1 chat)";
      lines.push("");
      lines.push(`  💬 [${i}] "${topic}"`);
      lines.push(
        `      ${conversation.threadType} · ${conversation.memberCount ?? "?"} members · last: ${lastMessage}`,
      );
    }
    return lines.join("\n");
  },
};

export const findConversation: ActionDefinition = {
  name: "find-conversation",
  title: "Find Teams Conversation",
  description:
    "Find a conversation by topic name (case-insensitive partial match). " +
    "When Substrate search is available, also matches by member names. " +
    "For 1:1 chats (which have no topic), use find-one-on-one instead. " +
    "Use the returned conversation ID for subsequent operations like get-messages or send-message.",
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
  formatMarkdown: (result) => {
    if (!result) return "No conversation found.";
    const conversation = result as Conversation;
    const lastMessage = conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
    return [
      `## Found: "${conversation.topic}"`,
      "",
      `- **ID:** ${conversation.id}`,
      `- **Type:** ${conversation.threadType}`,
      `- **Members:** ${conversation.memberCount ?? "?"}`,
      `- **Last message:** ${lastMessage}`,
    ].join("\n");
  },
  formatToon: (result) => {
    if (!result) return "\n  🔍 No conversation found.";
    const conversation = result as Conversation;
    const lastMessage = conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
    return [
      toonHeader("🔍", `Found: "${conversation.topic}"`),
      `  🆔 ${conversation.id}`,
      `  📁 ${conversation.threadType} · ${conversation.memberCount ?? "?"} members · last: ${lastMessage}`,
    ].join("\n");
  },
};

export const findOneOnOne: ActionDefinition = {
  name: "find-one-on-one",
  title: "Find 1:1 Conversation",
  description:
    "Find a 1:1 conversation with a person by name. " +
    "Uses Substrate people/chat search when available, " +
    "falls back to scanning message senders. " +
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
  formatMarkdown: (result) => {
    if (!result) return "No 1:1 conversation found.";
    const searchResult = result as OneOnOneSearchResult;
    return [
      `## Found 1:1 with ${searchResult.memberDisplayName}`,
      "",
      `- **Conversation ID:** ${searchResult.conversationId}`,
    ].join("\n");
  },
  formatToon: (result) => {
    if (!result) return "\n  🔍 No 1:1 conversation found.";
    const searchResult = result as OneOnOneSearchResult;
    return [
      toonHeader("🔍", `Found 1:1 with ${searchResult.memberDisplayName}`),
      `  🆔 ${searchResult.conversationId}`,
    ].join("\n");
  },
};
