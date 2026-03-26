/**
 * Conversation-related action definitions.
 *
 * Actions: list-conversations, find-conversation, find-one-on-one.
 */

import type { Conversation, OneOnOneSearchResult } from "../types.js";
import type { ActionDefinition } from "./formatters.js";

export const listConversations: ActionDefinition = {
  name: "list-conversations",
  title: "List Teams Conversations",
  description:
    "List conversations (chats, group chats, meetings, channels). " +
    "Returns conversation ID, topic, type, and last message time.",
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
  formatConcise: (result) => {
    const conversations = result as Conversation[];
    const lines = [`## Conversations (${conversations.length})`, ""];
    if (conversations.length === 0) return lines.join("\n");
    lines.push("| # | Topic | Type | ID | Last Message |");
    lines.push("|---|-------|------|----|--------------|");
    for (let i = 0; i < conversations.length; i++) {
      const conversation = conversations[i];
      const lastMessage =
        conversation.lastMessageTime?.slice(0, 10) ?? "";
      const topic = conversation.topic || "(untitled)";
      lines.push(
        `| ${i} | ${topic} | ${conversation.threadType} | ${conversation.id} | ${lastMessage} |`,
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
  formatConcise: (result) => {
    if (!result) return "No conversation found.";
    const conversation = result as Conversation;
    const lines = [
      `## Found: "${conversation.topic}"`,
      "",
      `- **ID:** ${conversation.id}`,
      `- **Type:** ${conversation.threadType}`,
    ];
    if (conversation.lastMessageTime) {
      lines.push(`- **Last message:** ${conversation.lastMessageTime.slice(0, 10)}`);
    }
    return lines.join("\n");
  },
};

export const findOneOnOne: ActionDefinition = {
  name: "find-one-on-one",
  title: "Find 1:1 Conversation",
  description:
    "Find or create a 1:1 conversation with a person by name. " +
    "Uses Substrate people/chat search when available — if a conversation " +
    "already exists it is returned; if the person is found in the org directory " +
    "but no chat exists yet, a new 1:1 is started. " +
    "Falls back to scanning message senders when Substrate is unavailable. " +
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
  formatConcise: (result) => {
    if (!result) return "No 1:1 conversation found.";
    const searchResult = result as OneOnOneSearchResult;
    return [
      `## Found 1:1 with ${searchResult.memberDisplayName}`,
      "",
      `- **Conversation ID:** ${searchResult.conversationId}`,
    ].join("\n");
  },
};
