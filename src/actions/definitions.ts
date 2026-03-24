/**
 * All action definitions for the Teams API.
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

import type { TeamsClient } from "../teams-client.js";
import type {
  Conversation,
  Message,
  Member,
  MessageFormat,
  OneOnOneSearchResult,
  PersonSearchResult,
  ChatSearchResult,
  TranscriptResult,
} from "../types.js";
import { isTextMessageType } from "../constants.js";
import {
  type ActionDefinition,
  toonHeader,
  extractQuote,
  buildSenderLookup,
  formatTimestamp,
  groupBySpeaker,
} from "./formatters.js";
import {
  conversationParameters,
  resolveConversationId,
} from "./conversation-resolution.js";

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

const findConversation: ActionDefinition = {
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

const findOneOnOne: ActionDefinition = {
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

const getMessages: ActionDefinition = {
  name: "get-messages",
  title: "Get Messages",
  description:
    "Get messages from a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required. " +
    "Messages include reactions, mentions, followers (thread subscribers), and quoted message references.",
  parameters: [
    ...conversationParameters,
    {
      name: "limit",
      type: "number",
      description:
        "Maximum number of messages to return. " +
        "Omit to fetch the entire conversation history.",
      required: false,
    },
    {
      name: "textOnly",
      type: "boolean",
      description:
        "Only return text messages, excluding system events (default: true)",
      required: false,
      default: true,
    },
    {
      name: "order",
      type: "string",
      description:
        "Message order: oldest-first (chronological, default) or newest-first",
      required: false,
      default: "oldest-first",
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    const limit = parameters.limit as number | undefined;
    const textOnly = (parameters.textOnly as boolean | undefined) ?? true;
    const onProgress = parameters.onProgress as
      | ((count: number) => void)
      | undefined;

    let messages = await client.getMessages(conversationId, {
      limit,
      onProgress,
    });

    if (textOnly) {
      messages = messages.filter(
        (message) =>
          isTextMessageType(message.messageType) && !message.isDeleted,
      );
    }

    const order = (parameters.order as string | undefined) ?? "oldest-first";
    if (order === "oldest-first") {
      messages = [...messages].reverse();
    }

    return messages;
  },
  formatResult: (result) => {
    const messages = result as Message[];
    const senderLookup = buildSenderLookup(messages);
    const lines = [`\n${messages.length} messages:\n`];
    for (const message of messages) {
      const time = message.originalArrivalTime.slice(0, 19).replace("T", " ");
      const sender = message.senderDisplayName || "(system)";
      const { quote, body } = extractQuote(message.content);

      if (quote && message.quotedMessageId) {
        const quotedSender =
          senderLookup.get(message.quotedMessageId) ?? "unknown";
        lines.push(`  [${time}] ${sender}:`);
        lines.push(
          `    > [replying to ${quotedSender}]: ${quote.slice(0, 80)}`,
        );
        lines.push(`    ${body.slice(0, 120)}`);
      } else {
        lines.push(`  [${time}] ${sender}: ${body.slice(0, 120)}`);
      }
      if (message.followers.length > 0) {
        lines.push(`    [${message.followers.length} follower(s)]`);
      }
    }
    return lines.join("\n");
  },
  formatMarkdown: (result) => {
    const messages = result as Message[];
    const senderLookup = buildSenderLookup(messages);
    const lines = [`## Messages (${messages.length})`, ""];
    let previousSender = "";
    for (const message of messages) {
      const time = message.originalArrivalTime.slice(0, 19).replace("T", " ");
      const sender = message.senderDisplayName || "(system)";
      const { quote, body } = extractQuote(message.content);

      if (sender === previousSender) {
        lines.push(`*${time}*`, "");
      } else {
        lines.push(`### ${sender} — ${time}`, "");
        previousSender = sender;
      }

      if (quote && message.quotedMessageId) {
        const quotedSender =
          senderLookup.get(message.quotedMessageId) ?? "unknown";
        lines.push(`> **[replying to ${quotedSender}]:** ${quote}`, "");
      }

      lines.push(body, "");

      if (message.followers.length > 0) {
        lines.push(`*${message.followers.length} follower(s)*`, "");
      }
    }
    return lines.join("\n");
  },
  formatToon: (result) => {
    const messages = result as Message[];
    const senderLookup = buildSenderLookup(messages);
    const lines = [toonHeader("💬", `${messages.length} Messages`)];
    let previousSender = "";
    for (const message of messages) {
      const time = message.originalArrivalTime.slice(0, 19).replace("T", " ");
      const sender = message.senderDisplayName || "(system)";
      const { quote, body } = extractQuote(message.content);

      lines.push("");
      if (sender === previousSender) {
        lines.push(`      ${time}`);
      } else {
        lines.push(`  🗣️  ${sender} · ${time}`);
        previousSender = sender;
      }

      if (quote && message.quotedMessageId) {
        const quotedSender =
          senderLookup.get(message.quotedMessageId) ?? "unknown";
        lines.push(
          `      > [replying to ${quotedSender}]: ${quote.slice(0, 80)}`,
        );
      }
      lines.push(`      ${body.slice(0, 120)}`);
      if (message.followers.length > 0) {
        lines.push(`      👥 ${message.followers.length} follower(s)`);
      }
    }
    return lines.join("\n");
  },
};

const sendMessage: ActionDefinition = {
  name: "send-message",
  title: "Send Message",
  description:
    "Send a message to a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required. " +
    "Content is interpreted as Markdown by default and converted to rich HTML.",
  parameters: [
    ...conversationParameters,
    {
      name: "content",
      type: "string",
      description: "Message content to send",
      required: true,
    },
    {
      name: "messageFormat",
      type: "string",
      description:
        'Content format: "markdown" (default, converted to HTML), "html" (raw HTML), or "text" (plain text)',
      required: false,
      default: "markdown",
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId, label } = await resolveConversationId(
      client,
      parameters,
    );
    const content = parameters.content as string;
    const messageFormat =
      (parameters.messageFormat as MessageFormat | undefined) ?? "markdown";
    const result = await client.sendMessage(
      conversationId,
      content,
      messageFormat,
    );
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
  formatMarkdown: (result) => {
    const { messageId, arrivalTime, conversation } = result as {
      messageId: string;
      arrivalTime: number;
      conversation: string;
    };
    return [
      "## Message Sent",
      "",
      `- **To:** ${conversation}`,
      `- **Message ID:** ${messageId}`,
      `- **Arrival time:** ${arrivalTime}`,
    ].join("\n");
  },
  formatToon: (result) => {
    const { messageId, arrivalTime, conversation } = result as {
      messageId: string;
      arrivalTime: number;
      conversation: string;
    };
    return [
      toonHeader("✅", "Message Sent!"),
      `  📨 To: "${conversation}"`,
      `  🆔 ${messageId}`,
      `  ⏰ ${arrivalTime}`,
    ].join("\n");
  },
};

const editMessageAction: ActionDefinition = {
  name: "edit-message",
  title: "Edit Message",
  description:
    "Edit an existing message in a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required. " +
    "The message to edit is identified by --message-id. " +
    "Content is interpreted as Markdown by default and converted to rich HTML.",
  parameters: [
    ...conversationParameters,
    {
      name: "messageId",
      type: "string",
      description: "ID of the message to edit",
      required: true,
    },
    {
      name: "content",
      type: "string",
      description: "New message content",
      required: true,
    },
    {
      name: "messageFormat",
      type: "string",
      description:
        'Content format: "markdown" (default, converted to HTML), "html" (raw HTML), or "text" (plain text)',
      required: false,
      default: "markdown",
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId, label } = await resolveConversationId(
      client,
      parameters,
    );
    const messageId = parameters.messageId as string;
    const content = parameters.content as string;
    const messageFormat =
      (parameters.messageFormat as MessageFormat | undefined) ?? "markdown";
    const result = await client.editMessage(
      conversationId,
      messageId,
      content,
      messageFormat,
    );
    return { ...result, conversation: label };
  },
  formatResult: (result) => {
    const { messageId, editTime, conversation } = result as {
      messageId: string;
      editTime: string;
      conversation: string;
    };
    return [
      `Message edited in "${conversation}"`,
      `  Message ID: ${messageId}`,
      `  Edit time: ${editTime}`,
    ].join("\n");
  },
  formatMarkdown: (result) => {
    const { messageId, editTime, conversation } = result as {
      messageId: string;
      editTime: string;
      conversation: string;
    };
    return [
      "## Message Edited",
      "",
      `- **In:** ${conversation}`,
      `- **Message ID:** ${messageId}`,
      `- **Edit time:** ${editTime}`,
    ].join("\n");
  },
  formatToon: (result) => {
    const { messageId, editTime, conversation } = result as {
      messageId: string;
      editTime: string;
      conversation: string;
    };
    return [
      toonHeader("✏️", "Message Edited!"),
      `  💬 In: "${conversation}"`,
      `  🆔 ${messageId}`,
      `  ⏰ ${editTime}`,
    ].join("\n");
  },
};

const deleteMessageAction: ActionDefinition = {
  name: "delete-message",
  title: "Delete Message",
  description:
    "Delete a message from a conversation. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "At least one identifier is required. " +
    "The message to delete is identified by --message-id.",
  parameters: [
    ...conversationParameters,
    {
      name: "messageId",
      type: "string",
      description: "ID of the message to delete",
      required: true,
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId, label } = await resolveConversationId(
      client,
      parameters,
    );
    const messageId = parameters.messageId as string;
    const result = await client.deleteMessage(conversationId, messageId);
    return { ...result, conversation: label };
  },
  formatResult: (result) => {
    const { messageId, conversation } = result as {
      messageId: string;
      conversation: string;
    };
    return [
      `Message deleted from "${conversation}"`,
      `  Message ID: ${messageId}`,
    ].join("\n");
  },
  formatMarkdown: (result) => {
    const { messageId, conversation } = result as {
      messageId: string;
      conversation: string;
    };
    return [
      "## Message Deleted",
      "",
      `- **From:** ${conversation}`,
      `- **Message ID:** ${messageId}`,
    ].join("\n");
  },
  formatToon: (result) => {
    const { messageId, conversation } = result as {
      messageId: string;
      conversation: string;
    };
    return [
      toonHeader("🗑️", "Message Deleted!"),
      `  💬 From: "${conversation}"`,
      `  🆔 ${messageId}`,
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
    "Display names are resolved via the Teams profile API when available, with message history as fallback. " +
    "Note: 1:1 chat members may have empty display names if profile resolution is unavailable.",
  parameters: [...conversationParameters],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    return client.getMembers(conversationId);
  },
  formatResult: (result) => {
    const members = result as Member[];
    const people = members.filter((member) => member.memberType === "person");
    const bots = members.filter((member) => member.memberType === "bot");
    const lines = [`\n${people.length} people, ${bots.length} bots:\n`];
    for (const member of people) {
      const name = member.displayName || "(unknown)";
      lines.push(`  ${name} (${member.role}) — ${member.id}`);
    }
    if (bots.length > 0) {
      lines.push("");
      lines.push("  Bots/Apps:");
      for (const bot of bots) {
        const name = bot.displayName || "(unnamed bot)";
        lines.push(`  ${name} — ${bot.id}`);
      }
    }
    return lines.join("\n");
  },
  formatMarkdown: (result) => {
    const members = result as Member[];
    const people = members.filter((member) => member.memberType === "person");
    const bots = members.filter((member) => member.memberType === "bot");
    const lines = [
      `## Members (${people.length} people, ${bots.length} bots)`,
      "",
    ];
    if (people.length > 0) {
      lines.push("| Name | Role | ID |");
      lines.push("|------|------|----|");
      for (const member of people) {
        const name = member.displayName || "(unknown)";
        lines.push(`| ${name} | ${member.role} | ${member.id} |`);
      }
    }
    if (bots.length > 0) {
      lines.push("", "### Bots/Apps", "");
      lines.push("| Name | ID |");
      lines.push("|------|----|");
      for (const bot of bots) {
        const name = bot.displayName || "(unnamed bot)";
        lines.push(`| ${name} | ${bot.id} |`);
      }
    }
    return lines.join("\n");
  },
  formatToon: (result) => {
    const members = result as Member[];
    const people = members.filter((member) => member.memberType === "person");
    const bots = members.filter((member) => member.memberType === "bot");
    const lines = [
      toonHeader("👥", `${people.length} People, ${bots.length} Bots`),
    ];
    for (const member of people) {
      const name = member.displayName || "(unknown)";
      lines.push("");
      lines.push(`  👤 ${name} · ${member.role}`);
      lines.push(`     ${member.id}`);
    }
    if (bots.length > 0) {
      lines.push("");
      lines.push("  🤖 Bots/Apps:");
      for (const bot of bots) {
        const name = bot.displayName || "(unnamed bot)";
        lines.push(`     🤖 ${name} — ${bot.id}`);
      }
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
  formatMarkdown: (result) => {
    const { displayName, region } = result as {
      displayName: string;
      region: string;
    };
    return [`## ${displayName}`, "", `- **Region:** ${region}`].join("\n");
  },
  formatToon: (result) => {
    const { displayName, region } = result as {
      displayName: string;
      region: string;
    };
    return [toonHeader("🙋", displayName), `  📍 region: ${region}`].join("\n");
  },
};

const getTranscript: ActionDefinition = {
  name: "get-transcript",
  title: "Get Meeting Transcript",
  description:
    "Get the meeting transcript from a conversation that contains a recorded meeting. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "Use --raw-vtt to get the original VTT file instead of parsed output.",
  parameters: [
    ...conversationParameters,
    {
      name: "rawVtt",
      type: "boolean",
      description:
        "Return the original VTT file content instead of parsed transcript (default: false)",
      required: false,
      default: false,
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    const rawVtt = (parameters.rawVtt as boolean | undefined) ?? false;

    const transcriptResult = await client.getTranscript(conversationId);

    if (rawVtt) {
      return { rawVtt: transcriptResult.rawVtt, format: "vtt" as const };
    }

    return transcriptResult;
  },
  formatResult: (result) => {
    const data = result as TranscriptResult | { rawVtt: string; format: "vtt" };

    if ("format" in data && data.format === "vtt") {
      return data.rawVtt;
    }

    const transcript = data as TranscriptResult;
    const groups = groupBySpeaker(transcript.entries);
    const lines = [
      `\nTranscript: ${transcript.meetingTitle} (${transcript.entries.length} segments)\n`,
    ];

    for (const group of groups) {
      const time = formatTimestamp(group.startTime);
      lines.push(`  [${time}] ${group.speaker}:`);
      lines.push(`    ${group.segments.join(" ")}`);
    }

    return lines.join("\n");
  },
  formatMarkdown: (result) => {
    const data = result as TranscriptResult | { rawVtt: string; format: "vtt" };

    if ("format" in data && data.format === "vtt") {
      return ["```vtt", data.rawVtt, "```"].join("\n");
    }

    const transcript = data as TranscriptResult;
    const groups = groupBySpeaker(transcript.entries);
    const lines = [
      `## Transcript: ${transcript.meetingTitle}`,
      "",
      `*${transcript.entries.length} segments*`,
      "",
    ];

    for (const group of groups) {
      const time = formatTimestamp(group.startTime);
      lines.push(`**${group.speaker}** *(${time})*`, "");
      lines.push(group.segments.join(" "), "");
    }

    return lines.join("\n");
  },
  formatToon: (result) => {
    const data = result as TranscriptResult | { rawVtt: string; format: "vtt" };

    if ("format" in data && data.format === "vtt") {
      return data.rawVtt;
    }

    const transcript = data as TranscriptResult;
    const groups = groupBySpeaker(transcript.entries);
    const lines = [
      toonHeader(
        "🎙️",
        `Transcript: ${transcript.meetingTitle} (${transcript.entries.length} segments)`,
      ),
    ];

    for (const group of groups) {
      const time = formatTimestamp(group.startTime);
      lines.push("");
      lines.push(`  🗣️  ${group.speaker} · ${time}`);
      lines.push(`      ${group.segments.join(" ")}`);
    }

    return lines.join("\n");
  },
};

const findPeopleAction: ActionDefinition = {
  name: "find-people",
  title: "Find People",
  description:
    "Search for people in the organization directory by name. " +
    "Uses the Substrate search API (requires authentication via auto-login or interactive). " +
    "Returns matching people with emails, job titles, and departments.",
  parameters: [
    {
      name: "query",
      type: "string",
      description: "Name or partial name to search for",
      required: true,
    },
    {
      name: "maxResults",
      type: "number",
      description: "Maximum results to return (default: 10)",
      required: false,
      default: 10,
    },
  ],
  execute: async (client, parameters) => {
    const query = parameters.query as string;
    const maxResults = (parameters.maxResults as number) ?? 10;
    return client.findPeople(query, maxResults);
  },
  formatResult: (result) => {
    const people = result as PersonSearchResult[];
    if (people.length === 0) return "No people found.";
    return people
      .map(
        (person) =>
          `${person.displayName} <${person.email}> — ${person.jobTitle || "no title"}, ${person.department || "no department"}`,
      )
      .join("\n");
  },
  formatMarkdown: (result) => {
    const people = result as PersonSearchResult[];
    if (people.length === 0) return "No people found.";
    const lines = [`## People (${people.length} found)`, ""];
    for (const person of people) {
      lines.push(`### ${person.displayName}`);
      lines.push(`- **Email:** ${person.email}`);
      if (person.jobTitle) lines.push(`- **Title:** ${person.jobTitle}`);
      if (person.department)
        lines.push(`- **Department:** ${person.department}`);
      lines.push(`- **MRI:** ${person.mri}`);
      lines.push("");
    }
    return lines.join("\n");
  },
  formatToon: (result) => {
    const people = result as PersonSearchResult[];
    if (people.length === 0) return "\n  🔍 No people found.";
    const lines = [toonHeader("👥", `Found ${people.length} people`)];
    for (const person of people) {
      lines.push(`  👤 ${person.displayName}`);
      lines.push(
        `     📧 ${person.email} · ${person.jobTitle || "?"} · ${person.department || "?"}`,
      );
    }
    return lines.join("\n");
  },
};

const findChatsAction: ActionDefinition = {
  name: "find-chats",
  title: "Find Chats",
  description:
    "Search for chats by name or member name. " +
    "Uses the Substrate search API (requires authentication via auto-login or interactive). " +
    "Returns matching chats with member lists and thread IDs.",
  parameters: [
    {
      name: "query",
      type: "string",
      description: "Chat name or member name to search for",
      required: true,
    },
    {
      name: "maxResults",
      type: "number",
      description: "Maximum results to return (default: 10)",
      required: false,
      default: 10,
    },
  ],
  execute: async (client, parameters) => {
    const query = parameters.query as string;
    const maxResults = (parameters.maxResults as number) ?? 10;
    return client.findChats(query, maxResults);
  },
  formatResult: (result) => {
    const chats = result as ChatSearchResult[];
    if (chats.length === 0) return "No chats found.";
    return chats
      .map((chat) => {
        const name = chat.name || "(untitled)";
        const members = chat.matchingMembers
          .map((member) => member.displayName)
          .join(", ");
        return `${name} (${chat.threadType}, ${chat.totalMemberCount} members${members ? `, matched: ${members}` : ""}) — ${chat.threadId}`;
      })
      .join("\n");
  },
  formatMarkdown: (result) => {
    const chats = result as ChatSearchResult[];
    if (chats.length === 0) return "No chats found.";
    const lines = [`## Chats (${chats.length} found)`, ""];
    for (const chat of chats) {
      lines.push(`### ${chat.name || "(untitled)"}`);
      lines.push(`- **Thread ID:** ${chat.threadId}`);
      lines.push(`- **Type:** ${chat.threadType}`);
      lines.push(`- **Members:** ${chat.totalMemberCount}`);
      if (chat.matchingMembers.length > 0) {
        lines.push(
          `- **Matched:** ${chat.matchingMembers.map((member) => member.displayName).join(", ")}`,
        );
      }
      lines.push("");
    }
    return lines.join("\n");
  },
  formatToon: (result) => {
    const chats = result as ChatSearchResult[];
    if (chats.length === 0) return "\n  🔍 No chats found.";
    const lines = [toonHeader("💬", `Found ${chats.length} chats`)];
    for (const chat of chats) {
      lines.push(`  💬 ${chat.name || "(untitled)"}`);
      lines.push(
        `     📁 ${chat.threadType} · ${chat.totalMemberCount} members`,
      );
      if (chat.matchingMembers.length > 0) {
        const matched = chat.matchingMembers
          .map((member) => member.displayName)
          .join(", ");
        lines.push(`     🎯 Matched: ${matched}`);
      }
    }
    return lines.join("\n");
  },
};

// ── Registry ─────────────────────────────────────────────────────────

/**
 * Map-based action registry keyed by action name.
 * Ensures compile-time visibility of all actions and prevents
 * accidental omissions from the exported array.
 */
const actionRegistry = new Map<string, ActionDefinition>([
  ["list-conversations", listConversations],
  ["find-conversation", findConversation],
  ["find-one-on-one", findOneOnOne],
  ["find-people", findPeopleAction],
  ["find-chats", findChatsAction],
  ["get-messages", getMessages],
  ["send-message", sendMessage],
  ["edit-message", editMessageAction],
  ["delete-message", deleteMessageAction],
  ["get-members", getMembers],
  ["whoami", whoami],
  ["get-transcript", getTranscript],
]);

/** All registered actions, derived from the registry map. */
export const actions: ActionDefinition[] = Array.from(actionRegistry.values());
