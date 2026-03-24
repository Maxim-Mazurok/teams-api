/**
 * Message-related action definitions.
 *
 * Actions: get-messages, send-message, edit-message, delete-message.
 */

import type { Message, MessageFormat } from "../types.js";
import { isTextMessageType } from "../constants.js";
import {
  type ActionDefinition,
  toonHeader,
  extractQuote,
  buildSenderLookup,
} from "./formatters.js";
import {
  conversationParameters,
  resolveConversationId,
} from "./conversation-resolution.js";

export const getMessages: ActionDefinition = {
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

export const sendMessage: ActionDefinition = {
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

export const editMessageAction: ActionDefinition = {
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

export const deleteMessageAction: ActionDefinition = {
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
