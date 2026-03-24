/**
 * Message-related action definitions.
 *
 * Actions: get-messages, send-message, edit-message, delete-message.
 */

import { readFileSync } from "node:fs";
import { basename, extname } from "node:path";
import type { Message, MessageFormat, MessageContentPart, ScheduledMessage } from "../types.js";
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

/** Map file extension to MIME content type. */
function contentTypeFromExtension(filePath: string): string {
  const extension = extname(filePath).toLowerCase();
  const mimeTypes: Record<string, string> = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".bmp": "image/bmp",
  };
  return mimeTypes[extension] ?? "image/png";
}

/**
 * Build MessageContentPart[] from content text, image file paths, and file paths.
 *
 * If the content contains `[image]` placeholders, images are interleaved
 * at those positions. Otherwise, the text comes first followed by all images.
 * File attachments are always appended at the end.
 */
function buildContentParts(
  content: string,
  imagePaths: string[],
  filePaths: string[] = [],
): MessageContentPart[] {
  const parts: MessageContentPart[] = [];

  if (content.includes("[image]") && imagePaths.length > 0) {
    const textSegments = content.split("[image]");
    let imageIndex = 0;
    for (let i = 0; i < textSegments.length; i++) {
      const segment = textSegments[i].trim();
      if (segment) {
        parts.push({ type: "text", text: segment });
      }
      if (i < textSegments.length - 1 && imageIndex < imagePaths.length) {
        const filePath = imagePaths[imageIndex];
        parts.push({
          type: "image",
          data: readFileSync(filePath),
          fileName: basename(filePath),
          contentType: contentTypeFromExtension(filePath),
        });
        imageIndex++;
      }
    }
    // Append remaining images not mapped to placeholders
    for (; imageIndex < imagePaths.length; imageIndex++) {
      const filePath = imagePaths[imageIndex];
      parts.push({
        type: "image",
        data: readFileSync(filePath),
        fileName: basename(filePath),
        contentType: contentTypeFromExtension(filePath),
      });
    }
  } else {
    if (content) {
      parts.push({ type: "text", text: content });
    }
    for (const filePath of imagePaths) {
      parts.push({
        type: "image",
        data: readFileSync(filePath),
        fileName: basename(filePath),
        contentType: contentTypeFromExtension(filePath),
      });
    }
  }

  // Append file attachments
  for (const filePath of filePaths) {
    parts.push({
      type: "file",
      data: readFileSync(filePath),
      fileName: basename(filePath),
    });
  }

  return parts;
}

/** Build a concise attachment summary string for a message. */
function formatAttachmentSummary(message: Message): string {
  const parts: string[] = [];
  if (message.images.length > 0) {
    parts.push(
      message.images.length === 1
        ? "1 image"
        : `${message.images.length} images`,
    );
  }
  if (message.files.length > 0) {
    const fileNames = message.files.map((file) => file.fileName);
    parts.push(fileNames.join(", "));
  }
  return parts.length > 0 ? `[📎 ${parts.join("; ")}]` : "";
}

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
      const attachmentSummary = formatAttachmentSummary(message);
      if (attachmentSummary) {
        lines.push(`    ${attachmentSummary}`);
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
      const attachmentSummary = formatAttachmentSummary(message);
      if (attachmentSummary) {
        lines.push(attachmentSummary, "");
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
      const attachmentSummary = formatAttachmentSummary(message);
      if (attachmentSummary) {
        lines.push(`      ${attachmentSummary}`);
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
    "Content is interpreted as Markdown by default and converted to rich HTML. " +
    "To send images, provide file paths via --image. " +
    "To send file attachments (documents, videos, etc.), provide file paths via --file. " +
    "Use [image] placeholders in --content to interleave text and images " +
    "(e.g. --content 'Before [image] after [image] done' --image a.png --image b.png). " +
    "Without placeholders, text appears first followed by all images.",
  parameters: [
    ...conversationParameters,
    {
      name: "content",
      type: "string",
      description:
        "Message content to send. " +
        "Use [image] placeholders to position images within the text.",
      required: false,
    },
    {
      name: "image",
      type: "string[]",
      description: "Image file path(s) to attach inline (repeatable)",
      required: false,
    },
    {
      name: "file",
      type: "string[]",
      description:
        "File path(s) to attach as SharePoint-hosted file attachments (repeatable). " +
        "Supports any file type (documents, videos, archives, etc.).",
      required: false,
    },
    {
      name: "messageFormat",
      type: "string",
      description:
        'Content format: "markdown" (default, converted to HTML), "html" (raw HTML), or "text" (plain text). ' +
        "Ignored when images or files are provided (always uses HTML).",
      required: false,
      default: "markdown",
    },
    {
      name: "scheduleAt",
      type: "string",
      description:
        "Schedule the message to be sent at a future time. " +
        "Accepts an ISO 8601 timestamp (e.g. 2025-01-15T14:30:00Z). " +
        "Cannot be combined with --image or --file attachments.",
      required: false,
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId, label } = await resolveConversationId(
      client,
      parameters,
    );
    const content = (parameters.content as string | undefined) ?? "";
    const imagePaths = (parameters.image as string[] | undefined) ?? [];
    const filePaths = (parameters.file as string[] | undefined) ?? [];
    const scheduleAtRaw = parameters.scheduleAt as string | undefined;

    if (!content && imagePaths.length === 0 && filePaths.length === 0) {
      throw new Error(
        "At least one of --content, --image, or --file must be provided",
      );
    }

    if (
      scheduleAtRaw &&
      (imagePaths.length > 0 || filePaths.length > 0)
    ) {
      throw new Error(
        "Scheduled messages cannot include --image or --file attachments",
      );
    }

    if (scheduleAtRaw) {
      const scheduleAt = new Date(scheduleAtRaw);
      if (isNaN(scheduleAt.getTime())) {
        throw new Error(
          `Invalid --scheduleAt value: "${scheduleAtRaw}". Use an ISO 8601 timestamp (e.g. 2025-01-15T14:30:00Z).`,
        );
      }
      if (scheduleAt.getTime() <= Date.now()) {
        throw new Error(
          "Scheduled time must be in the future",
        );
      }
      const messageFormat =
        (parameters.messageFormat as MessageFormat | undefined) ?? "markdown";
      const result = await client.scheduleMessage(
        conversationId,
        content,
        scheduleAt,
        messageFormat,
      );
      return { ...result, conversation: label, scheduled: true };
    }

    if (filePaths.length > 0) {
      const contentParts = buildContentParts(content, imagePaths, filePaths);
      const result = await client.sendMessageWithFiles(
        conversationId,
        contentParts,
      );
      return { ...result, conversation: label };
    }

    if (imagePaths.length > 0) {
      const contentParts = buildContentParts(content, imagePaths);
      const result = await client.sendMessageWithImages(
        conversationId,
        contentParts,
      );
      return { ...result, conversation: label };
    }

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
    const { messageId, arrivalTime, conversation, scheduled, scheduledTime } =
      result as {
        messageId: string;
        arrivalTime: number;
        conversation: string;
        scheduled?: boolean;
        scheduledTime?: string;
      };
    if (scheduled) {
      return [
        `Message scheduled for "${conversation}"`,
        `  Message ID: ${messageId}`,
        `  Scheduled for: ${scheduledTime}`,
      ].join("\n");
    }
    return [
      `Message sent to "${conversation}"`,
      `  Message ID: ${messageId}`,
      `  Arrival time: ${arrivalTime}`,
    ].join("\n");
  },
  formatMarkdown: (result) => {
    const { messageId, arrivalTime, conversation, scheduled, scheduledTime } =
      result as {
        messageId: string;
        arrivalTime: number;
        conversation: string;
        scheduled?: boolean;
        scheduledTime?: string;
      };
    if (scheduled) {
      return [
        "## Message Scheduled",
        "",
        `- **To:** ${conversation}`,
        `- **Message ID:** ${messageId}`,
        `- **Scheduled for:** ${scheduledTime}`,
      ].join("\n");
    }
    return [
      "## Message Sent",
      "",
      `- **To:** ${conversation}`,
      `- **Message ID:** ${messageId}`,
      `- **Arrival time:** ${arrivalTime}`,
    ].join("\n");
  },
  formatToon: (result) => {
    const { messageId, arrivalTime, conversation, scheduled, scheduledTime } =
      result as {
        messageId: string;
        arrivalTime: number;
        conversation: string;
        scheduled?: boolean;
        scheduledTime?: string;
      };
    if (scheduled) {
      return [
        toonHeader("📅", "Message Scheduled!"),
        `  📨 To: "${conversation}"`,
        `  🆔 ${messageId}`,
        `  ⏰ Scheduled for: ${scheduledTime}`,
      ].join("\n");
    }
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
