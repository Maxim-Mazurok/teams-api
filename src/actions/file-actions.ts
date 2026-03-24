/**
 * File download action definitions.
 *
 * Actions: download-file.
 */

import { writeFileSync } from "node:fs";
import { resolve, join } from "node:path";
import { type ActionDefinition, toonHeader } from "./formatters.js";
import {
  conversationParameters,
  resolveConversationId,
} from "./conversation-resolution.js";

interface DownloadResult {
  fileName: string;
  fileType: string;
  size: number;
  contentType: string;
  savedTo: string | null;
}

export const downloadFileAction: ActionDefinition = {
  name: "download-file",
  title: "Download File",
  description:
    "Download file attachment(s) from a message. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "Specify the message containing the file(s) via --message-id. " +
    "Use --output-directory to save files to disk. " +
    "Without --output-directory, returns file metadata only.",
  parameters: [
    ...conversationParameters,
    {
      name: "messageId",
      type: "string",
      description: "ID of the message containing file attachment(s)",
      required: true,
    },
    {
      name: "outputDirectory",
      type: "string",
      description: "Directory to save downloaded files to",
      required: false,
    },
  ],
  execute: async (client, parameters) => {
    const { conversationId } = await resolveConversationId(client, parameters);
    const messageId = parameters.messageId as string;
    const outputDirectory = parameters.outputDirectory as string | undefined;

    // Fetch messages to find the target one
    const messages = await client.getMessages(conversationId);
    const message = messages.find((message) => message.id === messageId);
    if (!message) {
      throw new Error(
        `Message ${messageId} not found in conversation ${conversationId}`,
      );
    }

    if (message.files.length === 0 && message.images.length === 0) {
      throw new Error(`Message ${messageId} has no file attachments or images`);
    }

    const results: DownloadResult[] = [];

    // Download file attachments (SharePoint)
    for (const file of message.files) {
      const { data, contentType, size, fileName } = await client.downloadFile(
        file.fileUrl,
        file.itemId,
      );

      let savedTo: string | null = null;
      if (outputDirectory) {
        const outputPath = resolve(join(outputDirectory, fileName));
        writeFileSync(outputPath, data);
        savedTo = outputPath;
      }

      results.push({
        fileName,
        fileType: file.fileType,
        size,
        contentType,
        savedTo,
      });
    }

    // Download inline images (AMS)
    for (const image of message.images) {
      const { data, contentType, size } = await client.downloadImage(
        image.amsObjectId,
        true,
      );

      const fileName = `${image.amsObjectId}.jpg`;
      let savedTo: string | null = null;
      if (outputDirectory) {
        const outputPath = resolve(join(outputDirectory, fileName));
        writeFileSync(outputPath, data);
        savedTo = outputPath;
      }

      results.push({
        fileName,
        fileType: "image",
        size,
        contentType,
        savedTo,
      });
    }

    return results;
  },
  formatResult: (result) => {
    const downloads = result as DownloadResult[];
    const lines = [`Downloaded ${downloads.length} file(s):`];
    for (const download of downloads) {
      lines.push(
        `  ${download.fileName} (${download.fileType}, ${download.size} bytes)`,
      );
      if (download.savedTo) {
        lines.push(`    Saved to: ${download.savedTo}`);
      }
    }
    return lines.join("\n");
  },
  formatMarkdown: (result) => {
    const downloads = result as DownloadResult[];
    const lines = [`## Downloaded ${downloads.length} file(s)`, ""];
    for (const download of downloads) {
      lines.push(
        `- **${download.fileName}** (${download.fileType}, ${download.size} bytes)`,
      );
      if (download.savedTo) {
        lines.push(`  - Saved to: \`${download.savedTo}\``);
      }
    }
    return lines.join("\n");
  },
  formatToon: (result) => {
    const downloads = result as DownloadResult[];
    const lines = [toonHeader("📥", `${downloads.length} File(s) Downloaded`)];
    for (const download of downloads) {
      lines.push("");
      lines.push(`  📄 ${download.fileName}`);
      lines.push(`     ${download.fileType} · ${download.size} bytes`);
      if (download.savedTo) {
        lines.push(`     💾 ${download.savedTo}`);
      }
    }
    return lines.join("\n");
  },
};
