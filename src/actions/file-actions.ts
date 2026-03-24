/**
 * File download action definitions.
 *
 * Actions: download-file.
 */

import { mkdirSync, writeFileSync } from "node:fs";
import { resolve, join } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import { type ActionDefinition, toonHeader } from "./formatters.js";
import {
  conversationParameters,
  resolveConversationId,
} from "./conversation-resolution.js";

export interface DownloadResult {
  fileName: string;
  fileType: string;
  size: number;
  contentType: string;
  savedTo: string;
  data: Buffer;
}

export const downloadFileAction: ActionDefinition = {
  name: "download-file",
  title: "Download File",
  description:
    "Download file attachment(s) from a message. " +
    "Identify the conversation by topic name (--chat), " +
    "person name for 1:1 chats (--to), or direct ID (--conversation-id). " +
    "Specify the message containing the file(s) via --message-id. " +
    "Use --output-directory to save files to a specific directory. " +
    "Without --output-directory, files are saved to a temporary directory. " +
    "File contents are returned inline in the response.",
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

    // Determine output directory — use provided directory or create a unique temp directory
    const resolvedOutputDirectory = outputDirectory
      ? resolve(outputDirectory)
      : resolve(join(tmpdir(), `teams-download-${randomUUID()}`));
    mkdirSync(resolvedOutputDirectory, { recursive: true });

    const results: DownloadResult[] = [];

    // Download file attachments (SharePoint)
    for (const file of message.files) {
      const { data, contentType, size, fileName } = await client.downloadFile(
        file.fileUrl,
        file.itemId,
      );

      const outputPath = resolve(join(resolvedOutputDirectory, fileName));
      writeFileSync(outputPath, data);

      results.push({
        fileName,
        fileType: file.fileType,
        size,
        contentType,
        savedTo: outputPath,
        data,
      });
    }

    // Download inline images (AMS)
    for (const image of message.images) {
      const { data, contentType, size } = await client.downloadImage(
        image.amsObjectId,
        true,
      );

      const fileName = `${image.amsObjectId}.jpg`;
      const outputPath = resolve(join(resolvedOutputDirectory, fileName));
      writeFileSync(outputPath, data);

      results.push({
        fileName,
        fileType: "image",
        size,
        contentType,
        savedTo: outputPath,
        data,
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
      lines.push(`    Saved to: ${download.savedTo}`);
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
      lines.push(`  - Saved to: \`${download.savedTo}\``);
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
      lines.push(`     💾 ${download.savedTo}`);
    }
    return lines.join("\n");
  },
};
