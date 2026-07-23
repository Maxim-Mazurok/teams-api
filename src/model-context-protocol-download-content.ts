import type { ContentBlock } from "@modelcontextprotocol/sdk/types.js";
import type { DownloadResult } from "./actions/file-actions.js";

const TEXT_MIME_PREFIXES = [
  "text/",
  "application/json",
  "application/xml",
  "application/javascript",
  "application/typescript",
  "application/x-yaml",
  "application/yaml",
  "application/toml",
];

const TEXT_FILE_EXTENSIONS = new Set([
  "md",
  "txt",
  "csv",
  "tsv",
  "json",
  "xml",
  "yaml",
  "yml",
  "toml",
  "html",
  "htm",
  "css",
  "js",
  "ts",
  "jsx",
  "tsx",
  "py",
  "rb",
  "sh",
  "bash",
  "zsh",
  "ps1",
  "bat",
  "cmd",
  "sql",
  "graphql",
  "svg",
  "log",
  "ini",
  "cfg",
  "conf",
  "env",
  "properties",
]);

function isTextContent(mimeType: string, fileName: string): boolean {
  const lowerMimeType = mimeType.toLowerCase();
  if (TEXT_MIME_PREFIXES.some((prefix) => lowerMimeType.startsWith(prefix))) {
    return true;
  }

  const extension = fileName.includes(".")
    ? fileName.split(".").pop()?.toLowerCase()
    : undefined;
  return extension !== undefined && TEXT_FILE_EXTENSIONS.has(extension);
}

export function buildDownloadContentBlocks(
  downloads: DownloadResult[],
): ContentBlock[] {
  const contentBlocks: ContentBlock[] = [];
  for (const download of downloads) {
    if (download.contentType.startsWith("image/")) {
      contentBlocks.push({
        type: "image" as const,
        data: download.data.toString("base64"),
        mimeType: download.contentType,
      });
    } else if (isTextContent(download.contentType, download.fileName)) {
      contentBlocks.push({
        type: "resource" as const,
        resource: {
          uri: `file://${download.savedTo}`,
          mimeType: download.contentType,
          text: download.data.toString("utf-8"),
        },
      });
    }
  }
  return contentBlocks;
}
