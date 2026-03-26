#!/usr/bin/env npx tsx
/**
 * MCP Server for Teams API
 *
 * Thin adapter that maps unified action definitions to MCP tools.
 * All tools, parameters, descriptions, and execution logic come
 * from `src/actions.ts` — the single source of truth.
 *
 * Configuration:
 *   Set environment variables for authentication:
 *     TEAMS_TOKEN           — Use an existing skype token
 *     TEAMS_BEARER_TOKEN    — Optional middle-tier bearer token for profile resolution
 *     TEAMS_SUBSTRATE_TOKEN — Optional Substrate bearer token for people/chat search
 *     TEAMS_REGION          — API region (required with TEAMS_TOKEN, optional otherwise)
 *     TEAMS_EMAIL           — Corporate email (optional; the server prompts the AI agent if needed)
 *     TEAMS_AUTO            — Set to "true" to use auto-login (macOS + FIDO2)
 *     TEAMS_LOGIN           — Set to "true" to use interactive browser login (all platforms)
 *     TEAMS_DEBUG_PORT      — Chrome debug port (default: 9222)
 *     TEAMS_TELEMETRY       — Set to "true" to enable full debug telemetry (contributor use)
 *
 * Usage in VS Code settings (mcp config):
 *   {
 *     "mcpServers": {
 *       "teams": {
 *         "command": "npx",
 *         "args": ["-y", "teams-api"],
 *         "env": {
 *           "TEAMS_LOGIN": "true"
 *         }
 *       }
 *     }
 *   }
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import type { ContentBlock } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { TeamsClient } from "./teams-client.js";
import { actions } from "./actions/definitions.js";
import { formatOutput } from "./actions/formatters.js";
import type { ActionParameter, OutputFormat } from "./actions/formatters.js";
import type { DownloadResult } from "./actions/file-actions.js";
import { serverInstructions } from "./server-instructions.js";
import { recordToolCall, recordToolError, recordAuth } from "./telemetry.js";

let clientInstance: TeamsClient | null = null;

class NeedsEmailError extends Error {
  constructor() {
    super(
      "I need your corporate email address to log into Teams. " +
        "Please provide your email and call this tool again.",
    );
    this.name = "NeedsEmailError";
  }
}

async function getClient(toolEmail?: string): Promise<TeamsClient> {
  if (clientInstance) {
    return clientInstance;
  }

  const envToken = process.env.TEAMS_TOKEN;
  const envBearerToken = process.env.TEAMS_BEARER_TOKEN;
  const envSubstrateToken = process.env.TEAMS_SUBSTRATE_TOKEN;
  const envRegion = process.env.TEAMS_REGION;
  const email = process.env.TEAMS_EMAIL || toolEmail;
  const envAuto = process.env.TEAMS_AUTO === "true";
  const envLogin = process.env.TEAMS_LOGIN === "true";
  const envDebug = process.env.TEAMS_DEBUG === "true";
  const envDebugPort = process.env.TEAMS_DEBUG_PORT
    ? Number(process.env.TEAMS_DEBUG_PORT)
    : 9222;

  if (envToken) {
    if (!envRegion) {
      throw new Error("TEAMS_REGION is required when TEAMS_TOKEN is set");
    }
    clientInstance = TeamsClient.fromToken(
      envToken,
      envRegion,
      envBearerToken,
      envSubstrateToken,
    );
    if (email) {
      clientInstance.setEmail(email);
    }
    recordAuth({ strategy: "token", success: true });
  } else if (envAuto) {
    if (!email) {
      throw new NeedsEmailError();
    }
    try {
      clientInstance = await TeamsClient.create({
        email,
        region: envRegion,
        headless: true,
        verbose: false,
      });
      recordAuth({ strategy: "auto", success: true });
    } catch (err) {
      recordAuth({ strategy: "auto", success: false, error: err });
      throw err;
    }
  } else if (envLogin) {
    try {
      clientInstance = await TeamsClient.fromInteractiveLogin({
        region: envRegion,
        email,
        verbose: false,
      });
      if (email) {
        clientInstance.setEmail(email);
      }
      recordAuth({ strategy: "login", success: true });
    } catch (err) {
      recordAuth({ strategy: "login", success: false, error: err });
      throw err;
    }
  } else if (envDebug) {
    try {
      clientInstance = await TeamsClient.fromDebugSession({
        debugPort: envDebugPort,
        region: envRegion,
      });
      recordAuth({ strategy: "debug", success: true });
    } catch (err) {
      recordAuth({ strategy: "debug", success: false, error: err });
      throw err;
    }
  } else {
    // Default: smart login (cross-platform, zero-config)
    try {
      clientInstance = await TeamsClient.connect({
        email,
        region: envRegion,
        verbose: false,
      });
      if (email) {
        clientInstance.setEmail(email);
      }
      recordAuth({ strategy: "auto", success: true });
    } catch (err) {
      recordAuth({ strategy: "auto", success: false, error: err });
      throw err;
    }
  }

  return clientInstance;
}

function parameterToZod(parameter: ActionParameter): z.ZodTypeAny {
  let schema: z.ZodTypeAny;
  switch (parameter.type) {
    case "string":
      schema = z.string();
      break;
    case "number":
      schema = z.number();
      break;
    case "boolean":
      schema = z.boolean();
      break;
    case "string[]":
      schema = z.array(z.string());
      break;
  }
  schema = schema.describe(parameter.description);
  if (!parameter.required) {
    schema = schema.optional();
  }
  return schema;
}

const server = new McpServer(
  {
    name: "teams-api",
    version: "0.1.0",
  },
  {
    instructions: serverInstructions,
  },
);

// ── Register all actions as MCP tools ─────────────────────────────────

/** MIME types that are safe to return as inline text to the LLM. */
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

/** File extensions that are known to be text-based. */
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
  const lowerMime = mimeType.toLowerCase();
  if (TEXT_MIME_PREFIXES.some((prefix) => lowerMime.startsWith(prefix))) {
    return true;
  }
  // Fall back to file extension when MIME type is generic (e.g. application/octet-stream)
  const extension = fileName.includes(".")
    ? fileName.split(".").pop()?.toLowerCase()
    : undefined;
  return extension !== undefined && TEXT_FILE_EXTENSIONS.has(extension);
}

/**
 * Build MCP content blocks for file download results.
 *
 * Returns the file content inline so the LLM can read it directly:
 * - Text files → EmbeddedResource with text content
 * - Images → ImageContent with base64 data
 * - Other binary → EmbeddedResource with base64 blob
 */
function buildDownloadContentBlocks(
  downloads: DownloadResult[],
): ContentBlock[] {
  const blocks: ContentBlock[] = [];
  for (const download of downloads) {
    const fileUri = `file://${download.savedTo}`;

    if (download.contentType.startsWith("image/")) {
      blocks.push({
        type: "image" as const,
        data: download.data.toString("base64"),
        mimeType: download.contentType,
      });
    } else if (isTextContent(download.contentType, download.fileName)) {
      blocks.push({
        type: "resource" as const,
        resource: {
          uri: fileUri,
          mimeType: download.contentType,
          text: download.data.toString("utf-8"),
        },
      });
    } else {
      blocks.push({
        type: "resource" as const,
        resource: {
          uri: fileUri,
          mimeType: download.contentType,
          blob: download.data.toString("base64"),
        },
      });
    }
  }
  return blocks;
}

for (const action of actions) {
  const toolName = `teams_${action.name.replace(/-/g, "_")}`;

  const inputSchema: Record<string, z.ZodTypeAny> = {};
  for (const parameter of action.parameters) {
    inputSchema[parameter.name] = parameterToZod(parameter);
  }
  inputSchema["format"] = z
    .enum(["concise", "detailed"])
    .describe(
      "Output format. " +
        '"concise" (default): light Markdown with actionable IDs and key decision fields; nested collections may be summarized. ' +
        '"detailed": full JSON for programmatic processing or inspecting exact field values.',
    )
    .optional();
  inputSchema["email"] = z
    .string()
    .describe(
      "Corporate email address for Teams login. " +
        "Only needed if the server asks for it.",
    )
    .optional();

  server.registerTool(
    toolName,
    {
      title: action.title,
      description: action.description,
      inputSchema,
    },
    async (parameters) => {
      const outputFormat = (parameters.format as OutputFormat) ?? "concise";
      const start = Date.now();

      try {
        const client = await getClient(parameters.email as string | undefined);
        const result = await action.execute(
          client,
          parameters as Record<string, unknown>,
        );

        const output = formatOutput(action, result, outputFormat);
        const durationMs = Date.now() - start;

        recordToolCall({
          tool: action.name,
          format: outputFormat,
          parameters: parameters as Record<string, unknown>,
          result,
          output,
          durationMs,
        });

        // Build content blocks — file downloads get inline file content
        const contentBlocks: ContentBlock[] = [
          { type: "text" as const, text: output },
        ];

        if (action.name === "download-file" && Array.isArray(result)) {
          const downloads = result as DownloadResult[];
          contentBlocks.push(...buildDownloadContentBlocks(downloads));
        }

        return {
          content: contentBlocks,
        };
      } catch (error) {
        const durationMs = Date.now() - start;
        if (error instanceof NeedsEmailError) {
          recordToolError({
            tool: action.name,
            format: outputFormat,
            parameters: parameters as Record<string, unknown>,
            error,
            durationMs,
          });
          return {
            content: [{ type: "text" as const, text: error.message }],
            isError: true,
          };
        }
        recordToolError({
          tool: action.name,
          format: outputFormat,
          parameters: parameters as Record<string, unknown>,
          error,
          durationMs,
        });
        throw error;
      }
    },
  );
}

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error: Error) => {
  console.error("MCP server error:", error.message);
  process.exit(1);
});
