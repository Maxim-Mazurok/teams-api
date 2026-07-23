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
 *     TEAMS_IMAGE_DESCRIPTION_API_KEY — Optional vision API key for teams_describe_image
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
import { recordToolCall, recordToolError } from "./telemetry.js";
import { buildDownloadContentBlocks } from "./model-context-protocol-download-content.js";
import {
  AuthenticationInProgressError,
  McpAuthManager,
  NeedsEmailError,
} from "./mcp-auth.js";
import type { AuthLogFunction } from "./types.js";

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

function formatLogMessage(messages: unknown[]): string {
  return messages
    .map((message) =>
      message instanceof Error ? message.message : String(message),
    )
    .join(" ");
}

function createMcpAuthLogFunction(mcpServer: McpServer): AuthLogFunction {
  return (...messages: unknown[]) => {
    const data = formatLogMessage(messages);
    void mcpServer
      .sendLoggingMessage({
        level: "info",
        logger: "teams-api.auth",
        data,
      })
      .catch((error: unknown) => {
        if (process.env.TEAMS_TELEMETRY === "true") {
          console.error(
            "Failed to send MCP auth logging message:",
            error instanceof Error ? error.message : error,
          );
        }
      });
  };
}

const authManager = new McpAuthManager<TeamsClient>({
  clientFactory: TeamsClient,
  log: createMcpAuthLogFunction(server),
});

// ── Register all actions as MCP tools ─────────────────────────────────

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
        const client = await authManager.getClient(
          parameters.email as string | undefined,
        );
        const result = await action.execute(
          client,
          parameters as Record<string, unknown>,
        );

        const output = formatOutput(action, result, outputFormat);
        const durationMs = Date.now() - start;

        // Redact binary data before recording telemetry — Buffer serialises as
        // a large numeric array and would massively bloat telemetry.jsonl.
        const telemetryResult =
          action.name === "download-file" && Array.isArray(result)
            ? (result as DownloadResult[]).map(({ data, ...rest }) => ({
                ...rest,
                byteLength: data.byteLength,
              }))
            : result;

        recordToolCall({
          tool: action.name,
          format: outputFormat,
          parameters: parameters as Record<string, unknown>,
          result: telemetryResult,
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
        if (
          error instanceof NeedsEmailError ||
          error instanceof AuthenticationInProgressError
        ) {
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
  void authManager.authenticateOnStartup();
}

main().catch((error: Error) => {
  console.error("MCP server error:", error.message);
  process.exit(1);
});
