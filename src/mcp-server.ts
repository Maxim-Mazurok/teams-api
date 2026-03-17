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
 *     TEAMS_TOKEN     — Use an existing skype token
 *     TEAMS_REGION    — API region (default: "apac")
 *     TEAMS_EMAIL     — Corporate email (for auto-login or interactive login)
 *     TEAMS_AUTO      — Set to "true" to use auto-login (macOS + FIDO2)
 *     TEAMS_LOGIN     — Set to "true" to use interactive browser login (all platforms)
 *     TEAMS_DEBUG_PORT — Chrome debug port (default: 9222)
 *
 * Usage in VS Code settings (mcp config):
 *   {
 *     "mcpServers": {
 *       "teams": {
 *         "command": "npx",
 *         "args": ["-y", "tsx", "/path/to/teams-api/src/mcp-server.ts"],
 *         "env": {
 *           "TEAMS_AUTO": "true",
 *           "TEAMS_EMAIL": "user@company.com"
 *         }
 *       }
 *     }
 *   }
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { TeamsClient } from "./teams-client.js";
import { actions, formatOutput } from "./actions.js";
import type { ActionParameter, OutputFormat } from "./actions.js";

let clientInstance: TeamsClient | null = null;

async function getClient(): Promise<TeamsClient> {
  if (clientInstance) {
    return clientInstance;
  }

  const envToken = process.env.TEAMS_TOKEN;
  const envRegion = process.env.TEAMS_REGION ?? "apac";
  const envEmail = process.env.TEAMS_EMAIL;
  const envAuto = process.env.TEAMS_AUTO === "true";
  const envLogin = process.env.TEAMS_LOGIN === "true";
  const envDebugPort = process.env.TEAMS_DEBUG_PORT
    ? Number(process.env.TEAMS_DEBUG_PORT)
    : 9222;

  if (envToken) {
    clientInstance = TeamsClient.fromToken(envToken, envRegion);
  } else if (envAuto && envEmail) {
    clientInstance = await TeamsClient.create({
      email: envEmail,
      headless: true,
      verbose: false,
    });
  } else if (envLogin) {
    clientInstance = await TeamsClient.fromInteractiveLogin({
      region: envRegion,
      email: envEmail,
      verbose: false,
    });
  } else {
    clientInstance = await TeamsClient.fromDebugSession({
      debugPort: envDebugPort,
    });
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
  }
  schema = schema.describe(parameter.description);
  if (!parameter.required) {
    schema = schema.optional();
  }
  return schema;
}

const server = new McpServer({
  name: "teams-api",
  version: "0.1.0",
});

// ── Register all actions as MCP tools ─────────────────────────────────

for (const action of actions) {
  const toolName = `teams_${action.name.replace(/-/g, "_")}`;

  const inputSchema: Record<string, z.ZodTypeAny> = {};
  for (const parameter of action.parameters) {
    inputSchema[parameter.name] = parameterToZod(parameter);
  }
  inputSchema["format"] = z
    .enum(["json", "text", "md", "toon"])
    .describe("Output format (default: toon)")
    .optional();

  server.registerTool(
    toolName,
    {
      title: action.title,
      description: action.description,
      inputSchema,
    },
    async (parameters) => {
      const client = await getClient();
      const outputFormat = (parameters.format as OutputFormat) ?? "toon";
      const result = await action.execute(
        client,
        parameters as Record<string, unknown>,
      );

      return {
        content: [
          {
            type: "text" as const,
            text: formatOutput(action, result, outputFormat),
          },
        ],
      };
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
