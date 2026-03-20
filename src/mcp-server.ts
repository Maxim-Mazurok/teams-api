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
 *     TEAMS_EMAIL           — Corporate email (optional; enables token caching and auto-login on macOS)
 *     TEAMS_AUTO            — Set to "true" to use auto-login (macOS + FIDO2)
 *     TEAMS_LOGIN           — Set to "true" to use interactive browser login (all platforms)
 *     TEAMS_DEBUG           — Set to "true" to use Chrome debug session (requires Chrome with --remote-debugging-port)
 *     TEAMS_DEBUG_PORT      — Chrome debug port (default: 9222)
 *
 * Default (no env vars): Smart login — tries auto-login on macOS if Chrome + email
 * are available, falls back to interactive browser login on all platforms.
 *
 * Usage in VS Code settings (mcp config):
 *   {
 *     "mcpServers": {
 *       "teams": {
 *         "command": "npx",
 *         "args": ["-y", "teams-api"],
 *         "env": {}
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
  } else if (envAuto) {
    if (!email) {
      throw new NeedsEmailError();
    }
    clientInstance = await TeamsClient.create({
      email,
      region: envRegion,
      headless: true,
      verbose: false,
    });
  } else if (envLogin) {
    clientInstance = await TeamsClient.fromInteractiveLogin({
      region: envRegion,
      email,
      verbose: false,
    });
  } else if (envDebug) {
    clientInstance = await TeamsClient.fromDebugSession({
      debugPort: envDebugPort,
      region: envRegion,
    });
  } else {
    // Default: smart login (zero-config)
    clientInstance = await TeamsClient.connect({
      email,
      region: envRegion,
      verbose: false,
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
      try {
        const client = await getClient(parameters.email as string | undefined);
        const outputFormat = (parameters.format as OutputFormat) ?? "toon";
        const result = await action.execute(
          client,
          parameters as Record<string, unknown>,
        );

        const structuredContent =
          result !== null &&
          typeof result === "object" &&
          !Array.isArray(result)
            ? (result as Record<string, unknown>)
            : { data: result };

        return {
          content: [
            {
              type: "text" as const,
              text: formatOutput(action, result, outputFormat),
            },
          ],
          structuredContent,
        };
      } catch (error) {
        if (error instanceof NeedsEmailError) {
          return {
            content: [{ type: "text" as const, text: error.message }],
            isError: true,
          };
        }
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
