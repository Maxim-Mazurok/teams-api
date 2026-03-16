#!/usr/bin/env npx tsx
/**
 * MCP Server for Teams API
 *
 * Exposes Teams operations as MCP tools for AI agents.
 * Communicates via stdio transport.
 *
 * Configuration:
 *   Set environment variables for authentication:
 *     TEAMS_TOKEN     — Use an existing skype token
 *     TEAMS_REGION    — API region (default: "apac")
 *     TEAMS_EMAIL     — Corporate email (for auto-login)
 *     TEAMS_AUTO      — Set to "true" to use auto-login
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

let clientInstance: TeamsClient | null = null;

async function getClient(): Promise<TeamsClient> {
  if (clientInstance) {
    return clientInstance;
  }

  const envToken = process.env.TEAMS_TOKEN;
  const envRegion = process.env.TEAMS_REGION ?? "apac";
  const envEmail = process.env.TEAMS_EMAIL;
  const envAuto = process.env.TEAMS_AUTO === "true";
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
  } else {
    clientInstance = await TeamsClient.fromDebugSession({
      debugPort: envDebugPort,
    });
  }

  return clientInstance;
}

const server = new McpServer({
  name: "teams-api",
  version: "0.1.0",
});

server.registerTool(
  "teams_list_conversations",
  {
    title: "List Teams Conversations",
    description:
      "List Microsoft Teams conversations (chats, group chats, meetings, channels). " +
      "Returns conversation ID, topic/name, type, member count, and last message time. " +
      "Use the conversation ID from the results to fetch messages or send messages.",
    inputSchema: {
      limit: z
        .number()
        .optional()
        .describe("Maximum number of conversations to return (default: 50)"),
    },
  },
  async ({ limit }) => {
    const client = await getClient();
    const conversations = await client.listConversations({
      pageSize: limit ?? 50,
    });

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(conversations, null, 2),
        },
      ],
    };
  },
);

server.registerTool(
  "teams_find_conversation",
  {
    title: "Find Teams Conversation",
    description:
      "Find a Teams conversation by topic name (case-insensitive partial match). " +
      "Returns the matching conversation with its ID, or null if not found. " +
      "For 1:1 chats (which have no topic), use teams_find_one_on_one instead.",
    inputSchema: {
      query: z
        .string()
        .describe("Partial topic name to search for (case-insensitive)"),
    },
  },
  async ({ query }) => {
    const client = await getClient();
    const conversation = await client.findConversation(query);

    return {
      content: [
        {
          type: "text" as const,
          text: conversation
            ? JSON.stringify(conversation, null, 2)
            : "No conversation found matching the query.",
        },
      ],
    };
  },
);

server.registerTool(
  "teams_find_one_on_one",
  {
    title: "Find 1:1 Conversation",
    description:
      "Find a 1:1 Teams conversation with a specific person by name. " +
      "Searches untitled chats by scanning recent message sender names. " +
      "Also finds the self-chat (notes to self) if the name matches the current user. " +
      "Returns the conversation ID and matched member name, or null if not found.",
    inputSchema: {
      personName: z
        .string()
        .describe(
          "Name of the person to find (case-insensitive partial match)",
        ),
    },
  },
  async ({ personName }) => {
    const client = await getClient();
    const result = await client.findOneOnOneConversation(personName);

    return {
      content: [
        {
          type: "text" as const,
          text: result
            ? JSON.stringify(result, null, 2)
            : `No 1:1 conversation found with "${personName}".`,
        },
      ],
    };
  },
);

server.registerTool(
  "teams_get_messages",
  {
    title: "Get Messages",
    description:
      "Fetch messages from a Teams conversation. Returns an array of messages " +
      "with sender name, content (HTML or plain text), timestamp, reactions, " +
      "mentions, and reply references. Use the conversation ID from " +
      "teams_list_conversations or teams_find_conversation. " +
      "Messages are in reverse chronological order (newest first).",
    inputSchema: {
      conversationId: z
        .string()
        .describe("Conversation thread ID to fetch messages from"),
      maxPages: z
        .number()
        .optional()
        .describe("Max pagination pages to fetch (default: 5 for MCP use)"),
      textOnly: z
        .boolean()
        .optional()
        .describe(
          "If true, exclude system events and only return text messages (default: true)",
        ),
    },
  },
  async ({ conversationId, maxPages, textOnly }) => {
    const client = await getClient();
    let messages = await client.getMessages(conversationId, {
      maxPages: maxPages ?? 5,
      pageSize: 200,
    });

    const shouldFilterText = textOnly ?? true;
    if (shouldFilterText) {
      messages = messages.filter(
        (message) =>
          (message.messageType === "RichText/Html" ||
            message.messageType === "Text") &&
          !message.isDeleted,
      );
    }

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(messages, null, 2),
        },
      ],
    };
  },
);

server.registerTool(
  "teams_send_message",
  {
    title: "Send Message",
    description:
      "Send a plain-text message to a Teams conversation. " +
      "The sender identity is determined automatically from the authenticated user. " +
      "Use the conversation ID from teams_list_conversations, " +
      "teams_find_conversation, or teams_find_one_on_one.",
    inputSchema: {
      conversationId: z
        .string()
        .describe("Conversation thread ID to send the message to"),
      content: z.string().describe("Plain text message content to send"),
    },
  },
  async ({ conversationId, content }) => {
    const client = await getClient();
    const result = await client.sendMessage(conversationId, content);

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  },
);

server.registerTool(
  "teams_get_members",
  {
    title: "Get Conversation Members",
    description:
      "List members of a Teams conversation. Returns member ID (MRI), " +
      "display name, and role. Note: 1:1 chat members may have empty display names; " +
      "use teams_find_one_on_one to resolve names from message history.",
    inputSchema: {
      conversationId: z
        .string()
        .describe("Conversation thread ID to get members for"),
    },
  },
  async ({ conversationId }) => {
    const client = await getClient();
    const members = await client.getMembers(conversationId);

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(members, null, 2),
        },
      ],
    };
  },
);

server.registerTool(
  "teams_whoami",
  {
    title: "Current User Info",
    description:
      "Get the display name of the currently authenticated Teams user.",
    inputSchema: {},
  },
  async () => {
    const client = await getClient();
    const displayName = await client.getCurrentUserDisplayName();
    const token = client.getToken();

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify({ displayName, region: token.region }, null, 2),
        },
      ],
    };
  },
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error: Error) => {
  console.error("MCP server error:", error.message);
  process.exit(1);
});
