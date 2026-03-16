#!/usr/bin/env npx tsx
/**
 * teams-api CLI
 *
 * Full-featured command-line interface for Teams operations.
 *
 * Commands:
 *   conversations   List conversations
 *   messages         Get messages from a conversation
 *   send             Send a message
 *   members          List members of a conversation
 *   auth             Acquire and print a token
 */

import { Command } from "commander";
import { TeamsClient } from "./teams-client.js";
import type { AutoLoginOptions, ManualTokenOptions } from "./types.js";

const program = new Command();

program
  .name("teams-api")
  .description("AI-native Microsoft Teams integration CLI")
  .version("0.1.0");

// ── Shared auth options ────────────────────────────────────────────────

function addAuthOptions(command: Command): Command {
  return command
    .option("--auto", "Auto-acquire token via FIDO2 passkey")
    .option("--email <email>", "Corporate email for auto login")
    .option("--token <token>", "Use an existing skype token directly")
    .option(
      "--debug-port <port>",
      "Chrome debug port for manual token capture",
      "9222",
    )
    .option("--region <region>", "API region", "apac");
}

interface AuthFlags {
  auto?: boolean;
  email?: string;
  token?: string;
  debugPort: string;
  region: string;
}

async function createClient(flags: AuthFlags): Promise<TeamsClient> {
  if (flags.token) {
    return TeamsClient.fromToken(flags.token, flags.region);
  }

  if (flags.auto) {
    if (!flags.email) {
      console.error("Error: --email is required when using --auto");
      process.exit(1);
    }
    const autoLoginOptions: AutoLoginOptions = {
      email: flags.email,
      headless: true,
      verbose: true,
    };
    return TeamsClient.fromAutoLogin(autoLoginOptions);
  }

  const manualOptions: ManualTokenOptions = {
    debugPort: Number(flags.debugPort),
  };
  return TeamsClient.fromDebugSession(manualOptions);
}

// ── auth command ───────────────────────────────────────────────────────

addAuthOptions(
  program
    .command("auth")
    .description("Acquire a Teams token and print it to stdout"),
).action(async (flags: AuthFlags) => {
  const client = await createClient(flags);
  const token = client.getToken();
  console.log(JSON.stringify(token, null, 2));
});

// ── conversations command ─────────────────────────────────────────────

addAuthOptions(
  program
    .command("conversations")
    .description("List conversations")
    .option("--limit <n>", "Max conversations to return", "50")
    .option("--json", "Output as JSON"),
).action(async (flags: AuthFlags & { limit: string; json?: boolean }) => {
  const client = await createClient(flags);
  const conversations = await client.listConversations({
    pageSize: Number(flags.limit),
  });

  if (flags.json) {
    console.log(JSON.stringify(conversations, null, 2));
    return;
  }

  console.log(`\n${conversations.length} conversations:\n`);
  for (let i = 0; i < conversations.length; i++) {
    const conversation = conversations[i];
    const lastMessage = conversation.lastMessageTime?.slice(0, 10) ?? "unknown";
    const topic = conversation.topic || "(untitled 1:1 chat)";
    console.log(
      `  [${i}] ${conversation.threadType}: "${topic}" (members: ${conversation.memberCount ?? "?"}, last: ${lastMessage})`,
    );
  }
});

// ── messages command ──────────────────────────────────────────────────

addAuthOptions(
  program
    .command("messages")
    .description("Get messages from a conversation")
    .requiredOption("--chat <name>", "Conversation name/topic (partial match)")
    .option("--max-pages <n>", "Max pagination pages", "100")
    .option("--page-size <n>", "Messages per page", "200")
    .option(
      "--text-only",
      "Only show text/rich-text messages (exclude system events)",
    )
    .option("--json", "Output as JSON"),
).action(
  async (
    flags: AuthFlags & {
      chat: string;
      maxPages: string;
      pageSize: string;
      textOnly?: boolean;
      json?: boolean;
    },
  ) => {
    const client = await createClient(flags);

    const conversation = await client.findConversation(flags.chat);
    if (!conversation) {
      console.error(`No conversation matching "${flags.chat}" found.`);
      process.exit(1);
    }

    console.error(`Fetching messages from "${conversation.topic}"...`);

    let messages = await client.getMessages(conversation.id, {
      maxPages: Number(flags.maxPages),
      pageSize: Number(flags.pageSize),
      onProgress: (count) =>
        process.stderr.write(`\r  ${count} messages fetched...`),
    });
    process.stderr.write("\n");

    if (flags.textOnly) {
      messages = messages.filter(
        (message) =>
          (message.messageType === "RichText/Html" ||
            message.messageType === "Text") &&
          !message.isDeleted,
      );
    }

    if (flags.json) {
      console.log(JSON.stringify(messages, null, 2));
      return;
    }

    console.error(`\n${messages.length} messages:\n`);
    for (const message of messages) {
      const time = message.originalArrivalTime.slice(0, 19).replace("T", " ");
      const sender = message.senderDisplayName || "(system)";
      const preview = message.content.replace(/<[^>]*>/g, "").slice(0, 120);
      console.log(`  [${time}] ${sender}: ${preview}`);
    }
  },
);

// ── send command ──────────────────────────────────────────────────────

addAuthOptions(
  program
    .command("send")
    .description("Send a message to a conversation")
    .option("--to <name>", "Find 1:1 conversation by person name")
    .option("--chat <name>", "Find conversation by topic name")
    .requiredOption("--message <text>", "Message content (plain text)")
    .option("--json", "Output result as JSON"),
).action(
  async (
    flags: AuthFlags & {
      to?: string;
      chat?: string;
      message: string;
      json?: boolean;
    },
  ) => {
    if (!flags.to && !flags.chat) {
      console.error("Error: either --to <name> or --chat <name> is required.");
      process.exit(1);
    }

    const client = await createClient(flags);

    let conversationId: string;
    let conversationLabel: string;

    if (flags.to) {
      console.error(`Searching for 1:1 conversation with "${flags.to}"...`);
      const result = await client.findOneOnOneConversation(flags.to);
      if (!result) {
        console.error(`No 1:1 conversation found matching "${flags.to}".`);
        process.exit(1);
      }
      conversationId = result.conversationId;
      conversationLabel = result.memberDisplayName;
    } else {
      console.error(`Searching for conversation matching "${flags.chat}"...`);
      const conversation = await client.findConversation(flags.chat!);
      if (!conversation) {
        console.error(`No conversation matching "${flags.chat}" found.`);
        process.exit(1);
      }
      conversationId = conversation.id;
      conversationLabel = conversation.topic;
    }

    console.error(`Sending to "${conversationLabel}"...`);
    const result = await client.sendMessage(conversationId, flags.message);

    if (flags.json) {
      console.log(
        JSON.stringify({ ...result, conversation: conversationLabel }, null, 2),
      );
      return;
    }

    console.log(`Message sent to "${conversationLabel}"`);
    console.log(`  Message ID: ${result.messageId}`);
    console.log(`  Arrival time: ${result.arrivalTime}`);
  },
);

// ── members command ───────────────────────────────────────────────────

addAuthOptions(
  program
    .command("members")
    .description("List members of a conversation")
    .requiredOption("--chat <name>", "Conversation name/topic (partial match)")
    .option("--json", "Output as JSON"),
).action(async (flags: AuthFlags & { chat: string; json?: boolean }) => {
  const client = await createClient(flags);

  const conversation = await client.findConversation(flags.chat);
  if (!conversation) {
    console.error(`No conversation matching "${flags.chat}" found.`);
    process.exit(1);
  }

  const members = await client.getMembers(conversation.id);

  if (flags.json) {
    console.log(JSON.stringify(members, null, 2));
    return;
  }

  console.log(`\n${members.length} members in "${conversation.topic}":\n`);
  for (const member of members) {
    const name = member.displayName || "(unknown)";
    console.log(`  ${name} (${member.role}) — ${member.id}`);
  }
});

program.parseAsync().catch((error: Error) => {
  console.error("Fatal error:", error.message);
  process.exit(1);
});
