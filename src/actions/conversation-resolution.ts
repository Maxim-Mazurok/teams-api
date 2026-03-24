/**
 * Shared conversation resolution logic for actions.
 *
 * Provides the standard parameter definitions for identifying a conversation
 * and a resolver function that supports three identification strategies:
 * direct ID, topic name match, or person name (1:1 lookup).
 */

import type { TeamsClient } from "../teams-client.js";
import type { ActionParameter } from "./formatters.js";

/** Shared parameter definitions for conversation identification. */
export const conversationParameters: ActionParameter[] = [
  {
    name: "chat",
    type: "string",
    description:
      "Find conversation by topic name (partial match), person name (1:1 fallback), or direct thread ID",
    required: false,
  },
  {
    name: "to",
    type: "string",
    description: "Find 1:1 conversation by person name",
    required: false,
  },
  {
    name: "conversationId",
    type: "string",
    description: "Direct conversation thread ID",
    required: false,
  },
];

/**
 * Resolve a conversation ID from the standard identification parameters.
 *
 * Supports three ways to identify a conversation:
 *   1. conversationId — direct thread ID
 *   2. chat — topic name (partial match via findConversation)
 *   3. to — person name (1:1 lookup via findOneOnOneConversation)
 *
 * Returns both the resolved ID and a human-readable label.
 */
export async function resolveConversationId(
  client: TeamsClient,
  parameters: Record<string, unknown>,
): Promise<{ conversationId: string; label: string }> {
  const conversationId = parameters.conversationId as string | undefined;
  const chat = parameters.chat as string | undefined;
  const to = parameters.to as string | undefined;

  if (conversationId) {
    return { conversationId, label: conversationId };
  }

  if (chat) {
    // If the value looks like a raw conversation ID, use it directly
    if (chat.startsWith("19:") && chat.includes("@")) {
      return { conversationId: chat, label: chat };
    }

    const conversation = await client.findConversation(chat);
    if (conversation) {
      return { conversationId: conversation.id, label: conversation.topic };
    }

    // Fall back to 1:1 resolution (covers person names passed to --chat)
    const oneOnOne = await client.findOneOnOneConversation(chat);
    if (oneOnOne) {
      return {
        conversationId: oneOnOne.conversationId,
        label: oneOnOne.memberDisplayName,
      };
    }

    throw new Error(`No conversation matching "${chat}" found.`);
  }

  if (to) {
    const result = await client.findOneOnOneConversation(to);
    if (!result) {
      throw new Error(`No 1:1 conversation found with "${to}".`);
    }
    return {
      conversationId: result.conversationId,
      label: result.memberDisplayName,
    };
  }

  throw new Error("One of --conversation-id, --chat, or --to is required.");
}
