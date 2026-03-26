/**
 * Action registry — the single source of truth for all Teams API operations.
 *
 * CLI commands, MCP tools, and programmatic usage all derive from these
 * definitions. Individual action definitions live in domain-specific files;
 * this module assembles them into the canonical registry.
 */

import type { ActionDefinition } from "./formatters.js";
import {
  listConversations,
  findConversation,
  findOneOnOne,
} from "./conversation-actions.js";
import {
  getMessages,
  sendMessage,
  editMessageAction,
  deleteMessageAction,
  addReactionAction,
  removeReactionAction,
} from "./message-actions.js";
import { findPeopleAction, findChatsAction } from "./search-actions.js";
import { getMembers, whoami, getTranscript } from "./utility-actions.js";
import { downloadFileAction } from "./file-actions.js";

// ── Registry ─────────────────────────────────────────────────────────

/**
 * Map-based action registry keyed by action name.
 * Ensures compile-time visibility of all actions and prevents
 * accidental omissions from the exported array.
 */
const actionRegistry = new Map<string, ActionDefinition>([
  ["list-conversations", listConversations],
  ["find-conversation", findConversation],
  ["find-one-on-one", findOneOnOne],
  ["find-people", findPeopleAction],
  ["find-chats", findChatsAction],
  ["get-messages", getMessages],
  ["send-message", sendMessage],
  ["edit-message", editMessageAction],
  ["delete-message", deleteMessageAction],
  ["add-reaction", addReactionAction],
  ["remove-reaction", removeReactionAction],
  ["get-members", getMembers],
  ["whoami", whoami],
  ["get-transcript", getTranscript],
  ["download-file", downloadFileAction],
]);

/** All registered actions, derived from the registry map. */
export const actions: ActionDefinition[] = Array.from(actionRegistry.values());
