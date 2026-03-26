/**
 * Shared formatting utilities for action output.
 *
 * Contains HTML cleaning, quote extraction, message formatting helpers,
 * transcript grouping, and output format dispatch.
 */

import type { Message, TranscriptEntry } from "../types.js";
import { decodeHtmlEntities } from "../html-utils.js";

// ── Parameter & Action types ─────────────────────────────────────────

export interface ActionParameter {
  /** Parameter name in camelCase (CLI flags auto-converted to kebab-case). */
  name: string;
  /** Parameter type. Determines CLI flag syntax and MCP Zod schema. */
  type: "string" | "number" | "boolean" | "string[]";
  /** Description for CLI help, MCP tool description, and documentation. */
  description: string;
  /** Whether the parameter must be provided. */
  required: boolean;
  /** Default value when parameter is omitted. */
  default?: string | number | boolean;
}

export interface ActionDefinition {
  /** Kebab-case name. CLI command name; MCP tool name is `teams_` + snake_case. */
  name: string;
  /** Human-readable title (MCP tool title). */
  title: string;
  /** Full description shared across CLI help, MCP, and documentation. */
  description: string;
  /** Typed parameter definitions. */
  parameters: ActionParameter[];
  /** Execute the action against a TeamsClient. */
  execute: (
    client: import("../teams-client.js").TeamsClient,
    parameters: Record<string, unknown>,
  ) => Promise<unknown>;
  /** Format result as concise Markdown — human-readable, action-complete, with actionable IDs. */
  formatConcise: (result: unknown) => string;
}

export type OutputFormat = "concise" | "detailed";
export type MessageOrder = "newest-first" | "oldest-first";

// ── Output format dispatch ───────────────────────────────────────────

/** Format an action result in the specified output format. */
export function formatOutput(
  action: ActionDefinition,
  result: unknown,
  format: OutputFormat = "concise",
): string {
  switch (format) {
    case "detailed":
      return JSON.stringify(result, null, 2);
    case "concise":
      return action.formatConcise(result);
  }
}

// ── Shared formatting helpers ────────────────────────────────────────
/** Strip HTML tags and decode entities from message content. */
export function cleanContent(content: string): string {
  return decodeHtmlEntities(content.replace(/<[^>]*>/g, "")).trim();
}

/** Extract quoted text from HTML blockquotes, returning quote and body separately. */
export function extractQuote(content: string): {
  quote: string | null;
  body: string;
} {
  for (const tag of ["blockquote", "quote"]) {
    const pattern = new RegExp(`<${tag}[^>]*>[\\s\\S]*?<\\/${tag}>`, "i");
    const match = content.match(pattern);
    if (match) {
      const quote = cleanContent(match[0]);
      const remainder = content.replace(pattern, "");
      return { quote: quote || null, body: cleanContent(remainder) };
    }
  }
  return { quote: null, body: cleanContent(content) };
}

/** Build a map from message ID to sender display name. */
export function buildSenderLookup(messages: Message[]): Map<string, string> {
  const lookup = new Map<string, string>();
  for (const message of messages) {
    lookup.set(message.id, message.senderDisplayName || "(system)");
  }
  return lookup;
}

// ── Transcript formatting helpers ────────────────────────────────────

/** Format a VTT timestamp (HH:MM:SS.mmm) as a compact time string. */
export function formatTimestamp(timestamp: string): string {
  // Strip leading "00:" hours if zero, and trim milliseconds
  const withoutMilliseconds = timestamp.replace(/\.\d+$/, "");
  return withoutMilliseconds.replace(/^00:/, "");
}

/** Group consecutive entries by the same speaker for more readable output. */
export function groupBySpeaker(
  entries: TranscriptEntry[],
): Array<{ speaker: string; startTime: string; segments: string[] }> {
  const groups: Array<{
    speaker: string;
    startTime: string;
    segments: string[];
  }> = [];

  for (const entry of entries) {
    const lastGroup = groups[groups.length - 1];
    if (lastGroup && lastGroup.speaker === entry.speaker) {
      lastGroup.segments.push(entry.text);
    } else {
      groups.push({
        speaker: entry.speaker,
        startTime: entry.startTime,
        segments: [entry.text],
      });
    }
  }

  return groups;
}
