/**
 * Search-related action definitions.
 *
 * Actions: find-people, find-chats.
 */

import type { PersonSearchResult, ChatSearchResult } from "../types.js";
import type { ActionDefinition } from "./formatters.js";

export const findPeopleAction: ActionDefinition = {
  name: "find-people",
  title: "Find People",
  description:
    "Search for people in the organization directory by name. " +
    "Uses the Substrate search API (requires authentication via auto-login or interactive). " +
    "Returns matching people with emails, job titles, and departments.",
  parameters: [
    {
      name: "query",
      type: "string",
      description: "Name or partial name to search for",
      required: true,
    },
    {
      name: "maxResults",
      type: "number",
      description: "Maximum results to return (default: 10)",
      required: false,
      default: 10,
    },
  ],
  execute: async (client, parameters) => {
    const query = parameters.query as string;
    const maxResults = (parameters.maxResults as number) ?? 10;
    return client.findPeople(query, maxResults);
  },
  formatConcise: (result) => {
    const people = result as PersonSearchResult[];
    if (people.length === 0) return "No people found.";
    const lines = [`## People (${people.length} found)`, ""];
    for (const person of people) {
      lines.push(`### ${person.displayName}`);
      lines.push(`- **Email:** ${person.email}`);
      if (person.jobTitle) lines.push(`- **Title:** ${person.jobTitle}`);
      if (person.department)
        lines.push(`- **Department:** ${person.department}`);
      lines.push(`- **MRI:** ${person.mri}`);
      if (person.objectId) lines.push(`- **Object ID:** ${person.objectId}`);
      lines.push("");
    }
    return lines.join("\n");
  },
};

export const findChatsAction: ActionDefinition = {
  name: "find-chats",
  title: "Find Chats",
  description:
    "Search for chats by name or member name. " +
    "Uses the Substrate search API (requires authentication via auto-login or interactive). " +
    "Returns matching chats with member lists and thread IDs.",
  parameters: [
    {
      name: "query",
      type: "string",
      description: "Chat name or member name to search for",
      required: true,
    },
    {
      name: "maxResults",
      type: "number",
      description: "Maximum results to return (default: 10)",
      required: false,
      default: 10,
    },
  ],
  execute: async (client, parameters) => {
    const query = parameters.query as string;
    const maxResults = (parameters.maxResults as number) ?? 10;
    return client.findChats(query, maxResults);
  },
  formatConcise: (result) => {
    const chats = result as ChatSearchResult[];
    if (chats.length === 0) return "No chats found.";
    const lines = [`## Chats (${chats.length} found)`, ""];
    for (const chat of chats) {
      lines.push(`### ${chat.name || "(untitled)"}`);
      lines.push(`- **Thread ID:** ${chat.threadId}`);
      lines.push(`- **Type:** ${chat.threadType}`);
      lines.push(`- **Members:** ${chat.totalMemberCount}`);
      if (chat.matchingMembers.length > 0) {
        lines.push(
          `- **Matched:** ${chat.matchingMembers.map((member) => member.displayName).join(", ")}`,
        );
      }
      lines.push("");
    }
    return lines.join("\n");
  },
};
