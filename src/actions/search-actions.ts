/**
 * Search-related action definitions.
 *
 * Actions: get-profiles, find-people, find-chats.
 */

import type {
  PersonSearchResult,
  ChatSearchResult,
  UserProfile,
  SearchMessagesOptions,
  MessageSearchPage,
} from "../types.js";
import type { ActionDefinition } from "./formatters.js";

export const searchMessagesAction: ActionDefinition = {
  name: "search-messages",
  title: "Search Messages",
  description:
    "Search message content across accessible Teams chats and channel posts/replies. " +
    "Read-only; does not open chats, send messages, or mark messages read. " +
    "Requires a Substrate search token. Returns one page with message IDs, previews, Teams links, total estimate, and nextOffset. " +
    "Search visibility and retention policies apply; zero results do not prove no messages exist.",
  parameters: [
    {
      name: "query",
      type: "string",
      description:
        "Search keywords or advanced search query; use * for filter-only searches",
      required: true,
    },
    {
      name: "from",
      type: "string",
      description: "Sender email address to filter authored messages",
      required: false,
    },
    {
      name: "messageType",
      type: "string",
      description:
        "Message scope: all (default), chat, or channel (includes replies)",
      required: false,
      default: "all",
    },
    {
      name: "conversationId",
      type: "string",
      description: "Restrict results to this chat or channel conversation ID",
      required: false,
    },
    {
      name: "sentAfter",
      type: "string",
      description: "Inclusive start date in YYYY-MM-DD format",
      required: false,
    },
    {
      name: "sentBefore",
      type: "string",
      description: "Inclusive end date in YYYY-MM-DD format",
      required: false,
    },
    {
      name: "offset",
      type: "number",
      description:
        "Result offset; pass nextOffset from the previous page (default: 0)",
      required: false,
      default: 0,
    },
    {
      name: "pageSize",
      type: "number",
      description: "Number of results per page, from 1 to 100 (default: 25)",
      required: false,
      default: 25,
    },
  ],
  execute: (client, parameters) =>
    client.searchMessages({
      query: parameters.query as string,
      from: parameters.from as string | undefined,
      messageType:
        (parameters.messageType as SearchMessagesOptions["messageType"]) ??
        "all",
      conversationId: parameters.conversationId as string | undefined,
      sentAfter: parameters.sentAfter as string | undefined,
      sentBefore: parameters.sentBefore as string | undefined,
      offset: (parameters.offset as number | undefined) ?? 0,
      pageSize: (parameters.pageSize as number | undefined) ?? 25,
    }),
  formatConcise: (result) => {
    const page = result as MessageSearchPage;
    const lines = [
      `## Message Search (${page.messages.length} returned; approximately ${page.total} matches)`,
      "",
    ];
    if (!page.messages.length) lines.push("No matching messages found.", "");
    for (const message of page.messages) {
      lines.push(
        `### ${message.senderDisplayName || message.senderEmail || "(unknown)"}${message.sentAt ? ` | ${message.sentAt}` : ""}`,
      );
      if (message.senderEmail)
        lines.push(`- **Email:** ${message.senderEmail}`);
      if (message.conversationName)
        lines.push(`- **Conversation:** ${message.conversationName}`);
      if (message.conversationId)
        lines.push(`- **Conversation ID:** ${message.conversationId}`);
      if (message.messageId)
        lines.push(`- **Message ID:** ${message.messageId}`);
      if (message.parentMessageId)
        lines.push(`- **Parent message ID:** ${message.parentMessageId}`);
      if (message.subject) lines.push(`- **Subject:** ${message.subject}`);
      if (message.preview !== undefined) lines.push("", message.preview);
      if (message.webUrl) lines.push("", `[Open in Teams](${message.webUrl})`);
      lines.push("");
    }
    if (page.hasMore)
      lines.push(`More results available. **nextOffset:** ${page.nextOffset}`);
    return lines.join("\n");
  },
};

export const getProfilesAction: ActionDefinition = {
  name: "get-profiles",
  title: "Get User Profiles",
  description:
    "Get Teams profiles for one or more people by MRI. " +
    "Returns display name, email, job title, office location, and user type.",
  parameters: [
    {
      name: "userIdentifiers",
      type: "string[]",
      description:
        "One or more Teams MRIs (for example, 8:orgid:00000000-0000-0000-0000-000000000000)",
      required: true,
    },
  ],
  execute: async (client, parameters) => {
    const userIdentifiers = (
      (parameters.userIdentifiers as string[] | undefined) ?? []
    )
      .map((userIdentifier) => userIdentifier.trim())
      .filter((userIdentifier) => userIdentifier.length > 0);
    if (userIdentifiers.length === 0) {
      throw new Error("At least one user identifier is required");
    }
    return client.getProfiles(userIdentifiers);
  },
  formatConcise: (result) => {
    const profiles = result as UserProfile[];
    if (profiles.length === 0) return "No profiles found.";
    const lines = [`## Profiles (${profiles.length} found)`, ""];
    for (const profile of profiles) {
      lines.push(`### ${profile.displayName || "(unknown)"}`);
      if (profile.email) lines.push(`- **Email:** ${profile.email}`);
      if (profile.jobTitle) lines.push(`- **Title:** ${profile.jobTitle}`);
      if (profile.userLocation)
        lines.push(`- **Office location:** ${profile.userLocation}`);
      if (profile.userType) lines.push(`- **User type:** ${profile.userType}`);
      lines.push(`- **MRI:** ${profile.mri}`, "");
    }
    return lines.join("\n");
  },
};

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
      if (person.userLocation)
        lines.push(`- **Office location:** ${person.userLocation}`);
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
