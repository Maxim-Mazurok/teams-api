/**
 * Search-related action definitions.
 *
 * Actions: get-profiles, find-people, find-chats.
 */

import type {
  PersonSearchResult,
  ChatSearchResult,
  UserProfile,
} from "../types.js";
import type { ActionDefinition } from "./formatters.js";

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
