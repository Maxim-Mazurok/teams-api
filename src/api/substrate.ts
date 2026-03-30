/**
 * Substrate search API for people and chat discovery.
 *
 * HTTP calls to substrate.office.com for searching people and chats.
 * Requires a Substrate Bearer token (substrate.office.com audience).
 */

import type {
  TeamsToken,
  PersonSearchResult,
  ChatSearchResult,
} from "../types.js";
import { fetchWithRetry, ApiAuthError } from "./common.js";

const SUBSTRATE_SEARCH_BASE = "https://substrate.office.com";

/**
 * Search for people by name using the Substrate suggestions API.
 *
 * Requires a Substrate Bearer token (substrate.office.com audience).
 * Returns matching people with their MRIs, emails, job titles, etc.
 */
export async function searchPeople(
  token: TeamsToken,
  query: string,
  maxResults = 10,
): Promise<PersonSearchResult[]> {
  if (!token.substrateToken) {
    throw new ApiAuthError(
      "Substrate token is missing — re-authentication required for people search",
    );
  }

  const url = `${SUBSTRATE_SEARCH_BASE}/search/api/v1/suggestions?scenario=peoplepicker.newChat`;

  const body = {
    EntityRequests: [
      {
        Query: {
          QueryString: query,
          DisplayQueryString: query,
        },
        EntityType: "People",
        Size: maxResults,
        Fields: [
          "Id",
          "MRI",
          "DisplayName",
          "EmailAddresses",
          "PeopleType",
          "PeopleSubtype",
          "UserPrincipalName",
          "GivenName",
          "Surname",
          "JobTitle",
          "Department",
          "ExternalDirectoryObjectId",
        ],
        Filter: {
          And: [
            {
              Or: [
                { Term: { PeopleType: "Person" } },
                { Term: { PeopleType: "Other" } },
              ],
            },
            {
              Or: [
                { Term: { PeopleSubtype: "OrganizationUser" } },
                { Term: { PeopleSubtype: "MTOUser" } },
                { Term: { PeopleSubtype: "Guest" } },
              ],
            },
            { Or: [{ Term: { Flags: "NonHidden" } }] },
          ],
        },
        Provenances: ["Mailbox", "Directory"],
        From: 0,
      },
    ],
    Scenario: { Name: "peoplepicker.newChat" },
    AppName: "Microsoft Teams",
  };

  const response = await fetchWithRetry(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token.substrateToken}`,
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      throw new ApiAuthError(
        `Substrate search authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    if (response.status >= 500) {
      throw new Error(
        `Substrate search server error: ${response.status} ${response.statusText}`,
      );
    }
    return [];
  }

  const data = (await response.json()) as {
    Groups?: Array<{
      Type?: string;
      Suggestions?: Array<{
        DisplayName?: string;
        MRI?: string;
        EmailAddresses?: string[];
        UserPrincipalName?: string;
        JobTitle?: string;
        Department?: string;
        ExternalDirectoryObjectId?: string;
      }>;
    }>;
  };

  const peopleGroup = data.Groups?.find((group) => group.Type === "People");
  if (!peopleGroup?.Suggestions) return [];

  return peopleGroup.Suggestions.map((suggestion) => ({
    displayName: suggestion.DisplayName ?? "",
    mri: suggestion.MRI ?? "",
    email: suggestion.EmailAddresses?.[0] ?? suggestion.UserPrincipalName ?? "",
    jobTitle: suggestion.JobTitle ?? "",
    department: suggestion.Department ?? "",
    objectId: suggestion.ExternalDirectoryObjectId ?? "",
  }));
}

/**
 * Search for chats by name or member using the Substrate suggestions API.
 *
 * Requires a Substrate Bearer token (substrate.office.com audience).
 * Returns matching chats with their thread IDs, members, etc.
 */
export async function searchChats(
  token: TeamsToken,
  query: string,
  maxResults = 10,
): Promise<ChatSearchResult[]> {
  if (!token.substrateToken) {
    throw new ApiAuthError(
      "Substrate token is missing — re-authentication required for chat search",
    );
  }

  const url = `${SUBSTRATE_SEARCH_BASE}/search/api/v1/suggestions?scenario=peoplepicker.newChat`;

  const body = {
    EntityRequests: [
      {
        Query: {
          QueryString: query,
        },
        EntityType: "Chat",
        Size: maxResults,
      },
    ],
    Scenario: { Name: "peoplepicker.newChat" },
    AppName: "Microsoft Teams",
  };

  const response = await fetchWithRetry(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token.substrateToken}`,
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      throw new ApiAuthError(
        `Substrate search authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    if (response.status >= 500) {
      throw new Error(
        `Substrate search server error: ${response.status} ${response.statusText}`,
      );
    }
    return [];
  }

  const data = (await response.json()) as {
    Groups?: Array<{
      Type?: string;
      Suggestions?: Array<{
        Name?: string;
        ThreadId?: string;
        ThreadType?: string;
        MatchingMembers?: Array<{
          DisplayName?: string;
          MRI?: string;
        }>;
        ChatMembers?: Array<{
          DisplayName?: string;
          MRI?: string;
        }>;
        TotalChatMembersCount?: number;
      }>;
    }>;
  };

  const chatGroup = data.Groups?.find((group) => group.Type === "Chat");
  if (!chatGroup?.Suggestions) return [];

  return chatGroup.Suggestions.map((suggestion) => ({
    name: suggestion.Name ?? "",
    threadId: suggestion.ThreadId ?? "",
    threadType: suggestion.ThreadType ?? "",
    matchingMembers: (suggestion.MatchingMembers ?? []).map((member) => ({
      displayName: member.DisplayName ?? "",
      mri: member.MRI ?? "",
    })),
    chatMembers: (suggestion.ChatMembers ?? []).map((member) => ({
      displayName: member.DisplayName ?? "",
      mri: member.MRI ?? "",
    })),
    totalMemberCount: suggestion.TotalChatMembersCount ?? 0,
  }));
}
