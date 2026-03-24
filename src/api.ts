/**
 * Low-level REST API layer for Teams Chat Service.
 *
 * All HTTP calls to {region}.ng.msg.teams.microsoft.com/v1 are here.
 * This module is stateless — every function takes a TeamsToken explicitly.
 *
 * Higher-level orchestration (pagination, search) lives in teams-client.ts.
 */

import type {
  TeamsToken,
  Conversation,
  Message,
  MessageFormat,
  MessagesPage,
  Member,
  Reaction,
  Mention,
  SentMessage,
  EditedMessage,
  DeletedMessage,
  UserProfile,
  PersonSearchResult,
  ChatSearchResult,
  TranscriptEntry,
  TranscriptResult,
} from "./types.js";
import MarkdownIt from "markdown-it";

const markdownRenderer = new MarkdownIt({ html: true, breaks: true });

export class ApiAuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ApiAuthError";
  }
}

export class ApiRateLimitError extends Error {
  public readonly retryAfterMilliseconds: number;

  constructor(message: string, retryAfterMilliseconds: number) {
    super(message);
    this.name = "ApiRateLimitError";
    this.retryAfterMilliseconds = retryAfterMilliseconds;
  }
}

const MAX_RETRY_ATTEMPTS = 5;
const INITIAL_BACKOFF_MILLISECONDS = 2_000;

function parseRetryAfter(response: Response): number {
  const retryAfterHeader = response.headers.get("Retry-After");
  if (!retryAfterHeader) {
    return INITIAL_BACKOFF_MILLISECONDS;
  }
  const seconds = Number(retryAfterHeader);
  if (!Number.isNaN(seconds) && seconds > 0) {
    return seconds * 1_000;
  }
  return INITIAL_BACKOFF_MILLISECONDS;
}

function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

/**
 * Wrapper around `fetch` that automatically retries on 429 (rate limit) responses.
 *
 * Uses the `Retry-After` header when present; otherwise applies exponential backoff
 * starting at 2 seconds. Retries up to 5 times before throwing `ApiRateLimitError`.
 */
async function fetchWithRetry(
  input: string | URL | Request,
  init?: RequestInit,
): Promise<Response> {
  for (let attempt = 0; attempt <= MAX_RETRY_ATTEMPTS; attempt++) {
    const response = await fetch(input, init);

    if (response.status !== 429) {
      return response;
    }

    if (attempt === MAX_RETRY_ATTEMPTS) {
      const errorText = await response.text();
      throw new ApiRateLimitError(
        `Rate limit exceeded after ${MAX_RETRY_ATTEMPTS + 1} attempts: ${response.status} ${errorText}`,
        parseRetryAfter(response),
      );
    }

    const retryAfterMilliseconds = parseRetryAfter(response);
    const backoffMilliseconds = retryAfterMilliseconds * Math.pow(2, attempt);
    await delay(backoffMilliseconds);
  }

  throw new Error("Unreachable: fetchWithRetry loop exited unexpectedly");
}

const chatServiceBase = (region: string) =>
  `https://${region}.ng.msg.teams.microsoft.com/v1`;

function authHeaders(token: TeamsToken): Record<string, string> {
  return {
    Authentication: `skypetoken=${token.skypeToken}`,
  };
}

/**
 * Fetch conversations from the Teams Chat Service.
 * Returns one page of conversations (no built-in pagination).
 */
export async function fetchConversations(
  token: TeamsToken,
  pageSize: number,
): Promise<Conversation[]> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations?view=mychats&pageSize=${pageSize}`;

  const response = await fetchWithRetry(url, { headers: authHeaders(token) });
  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to fetch conversations: ${response.status} ${response.statusText}`,
    );
  }

  const data = (await response.json()) as {
    conversations: Array<{
      id: string;
      version: number;
      threadProperties?: {
        topic?: string;
        threadType?: string;
        memberCount?: string;
      };
      properties?: { displayName?: string; lastimreceivedtime?: string };
    }>;
  };

  return (data.conversations ?? []).map((conversation) => ({
    id: conversation.id,
    topic:
      conversation.threadProperties?.topic ??
      conversation.properties?.displayName ??
      "",
    threadType: conversation.threadProperties?.threadType ?? "chat",
    version: conversation.version,
    lastMessageTime: conversation.properties?.lastimreceivedtime ?? null,
    memberCount: conversation.threadProperties?.memberCount
      ? Number(conversation.threadProperties.memberCount)
      : null,
  }));
}

/**
 * Fetch one page of messages from a conversation.
 */
export async function fetchMessagesPage(
  token: TeamsToken,
  conversationId: string,
  pageSize: number,
  backwardLink?: string,
): Promise<MessagesPage> {
  const url =
    backwardLink ??
    `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages?pageSize=${pageSize}`;

  const response = await fetchWithRetry(url, { headers: authHeaders(token) });
  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to fetch messages: ${response.status} ${response.statusText}`,
    );
  }

  const data = (await response.json()) as {
    messages: Array<Record<string, unknown>>;
    _metadata?: { backwardLink?: string; syncState?: string };
  };

  const messages = (data.messages ?? []).map(parseRawMessage);

  return {
    messages,
    backwardLink: data._metadata?.backwardLink ?? null,
    syncState: data._metadata?.syncState ?? null,
  };
}

/**
 * Fetch members of a conversation.
 */
export async function fetchMembers(
  token: TeamsToken,
  conversationId: string,
): Promise<Member[]> {
  const url = `${chatServiceBase(token.region)}/threads/${encodeURIComponent(conversationId)}/members`;

  const response = await fetchWithRetry(url, { headers: authHeaders(token) });
  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to fetch members: ${response.status} ${response.statusText}`,
    );
  }

  const data = (await response.json()) as {
    members: Array<{
      id: string;
      userDisplayName?: string;
      role?: string;
    }>;
  };

  return (data.members ?? []).map((member) => ({
    id: member.id,
    displayName: member.userDisplayName ?? "",
    role: member.role ?? "member",
    memberType: member.id.startsWith("28:")
      ? ("bot" as const)
      : ("person" as const),
  }));
}

const MIDDLE_TIER_BASE = "https://teams.cloud.microsoft/api/mt";

/**
 * Resolve display names for a batch of MRIs via the Teams middle-tier profile endpoint.
 *
 * Requires a Bearer token (api.spaces.skype.com audience). Throws `ApiAuthError`
 * if the token is unavailable so callers can trigger re-authentication.
 */
export async function fetchProfiles(
  token: TeamsToken,
  mris: string[],
): Promise<UserProfile[]> {
  if (mris.length === 0) {
    return [];
  }
  if (!token.bearerToken) {
    throw new ApiAuthError(
      "Bearer token is missing — re-authentication required for profile resolution",
    );
  }

  const url = `${MIDDLE_TIER_BASE}/${token.region}/beta/users/fetchShortProfile?isMailAddress=false&enableGuest=true&skypeTeamsInfo=true`;

  const response = await fetchWithRetry(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token.bearerToken}`,
    },
    body: JSON.stringify(mris),
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Profile resolution authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    return [];
  }

  const data = (await response.json()) as {
    value?: Array<{
      mri?: string;
      displayName?: string;
      email?: string;
      jobTitle?: string;
      userType?: string;
    }>;
  };

  return (data.value ?? []).map((profile) => ({
    mri: profile.mri ?? "",
    displayName: profile.displayName ?? "",
    email: profile.email ?? "",
    jobTitle: profile.jobTitle ?? "",
    userType: profile.userType ?? "",
  }));
}

/**
 * Convert content to the appropriate format for the Teams API.
 *
 * - "text": sent as-is with messagetype "Text"
 * - "markdown": converted from Markdown to HTML, sent as "RichText/Html"
 * - "html": sent as-is with messagetype "RichText/Html"
 */
function resolveMessageContent(
  content: string,
  format: MessageFormat,
): { resolvedContent: string; messagetype: string; contenttype: string } {
  switch (format) {
    case "text":
      return {
        resolvedContent: content,
        messagetype: "Text",
        contenttype: "text",
      };
    case "html":
      return {
        resolvedContent: content,
        messagetype: "RichText/Html",
        contenttype: "text",
      };
    case "markdown": {
      const htmlContent = markdownRenderer.render(content);
      return {
        resolvedContent: htmlContent,
        messagetype: "RichText/Html",
        contenttype: "text",
      };
    }
  }
}

/**
 * Send a message to a conversation.
 *
 * The `format` parameter controls how `content` is interpreted:
 * - `"text"` — plain text, sent as-is
 * - `"markdown"` (default) — converted from Markdown to HTML
 * - `"html"` — raw HTML, sent as-is
 */
export async function postMessage(
  token: TeamsToken,
  conversationId: string,
  content: string,
  senderDisplayName: string,
  format: MessageFormat = "markdown",
): Promise<SentMessage> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`;

  const clientMessageId = String(Date.now());
  const { resolvedContent, messagetype, contenttype } = resolveMessageContent(
    content,
    format,
  );

  const body = {
    content: resolvedContent,
    messagetype,
    contenttype,
    clientmessageid: clientMessageId,
    imdisplayname: senderDisplayName,
    properties: {
      importance: "",
      subject: null,
    },
  };

  const response = await fetchWithRetry(url, {
    method: "POST",
    headers: {
      ...authHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    const errorText = await response.text();
    throw new Error(
      `Failed to send message: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  const data = (await response.json()) as {
    OriginalArrivalTime: number;
  };

  return {
    messageId: String(data.OriginalArrivalTime),
    arrivalTime: data.OriginalArrivalTime,
  };
}

/**
 * Edit an existing message in a conversation.
 *
 * The `format` parameter controls how `content` is interpreted:
 * - `"text"` — plain text, sent as-is
 * - `"markdown"` (default) — converted from Markdown to HTML
 * - `"html"` — raw HTML, sent as-is
 */
export async function editMessage(
  token: TeamsToken,
  conversationId: string,
  messageId: string,
  content: string,
  senderDisplayName: string,
  format: MessageFormat = "markdown",
): Promise<EditedMessage> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${encodeURIComponent(messageId)}`;

  const { resolvedContent, messagetype, contenttype } = resolveMessageContent(
    content,
    format,
  );

  const body = {
    content: resolvedContent,
    messagetype,
    contenttype,
    skypeeditedid: messageId,
    imdisplayname: senderDisplayName,
  };

  const response = await fetchWithRetry(url, {
    method: "PUT",
    headers: {
      ...authHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    const errorText = await response.text();
    throw new Error(
      `Failed to edit message: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  const responseText = await response.text();
  let editTime: string;

  if (responseText) {
    const data = JSON.parse(responseText) as { edittime?: string };
    editTime = data.edittime ?? new Date().toISOString();
  } else {
    editTime = new Date().toISOString();
  }

  return {
    messageId,
    editTime,
  };
}

/**
 * Delete a message from a conversation.
 *
 * Sends a DELETE to the Chat Service; handles empty response body.
 */
export async function deleteMessage(
  token: TeamsToken,
  conversationId: string,
  messageId: string,
): Promise<DeletedMessage> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${encodeURIComponent(messageId)}`;

  const response = await fetchWithRetry(url, {
    method: "DELETE",
    headers: authHeaders(token),
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    const errorText = await response.text();
    throw new Error(
      `Failed to delete message: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  return { messageId };
}

/**
 * Fetch user properties for the authenticated user.
 * Returns raw properties — display name may or may not be present.
 */
export async function fetchUserProperties(
  token: TeamsToken,
): Promise<Record<string, unknown>> {
  const url = `${chatServiceBase(token.region)}/users/ME/properties`;
  const response = await fetchWithRetry(url, { headers: authHeaders(token) });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to fetch user properties: ${response.status} ${response.statusText}`,
    );
  }

  return (await response.json()) as Record<string, unknown>;
}

/**
 * Parse a raw API message object into the public Message type.
 */
export function parseRawMessage(raw: Record<string, unknown>): Message {
  const properties = (raw.properties ?? {}) as Record<string, unknown>;

  const reactions = parseReactions(properties.emotions);
  const mentions = parseMentions(properties.mentions);

  let quotedMessageId: string | null = null;
  const content = String(raw.content ?? "");
  const quoteMatch = content.match(
    /itemtype="http:\/\/schema\.skype\.com\/Reply"\s+itemid="(\d+)"/,
  );
  if (quoteMatch) {
    quotedMessageId = quoteMatch[1];
  }

  return {
    id: String(raw.id ?? ""),
    messageType: String(raw.messagetype ?? ""),
    senderMri: String(raw.from ?? ""),
    senderDisplayName: String(raw.imdisplayname ?? ""),
    content,
    originalArrivalTime: String(raw.originalarrivaltime ?? ""),
    composeTime: String(raw.composetime ?? ""),
    editTime: properties.edittime ? String(properties.edittime) : null,
    subject: properties.subject ? String(properties.subject) : null,
    isDeleted:
      raw.messagetype === "MessageDelete" ||
      String(properties.deletetime ?? "") !== "",
    reactions,
    mentions,
    quotedMessageId,
  };
}

function parseReactions(rawEmotions: unknown): Reaction[] {
  if (typeof rawEmotions === "string") {
    try {
      return JSON.parse(rawEmotions) as Reaction[];
    } catch {
      return [];
    }
  }

  if (Array.isArray(rawEmotions)) {
    return rawEmotions as Reaction[];
  }

  return [];
}

function parseMentions(rawMentions: unknown): Mention[] {
  let parsed: Array<{ id?: string; displayName?: string }>;

  if (typeof rawMentions === "string") {
    try {
      parsed = JSON.parse(rawMentions) as Array<{
        id?: string;
        displayName?: string;
      }>;
    } catch {
      return [];
    }
  } else if (Array.isArray(rawMentions)) {
    parsed = rawMentions as Array<{ id?: string; displayName?: string }>;
  } else {
    return [];
  }

  return parsed.map((mention) => ({
    id: mention.id ?? "",
    displayName: mention.displayName ?? "",
  }));
}

// ── Transcript helpers ────────────────────────────────────────────────

/**
 * Extract the AMS transcript URL from a `RichText/Media_CallRecording` message.
 *
 * The message content is XML containing `<item type="amsTranscript" uri="...">`.
 * Returns null if no transcript URL is found.
 */
export function extractTranscriptUrl(messageContent: string): string | null {
  const match = messageContent.match(
    /<item\s+[^>]*type="amsTranscript"[^>]*\buri="([^"]+)"/,
  );
  return match?.[1] ?? null;
}

/**
 * Extract the meeting title from a `RichText/Media_CallRecording` message.
 *
 * The title is in the `<OriginalName>` element inside the XML content.
 */
export function extractMeetingTitle(messageContent: string): string {
  const match = messageContent.match(/<OriginalName\b[^>]*v="([^"]*)"[^>]*\/>/);
  return match?.[1] ?? "Unknown Meeting";
}

/**
 * Check whether a `RichText/Media_CallRecording` message represents
 * a successful recording (as opposed to started/failed).
 */
export function isSuccessfulRecording(messageContent: string): boolean {
  return /<RecordingStatus\b[^>]*status="Success"/.test(messageContent);
}

/**
 * Decode HTML entities to plain text (used for VTT content).
 */
function decodeVttEntities(text: string): string {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, code: string) =>
      String.fromCharCode(Number(code)),
    );
}

/**
 * Parse VTT content into structured transcript entries.
 *
 * Handles the Teams VTT format with `<v Speaker Name>text</v>` tags
 * and HTML entities.
 */
export function parseVtt(vttContent: string): TranscriptEntry[] {
  const entries: TranscriptEntry[] = [];
  const lines = vttContent.split("\n");

  let currentStartTime = "";
  let currentEndTime = "";

  for (const line of lines) {
    // Match timestamp lines: "00:00:00.000 --> 00:00:05.000"
    const timestampMatch = line.match(
      /^(\d{2}:\d{2}:\d{2}\.\d{3})\s+-->\s+(\d{2}:\d{2}:\d{2}\.\d{3})/,
    );
    if (timestampMatch) {
      currentStartTime = timestampMatch[1];
      currentEndTime = timestampMatch[2];
      continue;
    }

    // Match speaker lines: "<v Speaker Name>text</v>"
    const speakerMatch = line.match(/^<v\s+([^>]+)>(.+)<\/v>\s*$/);
    if (speakerMatch && currentStartTime) {
      entries.push({
        speaker: decodeVttEntities(speakerMatch[1]),
        startTime: currentStartTime,
        endTime: currentEndTime,
        text: decodeVttEntities(speakerMatch[2]),
      });
      currentStartTime = "";
      currentEndTime = "";
    }
  }

  return entries;
}

/**
 * Fetch a transcript from the AMS (Async Media Service).
 *
 * Uses `Authorization: skype_token <token>` header (note: different format
 * from the Chat Service `Authentication: skypetoken=<token>` header).
 */
export async function fetchTranscriptVtt(
  token: TeamsToken,
  amsTranscriptUrl: string,
): Promise<string> {
  const response = await fetchWithRetry(amsTranscriptUrl, {
    headers: {
      Authorization: `skype_token ${token.skypeToken}`,
    },
  });

  if (!response.ok) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Transcript fetch authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to fetch transcript: ${response.status} ${response.statusText}`,
    );
  }

  return response.text();
}

/**
 * Fetch and parse the transcript for a conversation.
 *
 * 1. Fetches messages from the conversation
 * 2. Finds the latest `RichText/Media_CallRecording` message with status="Success"
 * 3. Extracts the AMS transcript URL from the XML content
 * 4. Fetches the VTT content from AMS
 * 5. Parses the VTT into structured entries
 */
export async function fetchTranscript(
  token: TeamsToken,
  conversationId: string,
): Promise<TranscriptResult> {
  // Fetch messages to find the recording message
  const pageSize = 200;
  const maxPages = 20;
  let backwardLink: string | undefined;
  let amsTranscriptUrl: string | null = null;
  let meetingTitle = "Unknown Meeting";

  for (let pageIndex = 0; pageIndex < maxPages; pageIndex++) {
    const page = await fetchMessagesPage(
      token,
      conversationId,
      pageSize,
      backwardLink,
    );

    for (const message of page.messages) {
      if (
        message.messageType === "RichText/Media_CallRecording" &&
        isSuccessfulRecording(message.content)
      ) {
        const transcriptUrl = extractTranscriptUrl(message.content);
        if (transcriptUrl) {
          amsTranscriptUrl = transcriptUrl;
          meetingTitle = extractMeetingTitle(message.content);
          break;
        }
      }
    }

    if (amsTranscriptUrl || !page.backwardLink) break;
    backwardLink = page.backwardLink;
  }

  if (!amsTranscriptUrl) {
    throw new Error(
      "No meeting transcript found in this conversation. " +
        "Make sure the conversation contains a recorded meeting with a transcript.",
    );
  }

  const rawVtt = await fetchTranscriptVtt(token, amsTranscriptUrl);
  const entries = parseVtt(rawVtt);

  return { meetingTitle, rawVtt, entries };
}

// ── Substrate search API ──────────────────────────────────────────────

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
    if (response.status === 401) {
      throw new ApiAuthError(
        `Substrate search authentication failed: ${response.status} ${response.statusText}`,
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
    if (response.status === 401) {
      throw new ApiAuthError(
        `Substrate search authentication failed: ${response.status} ${response.statusText}`,
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
