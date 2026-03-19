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
  UserProfile,
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
 * Requires a Bearer token (api.spaces.skype.com audience). Returns an empty array
 * if the token is unavailable or the request fails.
 */
export async function fetchProfiles(
  token: TeamsToken,
  mris: string[],
): Promise<UserProfile[]> {
  if (!token.bearerToken || mris.length === 0) {
    return [];
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
    id?: string;
  };

  return {
    messageId: data.id ?? clientMessageId,
    arrivalTime: data.OriginalArrivalTime,
  };
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
