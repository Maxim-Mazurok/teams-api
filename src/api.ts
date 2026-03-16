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
  MessagesPage,
  Member,
  Reaction,
  Mention,
  SentMessage,
} from "./types.js";

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

  const response = await fetch(url, { headers: authHeaders(token) });
  if (!response.ok) {
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

  const response = await fetch(url, { headers: authHeaders(token) });
  if (!response.ok) {
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

  const response = await fetch(url, { headers: authHeaders(token) });
  if (!response.ok) {
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
  }));
}

/**
 * Send a plain-text message to a conversation.
 */
export async function postMessage(
  token: TeamsToken,
  conversationId: string,
  content: string,
  senderDisplayName: string,
): Promise<SentMessage> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`;

  const clientMessageId = String(Date.now());

  const body = {
    content,
    messagetype: "Text",
    contenttype: "text",
    clientmessageid: clientMessageId,
    imdisplayname: senderDisplayName,
    properties: {
      importance: "",
      subject: null,
    },
  };

  const response = await fetch(url, {
    method: "POST",
    headers: {
      ...authHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
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
  const response = await fetch(url, { headers: authHeaders(token) });

  if (!response.ok) {
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
