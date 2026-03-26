/**
 * REST API layer for the Teams Chat Service.
 *
 * All HTTP calls to {region}.ng.msg.teams.microsoft.com/v1 are here.
 * This module is stateless — every function takes a TeamsToken explicitly.
 */

import type {
  TeamsToken,
  Conversation,
  Message,
  MessageFormat,
  MessagesPage,
  Member,
  Reaction,
  Follower,
  Mention,
  SentMessage,
  EditedMessage,
  DeletedMessage,
  ReactionResult,
  ScheduledMessage,
} from "../types.js";
import { fetchWithRetry, ApiAuthError } from "./common.js";
import { parseInlineImages, parseFileAttachments } from "./attachments.js";
import MarkdownIt from "markdown-it";

const markdownRenderer = new MarkdownIt({ html: true, breaks: true });

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
  amsReferences: string[] = [],
  filesJson?: string,
  subject?: string,
): Promise<SentMessage> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`;

  const clientMessageId = String(Date.now());
  const { resolvedContent, messagetype, contenttype } = resolveMessageContent(
    content,
    format,
  );

  const body: Record<string, unknown> = {
    content: resolvedContent,
    messagetype,
    contenttype,
    clientmessageid: clientMessageId,
    imdisplayname: senderDisplayName,
    properties: {
      importance: "",
      subject: subject ?? null,
      ...(filesJson ? { files: filesJson } : {}),
    },
  };

  if (amsReferences.length > 0) {
    body.amsreferences = amsReferences;
  }

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
 * Schedule a message to be sent at a future time.
 *
 * Creates a "ScheduledDraft" via the Chat Service drafts endpoint.
 * The server delivers the message automatically at the specified time.
 *
 * The `format` parameter controls how `content` is interpreted:
 * - `"text"` — plain text, sent as-is
 * - `"markdown"` (default) — converted from Markdown to HTML
 * - `"html"` — raw HTML, sent as-is
 */
export async function postScheduledMessage(
  token: TeamsToken,
  conversationId: string,
  content: string,
  senderDisplayName: string,
  scheduleAt: Date,
  format: MessageFormat = "markdown",
  amsReferences: string[] = [],
  filesJson?: string,
  subject?: string,
): Promise<ScheduledMessage> {
  const url = `${chatServiceBase(token.region)}/users/ME/drafts`;

  const clientMessageId = String(Date.now());
  const { resolvedContent, messagetype } = resolveMessageContent(
    content,
    format,
  );
  const now = new Date().toISOString();

  const body = {
    draftDetails: {
      sendAt: String(scheduleAt.getTime()),
    },
    draftType: "ScheduledDraft",
    innerThreadId: conversationId,
    message: {
      id: "-1",
      type: "Message",
      conversationid: conversationId,
      conversationLink: `blah/${conversationId}`,
      composetime: now,
      originalarrivaltime: now,
      content: resolvedContent,
      messagetype,
      contenttype: "Text",
      imdisplayname: senderDisplayName,
      clientmessageid: clientMessageId,
      callId: "",
      state: 0,
      version: "0",
      amsreferences: amsReferences,
      properties: {
        importance: "",
        subject: subject ?? "",
        cards: "[]",
        links: "[]",
        mentions: "[]",
        onbehalfof: null,
        ...(filesJson ? { files: filesJson } : { files: "[]" }),
        policyViolation: null,
        formatVariant: "TEAMS",
        draftId: clientMessageId,
      },
      crossPostChannels: [],
      draftDetails: {
        sendAt: scheduleAt.toISOString(),
      },
      threadtype: "streamofdrafts",
      innerThreadId: conversationId,
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
      `Failed to schedule message: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  const data = (await response.json()) as {
    OriginalArrivalTime: number;
  };

  return {
    messageId: String(data.OriginalArrivalTime),
    arrivalTime: data.OriginalArrivalTime,
    scheduledTime: scheduleAt.toISOString(),
  };
}

/**
 * Add a reaction (emotion) to a message.
 *
 * Uses the Chat Service emotions property endpoint.
 * The `reactionKey` is the emotion name (e.g. "like", "heart", "laugh", "surprised").
 */
export async function addReaction(
  token: TeamsToken,
  conversationId: string,
  messageId: string,
  reactionKey: string,
): Promise<ReactionResult> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${encodeURIComponent(messageId)}/properties?name=emotions`;

  const body = {
    emotions: { key: reactionKey, value: messageId },
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
      `Failed to add reaction: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  return { messageId, reactionKey };
}

/**
 * Remove a reaction (emotion) from a message.
 *
 * Uses the Chat Service emotions property endpoint with DELETE method.
 * Only removes the current user's reaction of the specified type.
 */
export async function removeReaction(
  token: TeamsToken,
  conversationId: string,
  messageId: string,
  reactionKey: string,
): Promise<ReactionResult> {
  const url = `${chatServiceBase(token.region)}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${encodeURIComponent(messageId)}/properties?name=emotions`;

  const body = {
    emotions: { key: reactionKey, value: messageId },
  };

  const response = await fetchWithRetry(url, {
    method: "DELETE",
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
      `Failed to remove reaction: ${response.status} ${response.statusText} — ${errorText}`,
    );
  }

  return { messageId, reactionKey };
}

/**
 * Create a new 1:1 conversation with a given member, or return the existing one.
 *
 * Uses `PUT /users/ME/conversations` with a single member MRI.
 * A 201 response means the conversation was created; a 200 response (or 409 on
 * some server versions) means it already existed — both return the conversation ID.
 */
export async function createOneOnOneConversation(
  token: TeamsToken,
  memberMri: string,
): Promise<{ id: string }> {
  if (!memberMri || !/^8:orgid:[0-9a-f-]+$/i.test(memberMri)) {
    throw new Error(
      `Invalid member MRI for 1:1 conversation creation: "${memberMri}". Expected format: 8:orgid:{uuid}`,
    );
  }

  const url = `${chatServiceBase(token.region)}/users/ME/conversations`;
  const response = await fetchWithRetry(url, {
    method: "PUT",
    headers: {
      ...authHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      members: [{ mri: memberMri, role: "Admin" }],
    }),
  });

  // 200 = already exists, 201 = newly created, 409 = conflict / already exists
  if (!response.ok && response.status !== 409) {
    if (response.status === 401) {
      throw new ApiAuthError(
        `Authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    throw new Error(
      `Failed to create conversation: ${response.status} ${response.statusText}`,
    );
  }

  const data = (await response.json()) as { id?: string };
  if (!data.id) {
    throw new Error("No conversation ID returned when creating 1:1 conversation");
  }
  return { id: data.id };
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

  const allEmotions = parseEmotions(properties.emotions);
  const reactions = allEmotions.filter(
    (emotion) => emotion.key !== "follow" && emotion.users.length > 0,
  );
  const followEntry = allEmotions.find((emotion) => emotion.key === "follow");
  const followers = (followEntry?.users ?? [])
    .filter((user) => String(user.value) === "0")
    .map((user) => ({ mri: user.mri, time: user.time }));
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
    followers,
    mentions,
    quotedMessageId,
    images: parseInlineImages(content),
    files: parseFileAttachments(properties.files),
  };
}

function parseEmotions(rawEmotions: unknown): Array<{
  key: string;
  users: Array<{ mri: string; time: number; value?: unknown }>;
}> {
  if (typeof rawEmotions === "string") {
    try {
      return JSON.parse(rawEmotions) as Array<{
        key: string;
        users: Array<{ mri: string; time: number; value?: unknown }>;
      }>;
    } catch {
      return [];
    }
  }

  if (Array.isArray(rawEmotions)) {
    return rawEmotions as Array<{
      key: string;
      users: Array<{ mri: string; time: number; value?: unknown }>;
    }>;
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
