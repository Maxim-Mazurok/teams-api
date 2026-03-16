/**
 * TeamsClient — the public API for interacting with Microsoft Teams.
 *
 * This is the primary entry point for the teams-api package.
 * It wraps authentication, conversation listing, message reading/sending,
 * and member management behind a clean, strongly-typed interface.
 *
 * @example
 *   // Auto-login (macOS with FIDO2 passkey)
 *   const client = await TeamsClient.fromAutoLogin({
 *     email: "user@company.com",
 *   });
 *
 *   // List conversations
 *   const conversations = await client.listConversations();
 *
 *   // Read messages
 *   const messages = await client.getMessages(conversations[0].id);
 *
 *   // Send a message
 *   await client.sendMessage(conversations[0].id, "Hello from the API!");
 *
 * @example
 *   // From an existing token
 *   const client = TeamsClient.fromToken("skype-token-here", "apac");
 */

import type {
  TeamsToken,
  AutoLoginOptions,
  ManualTokenOptions,
  Conversation,
  Message,
  Member,
  SentMessage,
  GetMessagesOptions,
  ListConversationsOptions,
  OneOnOneSearchResult,
  MessagesPage,
  SYSTEM_STREAM_TYPES,
} from "./types.js";
import {
  fetchConversations,
  fetchMessagesPage,
  fetchMembers,
  postMessage,
  fetchUserProperties,
} from "./api.js";
import {
  acquireTokenViaAutoLogin,
  acquireTokenViaDebugSession,
} from "./auth.js";

export {
  acquireTokenViaAutoLogin,
  acquireTokenViaDebugSession,
} from "./auth.js";
export type {
  TeamsToken,
  AutoLoginOptions,
  ManualTokenOptions,
  Conversation,
  Message,
  Member,
  SentMessage,
  GetMessagesOptions,
  ListConversationsOptions,
  OneOnOneSearchResult,
  MessagesPage,
  Mention,
  Reaction,
} from "./types.js";
export { SYSTEM_STREAM_TYPES } from "./types.js";
export { parseRawMessage } from "./api.js";

const SYSTEM_STREAMS: readonly string[] = [
  "streamofannotations",
  "streamofthreads",
  "streamofnotifications",
  "streamofmentions",
  "streamofnotes",
];

export class TeamsClient {
  private readonly token: TeamsToken;
  private cachedDisplayName: string | null = null;

  private constructor(token: TeamsToken) {
    this.token = token;
  }

  /**
   * Create a client by automatically logging in via FIDO2 passkey.
   *
   * Launches system Chrome with a fresh profile, completes the Microsoft
   * Entra ID FIDO2 login flow using a platform authenticator, and
   * captures the skype token. Zero user interaction required.
   */
  static async fromAutoLogin(options: AutoLoginOptions): Promise<TeamsClient> {
    const token = await acquireTokenViaAutoLogin(options);
    return new TeamsClient(token);
  }

  /**
   * Create a client by connecting to a running Chrome debug session.
   *
   * Requires Chrome started with --remote-debugging-port and Teams
   * already open and authenticated.
   */
  static async fromDebugSession(
    options?: ManualTokenOptions,
  ): Promise<TeamsClient> {
    const token = await acquireTokenViaDebugSession(options);
    return new TeamsClient(token);
  }

  /**
   * Create a client from an existing skype token.
   *
   * Use this when you already have a valid token (e.g. from a previous
   * session or external source). Token lifetime is ~24 hours.
   */
  static fromToken(skypeToken: string, region = "apac"): TeamsClient {
    return new TeamsClient({ skypeToken, region });
  }

  /** Get the underlying token (for persistence or debugging). */
  getToken(): TeamsToken {
    return { ...this.token };
  }

  /**
   * List conversations (chats, group chats, meetings, channels).
   *
   * By default, excludes system streams (annotations, notifications,
   * mentions). Set `excludeSystemStreams: false` to include everything.
   */
  async listConversations(
    options?: ListConversationsOptions,
  ): Promise<Conversation[]> {
    const pageSize = options?.pageSize ?? 50;
    const excludeSystem = options?.excludeSystemStreams ?? true;

    const conversations = await fetchConversations(this.token, pageSize);

    if (!excludeSystem) {
      return conversations;
    }

    return conversations.filter(
      (conversation) => !SYSTEM_STREAMS.includes(conversation.threadType),
    );
  }

  /**
   * Find a conversation by topic name (case-insensitive partial match).
   *
   * Searches conversations with topics. For 1:1 chats (which have no topic),
   * use `findOneOnOneConversation()` instead.
   */
  async findConversation(query: string): Promise<Conversation | null> {
    const conversations = await this.listConversations({ pageSize: 100 });
    const queryLower = query.toLowerCase();

    return (
      conversations.find(
        (conversation) =>
          conversation.topic &&
          conversation.topic.toLowerCase().includes(queryLower),
      ) ?? null
    );
  }

  /**
   * Find a 1:1 conversation with a specific person.
   *
   * Searches by scanning recent messages in untitled chats for sender
   * display names that match the query. Also checks the self-chat
   * (48:notes) if the query matches the current user's name.
   *
   * This is necessary because the Teams members API returns empty display
   * names for 1:1 chat participants.
   */
  async findOneOnOneConversation(
    personName: string,
  ): Promise<OneOnOneSearchResult | null> {
    const conversations = await fetchConversations(this.token, 100);
    const targetLower = personName.toLowerCase();

    // Check self-chat first
    const selfChat = conversations.find((conversation) =>
      conversation.id.startsWith("48:notes"),
    );
    if (selfChat) {
      const currentUserName = await this.getCurrentUserDisplayName();
      if (currentUserName.toLowerCase().includes(targetLower)) {
        return {
          conversationId: selfChat.id,
          memberDisplayName: `${currentUserName} (self)`,
        };
      }
    }

    // Search untitled 1:1 chats by scanning recent message senders
    const untitledChats = conversations.filter(
      (conversation) =>
        conversation.threadType === "chat" && !conversation.topic,
    );

    for (const chat of untitledChats) {
      try {
        const page = await fetchMessagesPage(this.token, chat.id, 10);
        const textMessages = page.messages.filter(
          (message) =>
            message.messageType === "RichText/Html" ||
            message.messageType === "Text",
        );

        const senderNames = [
          ...new Set(
            textMessages
              .map((message) => message.senderDisplayName)
              .filter((name) => name.length > 0),
          ),
        ];

        for (const senderName of senderNames) {
          if (senderName.toLowerCase().includes(targetLower)) {
            return {
              conversationId: chat.id,
              memberDisplayName: senderName,
            };
          }
        }
      } catch {
        continue;
      }
    }

    return null;
  }

  /**
   * Get all messages from a conversation.
   *
   * Follows pagination links to fetch the complete message history.
   * Use `maxPages` and `pageSize` to control how much is fetched.
   */
  async getMessages(
    conversationId: string,
    options?: GetMessagesOptions,
  ): Promise<Message[]> {
    const maxPages = options?.maxPages ?? 100;
    const pageSize = options?.pageSize ?? 200;
    const allMessages: Message[] = [];
    let backwardLink: string | undefined;

    for (let pageIndex = 0; pageIndex < maxPages; pageIndex++) {
      const result = await fetchMessagesPage(
        this.token,
        conversationId,
        pageSize,
        backwardLink,
      );
      allMessages.push(...result.messages);

      options?.onProgress?.(allMessages.length);

      if (!result.backwardLink) break;
      backwardLink = result.backwardLink;
    }

    return allMessages;
  }

  /**
   * Get one page of messages from a conversation.
   *
   * Returns the page along with a backwardLink for manual pagination.
   */
  async getMessagesPage(
    conversationId: string,
    pageSize = 50,
    backwardLink?: string,
  ): Promise<MessagesPage> {
    return fetchMessagesPage(
      this.token,
      conversationId,
      pageSize,
      backwardLink,
    );
  }

  /**
   * Send a plain-text message to a conversation.
   *
   * The sender display name is resolved automatically from the
   * current user's message history (via self-chat).
   */
  async sendMessage(
    conversationId: string,
    content: string,
  ): Promise<SentMessage> {
    const displayName = await this.getCurrentUserDisplayName();
    return postMessage(this.token, conversationId, content, displayName);
  }

  /**
   * Get members of a conversation.
   *
   * Note: For 1:1 chats, member display names may be empty.
   * Use `findOneOnOneConversation()` to resolve names from message history.
   */
  async getMembers(conversationId: string): Promise<Member[]> {
    return fetchMembers(this.token, conversationId);
  }

  /**
   * Get the display name of the currently authenticated user.
   *
   * Resolved by reading messages from the self-chat (48:notes) or
   * falling back to the user properties endpoint.
   */
  async getCurrentUserDisplayName(): Promise<string> {
    if (this.cachedDisplayName) {
      return this.cachedDisplayName;
    }

    const conversations = await fetchConversations(this.token, 10);

    // First try: self-chat messages (most reliable)
    for (const conversation of conversations) {
      try {
        if (!conversation.id.startsWith("48:notes")) continue;

        const page = await fetchMessagesPage(this.token, conversation.id, 10);
        const textMessage = page.messages.find(
          (message) =>
            (message.messageType === "RichText/Html" ||
              message.messageType === "Text") &&
            message.senderDisplayName.length > 0,
        );

        if (textMessage) {
          this.cachedDisplayName = textMessage.senderDisplayName;
          return this.cachedDisplayName;
        }
      } catch {
        continue;
      }
    }

    // Fallback: user properties endpoint
    try {
      const properties = await fetchUserProperties(this.token);
      if (typeof properties.displayname === "string") {
        this.cachedDisplayName = properties.displayname;
        return this.cachedDisplayName;
      }
    } catch {
      // Fall through
    }

    return "Unknown User";
  }
}
