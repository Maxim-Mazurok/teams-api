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
 *   const client = TeamsClient.fromToken(
 *     "skype-token-here",
 *     "apac",
 *     "optional-bearer-token",
 *     "optional-substrate-token",
 *   );
 */

import type {
  TeamsToken,
  AutoLoginOptions,
  InteractiveLoginOptions,
  ManualTokenOptions,
  Conversation,
  Message,
  MessageFormat,
  Member,
  SentMessage,
  EditedMessage,
  DeletedMessage,
  GetMessagesOptions,
  ListConversationsOptions,
  OneOnOneSearchResult,
  PersonSearchResult,
  ChatSearchResult,
  MessagesPage,
  TranscriptResult,
  SYSTEM_STREAM_TYPES,
} from "./types.js";
import {
  fetchConversations,
  fetchMessagesPage,
  fetchMembers,
  fetchProfiles,
  postMessage,
  editMessage,
  deleteMessage,
  fetchUserProperties,
  fetchTranscript,
  searchPeople,
  searchChats,
  ApiAuthError,
  ApiRateLimitError,
} from "./api.js";
import {
  acquireTokenViaAutoLogin,
  acquireTokenViaInteractiveLogin,
  acquireTokenViaDebugSession,
} from "./auth.js";
import { DEFAULT_TEAMS_REGION, resolveTeamsRegion } from "./region.js";
import { saveToken, loadToken, clearToken } from "./token-store.js";

export {
  acquireTokenViaAutoLogin,
  acquireTokenViaInteractiveLogin,
  acquireTokenViaDebugSession,
} from "./auth.js";
export type {
  TeamsToken,
  AutoLoginOptions,
  InteractiveLoginOptions,
  ManualTokenOptions,
  Conversation,
  Message,
  MessageFormat,
  Member,
  SentMessage,
  EditedMessage,
  DeletedMessage,
  GetMessagesOptions,
  ListConversationsOptions,
  OneOnOneSearchResult,
  MessagesPage,
  Mention,
  Reaction,
  UserProfile,
  PersonSearchResult,
  ChatSearchResult,
  TranscriptEntry,
  TranscriptResult,
} from "./types.js";
export { SYSTEM_STREAM_TYPES } from "./types.js";
export {
  parseRawMessage,
  ApiAuthError,
  ApiRateLimitError,
  fetchProfiles,
  searchPeople,
  searchChats,
  fetchTranscript,
  fetchTranscriptVtt,
  parseVtt,
  extractTranscriptUrl,
  extractMeetingTitle,
  isSuccessfulRecording,
} from "./api.js";
export { saveToken, loadToken, clearToken } from "./token-store.js";
export { actions, formatOutput } from "./actions.js";
export type {
  ActionDefinition,
  ActionParameter,
  OutputFormat,
} from "./actions.js";

const SYSTEM_STREAMS: readonly string[] = [
  "streamofannotations",
  "streamofthreads",
  "streamofnotifications",
  "streamofmentions",
  "streamofnotes",
];

/**
 * Extract the MRI suffix from a sender URL or return the input unchanged.
 * API messages use full URLs (e.g. ".../contacts/8:orgid:uuid"),
 * while member IDs are just the MRI suffix ("8:orgid:uuid").
 */
function extractMriSuffix(senderMri: string): string {
  const contactsPrefix = "/contacts/";
  const contactsIndex = senderMri.lastIndexOf(contactsPrefix);
  if (contactsIndex >= 0) {
    return senderMri.slice(contactsIndex + contactsPrefix.length);
  }
  return senderMri;
}

export class TeamsClient {
  private token: TeamsToken;
  private autoLoginOptions: AutoLoginOptions | null = null;
  private cachedDisplayName: string | null = null;

  private constructor(token: TeamsToken) {
    this.token = token;
  }

  /**
   * Create a client with secure token caching and automatic refresh.
   *
   * This is the recommended entry point. It:
   *   1. Checks the macOS Keychain for a cached, non-expired token
   *   2. If none found, launches auto-login to acquire a fresh token
   *   3. Saves the token to the Keychain for future use
   *   4. Automatically re-acquires the token if a 401 occurs mid-session
   *
   * Token lifetime is ~24 hours. Cached tokens are reused within 23 hours.
   */
  static async create(options: AutoLoginOptions): Promise<TeamsClient> {
    const log = options.verbose ? console.error.bind(console) : () => {};

    const cachedToken = loadToken(options.email);
    if (cachedToken) {
      // Re-authenticate if the cached token is missing substrate or bearer
      // tokens — these are required for people/chat search and profile
      // resolution. Tokens captured before these were added will be stale.
      if (!cachedToken.substrateToken || !cachedToken.bearerToken) {
        log(
          "Cached token is missing substrate or bearer token, re-authenticating...",
        );
        clearToken(options.email);
        const freshToken = await acquireTokenViaAutoLogin(options);
        saveToken(options.email, freshToken);
        log("Fresh token saved to Keychain");

        const client = new TeamsClient(freshToken);
        client.autoLoginOptions = options;
        return client;
      }

      const region = resolveTeamsRegion(options.region, cachedToken.region);
      const token =
        region === cachedToken.region
          ? cachedToken
          : { ...cachedToken, region };

      log("Using cached token from Keychain");
      if (token !== cachedToken) {
        log(`Overriding cached region with explicit value: ${region}`);
        saveToken(options.email, token);
      }

      const client = new TeamsClient(token);
      client.autoLoginOptions = options;
      return client;
    }

    log("No cached token found, acquiring via auto-login...");
    const token = await acquireTokenViaAutoLogin(options);
    saveToken(options.email, token);
    log("Token saved to Keychain");

    const client = new TeamsClient(token);
    client.autoLoginOptions = options;
    return client;
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
   * Create a client via interactive browser login.
   *
   * Opens a visible Chromium window where the user manually logs into
   * Teams. Works on all platforms (macOS, Windows, Linux) without
   * requiring FIDO2 passkeys or system Chrome.
   */
  static async fromInteractiveLogin(
    options?: InteractiveLoginOptions,
  ): Promise<TeamsClient> {
    const token = await acquireTokenViaInteractiveLogin(options);
    return new TeamsClient(token);
  }

  /**
   * Create a client from an existing token bundle.
   *
   * `bearerToken` enables profile/member resolution and `substrateToken`
   * enables reliable people/chat search. Both are optional; basic chat
   * operations only require `skypeToken`.
   */
  static fromToken(
    skypeToken: string,
    region = DEFAULT_TEAMS_REGION,
    bearerToken?: string,
    substrateToken?: string,
  ): TeamsClient {
    return new TeamsClient({ skypeToken, region, bearerToken, substrateToken });
  }

  /**
   * Clear a cached token from the macOS Keychain.
   *
   * Use this to force a fresh login on the next `create()` call.
   */
  static clearCachedToken(email: string): void {
    clearToken(email);
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
    return this.withTokenRefresh(async () => {
      const pageSize = options?.pageSize ?? 50;
      const excludeSystem = options?.excludeSystemStreams ?? true;

      const conversations = await fetchConversations(this.token, pageSize);

      if (!excludeSystem) {
        return conversations;
      }

      return conversations.filter(
        (conversation) => !SYSTEM_STREAMS.includes(conversation.threadType),
      );
    });
  }

  /**
   * Find a conversation by topic name (case-insensitive partial match).
   *
   * When a Substrate token is available, also searches via the Substrate
   * chat search API for matches by name or member. Falls back to local
   * topic matching when Substrate is unavailable.
   */
  async findConversation(query: string): Promise<Conversation | null> {
    return this.withTokenRefresh(async () => {
      const conversations = await this.listConversations({ pageSize: 100 });
      const queryLower = query.toLowerCase();

      // Try local topic matching first (fast, no extra API call)
      const topicMatch = conversations.find(
        (conversation) =>
          conversation.topic &&
          conversation.topic.toLowerCase().includes(queryLower),
      );
      if (topicMatch) return topicMatch;

      // Use Substrate chat search for broader matching (by member names, etc.)
      try {
        const chatResults = await searchChats(this.token, query, 5);
        for (const chatResult of chatResults) {
          const matchingConversation = conversations.find(
            (conversation) => conversation.id === chatResult.threadId,
          );
          if (matchingConversation) return matchingConversation;
        }
      } catch {
        // Substrate unavailable — local topic match above was the only option
      }

      return null;
    });
  }

  /**
   * Search for people in the organization directory by name.
   *
   * Primary: Substrate search API (requires Substrate token).
   * Fallback: resolves members of recent conversations via the
   * middle-tier profile API (requires Bearer token). Extracts
   * member UUIDs directly from 1:1 conversation IDs (zero extra
   * API calls for those) and scans a limited set of group chats.
   */
  async findPeople(
    query: string,
    maxResults = 10,
  ): Promise<PersonSearchResult[]> {
    return this.withTokenRefresh(async () => {
      let authError: ApiAuthError | null = null;

      try {
        const substrateResults = await searchPeople(
          this.token,
          query,
          maxResults,
        );
        if (substrateResults.length > 0) {
          return substrateResults;
        }
      } catch (error) {
        if (error instanceof ApiAuthError) {
          authError = error;
        }
        // Substrate unavailable — fall through to profile-based search
      }

      // Fallback: search members of recent conversations via profiles

      const queryLower = query.toLowerCase();
      const conversations = await fetchConversations(this.token, 100);
      const memberMris = new Set<string>();

      // Step 1: Extract UUIDs from 1:1 conversation IDs (no API calls needed)
      for (const conversation of conversations) {
        const match = conversation.id.match(
          /19:([a-f0-9-]+)_([a-f0-9-]+)@unq\.gbl\.spaces/i,
        );
        if (match) {
          memberMris.add(`8:orgid:${match[1]}`);
          memberMris.add(`8:orgid:${match[2]}`);
        }
      }

      // Step 2: Scan a limited set of group chats for their members
      const maxGroupChatsToScan = 10;
      let groupChatsScanned = 0;
      for (const conversation of conversations) {
        if (groupChatsScanned >= maxGroupChatsToScan) break;
        if (
          conversation.id.includes("@unq.gbl.spaces") ||
          conversation.id.startsWith("48:")
        ) {
          continue;
        }
        if (
          conversation.threadType === "chat" ||
          conversation.threadType === "topic"
        ) {
          try {
            const members = await fetchMembers(this.token, conversation.id);
            for (const member of members) {
              if (member.id.startsWith("8:orgid:")) {
                memberMris.add(member.id);
              }
            }
            groupChatsScanned++;
          } catch {
            continue;
          }
        }
      }

      if (memberMris.size === 0) {
        if (authError) throw authError;
        return [];
      }

      const profiles = await fetchProfiles(this.token, [...memberMris]);
      const matchingProfiles = profiles
        .filter((profile) =>
          profile.displayName.toLowerCase().includes(queryLower),
        )
        .slice(0, maxResults);

      return matchingProfiles.map((profile) => ({
        displayName: profile.displayName,
        mri: profile.mri,
        email: profile.email,
        jobTitle: profile.jobTitle,
        department: "",
        objectId: profile.mri.replace("8:orgid:", ""),
      }));
    });
  }

  /**
   * Search for chats by name or member name.
   *
   * Primary: Substrate search API (requires Substrate token).
   * Fallback: local topic matching on listed conversations.
   */
  async findChats(query: string, maxResults = 10): Promise<ChatSearchResult[]> {
    return this.withTokenRefresh(async () => {
      let authError: ApiAuthError | null = null;

      try {
        const substrateResults = await searchChats(
          this.token,
          query,
          maxResults,
        );
        if (substrateResults.length > 0) {
          return substrateResults;
        }
      } catch (error) {
        if (error instanceof ApiAuthError) {
          authError = error;
        }
        // Substrate unavailable — fall through to local topic matching
      }

      // Fallback: match conversation topics locally
      const conversations = await this.listConversations({ pageSize: 100 });
      const queryLower = query.toLowerCase();
      const matchingChats: ChatSearchResult[] = [];

      for (const conversation of conversations) {
        if (
          conversation.topic &&
          conversation.topic.toLowerCase().includes(queryLower)
        ) {
          matchingChats.push({
            name: conversation.topic,
            threadId: conversation.id,
            threadType: conversation.threadType,
            matchingMembers: [],
            chatMembers: [],
            totalMemberCount: conversation.memberCount ?? 0,
          });
          if (matchingChats.length >= maxResults) break;
        }
      }

      if (matchingChats.length === 0 && authError) throw authError;
      return matchingChats;
    });
  }

  /**
   * Find a 1:1 conversation with a specific person.
   *
   * Tries multiple strategies in order:
   *   1. Self-chat check (if the name matches the current user)
   *   2. Substrate people/chat search (if Substrate token available)
   *   3. Profile-based 1:1 matching (if Bearer token available) —
   *      extracts member UUIDs from conversation IDs and resolves
   *      display names via the middle-tier profile API
   *   4. Message sender scanning (slowest, always available)
   */
  async findOneOnOneConversation(
    personName: string,
  ): Promise<OneOnOneSearchResult | null> {
    return this.withTokenRefresh(async () => {
      const targetLower = personName.toLowerCase();
      let authError: ApiAuthError | null = null;

      // Check self-chat first
      const conversations = await fetchConversations(this.token, 100);
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

      // Strategy 1: Substrate search API for people + chat lookup
      try {
        const people = await searchPeople(this.token, personName, 5);
        const matchedPerson = people.find((person) =>
          person.displayName.toLowerCase().includes(targetLower),
        );

        if (matchedPerson) {
          const chats = await searchChats(this.token, personName, 10);

          // Look for a 1:1 chat (Chat type, exactly 2 members)
          for (const chat of chats) {
            if (
              chat.threadType === "Chat" &&
              chat.totalMemberCount === 2 &&
              chat.matchingMembers.some(
                (member) => member.mri === matchedPerson.mri,
              )
            ) {
              return {
                conversationId: chat.threadId,
                memberDisplayName: matchedPerson.displayName,
              };
            }
          }

          // Check if any conversation ID contains the person's UUID
          const personUuid = matchedPerson.mri.replace("8:orgid:", "");
          const matchingConversation = conversations.find(
            (conversation) =>
              conversation.id.includes(personUuid) &&
              conversation.threadType === "chat",
          );
          if (matchingConversation) {
            return {
              conversationId: matchingConversation.id,
              memberDisplayName: matchedPerson.displayName,
            };
          }
        }
      } catch (error) {
        if (error instanceof ApiAuthError) {
          authError = error;
        }
        // Substrate unavailable — fall through to profile-based matching
      }

      // Strategy 2: Profile-based matching for 1:1 chats (uses Bearer token)
      {
        const oneOnOneChats = conversations.filter(
          (conversation) =>
            conversation.id.includes("@unq.gbl.spaces") &&
            !conversation.id.startsWith("48:"),
        );

        // Extract all member UUIDs from 1:1 conversation IDs
        // Format: 19:{uuid1}_{uuid2}@unq.gbl.spaces
        const uuidToConversationId = new Map<string, string>();
        for (const chat of oneOnOneChats) {
          const match = chat.id.match(
            /19:([a-f0-9-]+)_([a-f0-9-]+)@unq\.gbl\.spaces/i,
          );
          if (match) {
            uuidToConversationId.set(match[1], chat.id);
            uuidToConversationId.set(match[2], chat.id);
          }
        }

        if (uuidToConversationId.size > 0) {
          const mris = [...uuidToConversationId.keys()].map(
            (uuid) => `8:orgid:${uuid}`,
          );
          try {
            const profiles = await fetchProfiles(this.token, mris);
            for (const profile of profiles) {
              if (profile.displayName.toLowerCase().includes(targetLower)) {
                const uuid = profile.mri.replace("8:orgid:", "");
                const conversationId = uuidToConversationId.get(uuid);
                if (conversationId) {
                  return {
                    conversationId,
                    memberDisplayName: profile.displayName,
                  };
                }
              }
            }
          } catch (error) {
            if (error instanceof ApiAuthError) {
              authError = error;
            }
            // Profile resolution failed, fall through to message scanning
          }
        }
      }

      // Strategy 3: scan message senders (slowest fallback)
      const untitledChats = conversations.filter(
        (conversation) =>
          conversation.threadType === "chat" &&
          !conversation.topic &&
          !conversation.id.startsWith("48:"),
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

      if (authError) throw authError;
      return null;
    });
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
    return this.withTokenRefresh(async () => {
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
    });
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
    return this.withTokenRefresh(async () => {
      return fetchMessagesPage(
        this.token,
        conversationId,
        pageSize,
        backwardLink,
      );
    });
  }

  /**
   * Send a message to a conversation.
   *
   * The `format` parameter controls how `content` is interpreted:
   * - `"text"` — plain text, sent as-is
   * - `"markdown"` (default) — converted from Markdown to HTML
   * - `"html"` — raw HTML, sent as-is
   */
  async sendMessage(
    conversationId: string,
    content: string,
    format: MessageFormat = "markdown",
  ): Promise<SentMessage> {
    return this.withTokenRefresh(async () => {
      const displayName = await this.getCurrentUserDisplayName();
      return postMessage(
        this.token,
        conversationId,
        content,
        displayName,
        format,
      );
    });
  }

  /**
   * Edit an existing message in a conversation.
   *
   * The `format` parameter controls how `content` is interpreted:
   * - `"text"` — plain text, sent as-is
   * - `"markdown"` (default) — converted from Markdown to HTML
   * - `"html"` — raw HTML, sent as-is
   */
  async editMessage(
    conversationId: string,
    messageId: string,
    content: string,
    format: MessageFormat = "markdown",
  ): Promise<EditedMessage> {
    return this.withTokenRefresh(async () => {
      const displayName = await this.getCurrentUserDisplayName();
      return editMessage(
        this.token,
        conversationId,
        messageId,
        content,
        displayName,
        format,
      );
    });
  }

  /** Delete a message from a conversation. */
  async deleteMessage(
    conversationId: string,
    messageId: string,
  ): Promise<DeletedMessage> {
    return this.withTokenRefresh(async () => {
      return deleteMessage(this.token, conversationId, messageId);
    });
  }

  /**
   * Get members of a conversation.
   *
   * Display names are resolved via the Teams middle-tier profile API when a
   * Bearer token is available. Falls back to scanning message history when it
   * is not.
   */
  async getMembers(conversationId: string): Promise<Member[]> {
    return this.withTokenRefresh(async () => {
      const members = await fetchMembers(this.token, conversationId);

      const unresolvedPeople = members.filter(
        (member) => member.memberType === "person" && !member.displayName,
      );

      if (unresolvedPeople.length === 0) {
        return members;
      }

      // Primary: resolve via middle-tier profile API (requires bearerToken)
      {
        const mris = unresolvedPeople.map((member) => member.id);
        try {
          const profiles = await fetchProfiles(this.token, mris);
          const profileLookup = new Map(
            profiles.map((profile) => [profile.mri, profile.displayName]),
          );
          for (const member of members) {
            if (!member.displayName) {
              member.displayName = profileLookup.get(member.id) ?? "";
            }
          }
          return members;
        } catch {
          // Fall through to message-history fallback
        }
      }

      // Fallback: resolve from message sender names in conversation history
      const unresolvedMris = new Set(
        unresolvedPeople.map((member) => member.id),
      );
      const nameLookup = new Map<string, string>();
      const maxPages = 10;
      let backwardLink: string | undefined;

      try {
        for (let page = 0; page < maxPages; page++) {
          const result = await fetchMessagesPage(
            this.token,
            conversationId,
            200,
            backwardLink,
          );

          for (const message of result.messages) {
            if (message.senderDisplayName) {
              const mriSuffix = extractMriSuffix(message.senderMri);
              if (unresolvedMris.has(mriSuffix)) {
                nameLookup.set(mriSuffix, message.senderDisplayName);
                unresolvedMris.delete(mriSuffix);
              }
            }
          }

          if (unresolvedMris.size === 0 || !result.backwardLink) {
            break;
          }
          backwardLink = result.backwardLink;
        }
      } catch {
        // If message fetch fails, return members with whatever names we have
      }

      for (const member of members) {
        if (!member.displayName) {
          member.displayName = nameLookup.get(member.id) ?? "";
        }
      }

      return members;
    });
  }

  /**
   * Get the meeting transcript for a conversation.
   *
   * Searches messages for a successful recording, extracts the AMS
   * transcript URL, fetches the VTT, and parses it into structured entries.
   */
  async getTranscript(conversationId: string): Promise<TranscriptResult> {
    return this.withTokenRefresh(async () => {
      return fetchTranscript(this.token, conversationId);
    });
  }

  /**
   * Get the display name of the currently authenticated user.
   *
   * Resolved by reading messages from the self-chat (48:notes) or
   * falling back to the user properties endpoint.
   */
  async getCurrentUserDisplayName(): Promise<string> {
    return this.withTokenRefresh(async () => {
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

      // Fallback: user properties endpoint (userDetails JSON field)
      try {
        const properties = await fetchUserProperties(this.token);
        if (typeof properties.displayname === "string") {
          this.cachedDisplayName = properties.displayname;
          return this.cachedDisplayName;
        }
        if (typeof properties.userDetails === "string") {
          try {
            const userDetails = JSON.parse(properties.userDetails) as {
              name?: string;
            };
            if (userDetails.name) {
              this.cachedDisplayName = userDetails.name;
              return this.cachedDisplayName;
            }
          } catch {
            // Malformed JSON — fall through
          }
        }
      } catch {
        // Fall through
      }

      return "Unknown User";
    });
  }

  /**
   * Re-acquire the token via auto-login and update the cached version.
   *
   * Called automatically by `withTokenRefresh` on 401 errors.
   */
  private async refreshToken(): Promise<void> {
    if (!this.autoLoginOptions) {
      throw new Error("Cannot refresh token: no auto-login options configured");
    }

    clearToken(this.autoLoginOptions.email);
    const freshToken = await acquireTokenViaAutoLogin(this.autoLoginOptions);
    saveToken(this.autoLoginOptions.email, freshToken);
    this.token = freshToken;
    this.cachedDisplayName = null;
  }

  /**
   * Wrap an operation with automatic token refresh on 401.
   *
   * If the operation throws an `ApiAuthError` and auto-login options
   * are configured, the token is re-acquired and the operation retried
   * exactly once.
   */
  private async withTokenRefresh<T>(operation: () => Promise<T>): Promise<T> {
    try {
      return await operation();
    } catch (error) {
      if (error instanceof ApiAuthError && this.autoLoginOptions) {
        await this.refreshToken();
        return await operation();
      }
      throw error;
    }
  }
}
