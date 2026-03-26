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
  SmartLoginOptions,
  Conversation,
  Message,
  MessageFormat,
  Member,
  SentMessage,
  EditedMessage,
  DeletedMessage,
  ScheduledMessage,
  GetMessagesOptions,
  ListConversationsOptions,
  OneOnOneSearchResult,
  PersonSearchResult,
  ChatSearchResult,
  MessagesPage,
  TranscriptResult,
  ImageAttachment,
  FileAttachment,
  MessageContentPart,
  UserProfile,
} from "./types.js";
import { SYSTEM_STREAM_TYPES } from "./types.js";
import { isTextMessageType } from "./constants.js";
import {
  fetchConversations,
  fetchMessagesPage,
  fetchMembers,
  postMessage,
  editMessage,
  deleteMessage,
  postScheduledMessage,
  fetchUserProperties,
} from "./api/chat-service.js";
import { ApiAuthError, ApiRateLimitError } from "./api/common.js";
import { fetchProfiles } from "./api/middle-tier.js";
import { searchPeople, searchChats } from "./api/substrate.js";
import { fetchTranscript } from "./api/transcripts.js";
import {
  fetchAmsImage,
  fetchSharePointFile,
  uploadAmsImage,
  uploadSharePointFile,
  buildAmsImageTag,
  buildFilesPropertyJson,
  type SharePointUploadResult,
} from "./api/attachments.js";
import { acquireTokenViaAutoLogin } from "./auth/auto-login.js";
import { acquireTokenViaInteractiveLogin } from "./auth/interactive.js";
import { acquireTokenViaDebugSession } from "./auth/debug-session.js";
import { acquireTokenViaSmartLogin } from "./smart-login.js";
import { DEFAULT_TEAMS_REGION, resolveTeamsRegion } from "./region.js";
import { saveToken, loadToken, clearToken } from "./token-store.js";

export { acquireTokenViaAutoLogin } from "./auth/auto-login.js";
export { acquireTokenViaInteractiveLogin } from "./auth/interactive.js";
export { acquireTokenViaDebugSession } from "./auth/debug-session.js";
export { acquireTokenViaSmartLogin } from "./smart-login.js";
export type {
  TeamsToken,
  AutoLoginOptions,
  InteractiveLoginOptions,
  ManualTokenOptions,
  SmartLoginOptions,
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
  Follower,
  UserProfile,
  PersonSearchResult,
  ChatSearchResult,
  TranscriptEntry,
  TranscriptResult,
  ImageAttachment,
  FileAttachment,
  MessageContentPart,
} from "./types.js";
export { SYSTEM_STREAM_TYPES } from "./types.js";
export { isTextMessageType } from "./constants.js";
export { parseRawMessage } from "./api/chat-service.js";
export { ApiAuthError, ApiRateLimitError } from "./api/common.js";
export { fetchProfiles } from "./api/middle-tier.js";
export { searchPeople, searchChats } from "./api/substrate.js";
export {
  fetchTranscript,
  fetchTranscriptVtt,
  parseVtt,
  extractTranscriptUrl,
  extractMeetingTitle,
  isSuccessfulRecording,
} from "./api/transcripts.js";
export {
  fetchAmsImage,
  uploadAmsImage,
  buildAmsImageTag,
  parseInlineImages,
  parseFileAttachments,
  uploadSharePointFile,
  buildFilesPropertyJson,
} from "./api/attachments.js";
export type { SharePointUploadResult } from "./api/attachments.js";
export { saveToken, loadToken, clearToken } from "./token-store.js";
export { actions } from "./actions/definitions.js";
export { formatOutput } from "./actions/formatters.js";
export type {
  ActionDefinition,
  ActionParameter,
  OutputFormat,
} from "./actions/formatters.js";

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

interface CurrentUserIdentity {
  displayName: string;
  userPrincipalName: string | null;
}

function parseCurrentUserIdentity(
  properties: Record<string, unknown>,
): CurrentUserIdentity | null {
  let displayName =
    typeof properties.displayname === "string" && properties.displayname
      ? properties.displayname
      : null;
  let userPrincipalName: string | null = null;

  if (typeof properties.userDetails === "string") {
    try {
      const userDetails = JSON.parse(properties.userDetails) as {
        name?: string;
        upn?: string;
      };
      if (!displayName && userDetails.name) {
        displayName = userDetails.name;
      }
      if (userDetails.upn) {
        userPrincipalName = userDetails.upn;
      }
    } catch {
      // Ignore malformed userDetails JSON and continue with other fields.
    }
  }

  if (
    !displayName &&
    typeof properties.primaryMemberName === "string" &&
    properties.primaryMemberName
  ) {
    displayName = properties.primaryMemberName;
  }

  if (!displayName && !userPrincipalName) {
    return null;
  }

  return {
    displayName: displayName ?? "Unknown User",
    userPrincipalName,
  };
}

export class TeamsClient {
  private token: TeamsToken;
  private autoLoginOptions: AutoLoginOptions | null = null;
  private smartLoginOptions: SmartLoginOptions | null = null;
  private cachedDisplayName: string | null = null;
  private cachedUserPrincipalName: string | null = null;
  private userEmail: string | null = null;

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
      if (
        !cachedToken.substrateToken ||
        !cachedToken.bearerToken ||
        !cachedToken.amsToken ||
        !cachedToken.sharePointHost
      ) {
        log(
          "Cached token is missing substrate, bearer, AMS token, or SharePoint host, re-authenticating...",
        );
        clearToken(options.email);
        const freshToken = await acquireTokenViaAutoLogin(options);
        saveToken(options.email, freshToken);
        log("Fresh token saved to Keychain");

        const client = new TeamsClient(freshToken);
        client.autoLoginOptions = options;
        client.userEmail = options.email;
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
      client.userEmail = options.email;
      return client;
    }

    log("No cached token found, acquiring via auto-login...");
    const token = await acquireTokenViaAutoLogin(options);
    saveToken(options.email, token);
    log("Token saved to Keychain");

    const client = new TeamsClient(token);
    client.autoLoginOptions = options;
    client.userEmail = options.email;
    return client;
  }

  /**
   * Create a client with smart cross-platform login.
   *
   * Zero-config default: automatically picks the best auth strategy:
   *   - macOS with Chrome + email: tries auto-login (FIDO2), falls back to interactive
   *   - All other cases: interactive browser login (works everywhere)
   *
   * Tokens are cached in the platform credential store and reused
   * until they expire (~23 hours).
   */
  static async connect(options?: SmartLoginOptions): Promise<TeamsClient> {
    const log = options?.verbose ? console.error.bind(console) : () => {};

    // Check for cached token if email is available
    if (options?.email) {
      const cachedToken = loadToken(options.email);
      if (cachedToken) {
        if (
          !cachedToken.substrateToken ||
          !cachedToken.bearerToken ||
          !cachedToken.amsToken ||
          !cachedToken.sharePointHost
        ) {
          log(
            "Cached token is missing required fields, re-authenticating...",
          );
          clearToken(options.email);
        } else {
          const region = resolveTeamsRegion(options.region, cachedToken.region);
          const token =
            region === cachedToken.region
              ? cachedToken
              : { ...cachedToken, region };
          log("Using cached token from credential store");
          if (token !== cachedToken) {
            log(`Overriding cached region with explicit value: ${region}`);
            saveToken(options.email, token);
          }
          const client = new TeamsClient(token);
          client.smartLoginOptions = options;
          client.userEmail = options.email;
          return client;
        }
      }
    }

    log("Acquiring token via smart login...");
    const token = await acquireTokenViaSmartLogin(options);

    if (options?.email) {
      saveToken(options.email, token);
      log("Token saved to credential store");
    }

    const client = new TeamsClient(token);
    client.smartLoginOptions = options ?? null;
    if (options?.email) {
      client.userEmail = options.email;
    }
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
    const client = new TeamsClient(token);
    client.userEmail = options.email;
    return client;
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
    amsToken?: string,
    sharePointToken?: string,
    sharePointHost?: string,
  ): TeamsClient {
    return new TeamsClient({
      skypeToken,
      region,
      bearerToken,
      substrateToken,
      amsToken,
      sharePointToken,
      sharePointHost,
    });
  }

  /**
   * Clear a cached token from the platform credential store.
   *
   * Use this to force a fresh login on the next `create()` or `connect()` call.
   */
  static clearCachedToken(email: string): void {
    clearToken(email);
  }

  /** Get the underlying token (for persistence or debugging). */
  getToken(): TeamsToken {
    return { ...this.token };
  }

  /**
   * Set the user's email address (needed for SharePoint file uploads).
   *
   * Automatically set when using `create()` or `fromAutoLogin()`.
   * Call this manually when using `fromToken()`, `fromDebugSession()`,
   * or `fromInteractiveLogin()` if file upload is needed.
   */
  setEmail(email: string): void {
    this.userEmail = email;
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

      const filtered = excludeSystem
        ? conversations.filter(
            (conversation) =>
              !(SYSTEM_STREAM_TYPES as readonly string[]).includes(
                conversation.threadType,
              ),
          )
        : conversations;

      // Resolve display names for untitled 1:1 chats.
      await this.resolveOneOnOneDisplayNames(filtered);

      return filtered;
    });
  }

  /**
   * Enrich untitled 1:1 conversations with the other member's display name.
   */
  private async resolveOneOnOneDisplayNames(
    conversations: Conversation[],
  ): Promise<void> {
    const untitledOneOnOnes = conversations.filter(
      (conversation) =>
        !conversation.topic &&
        conversation.id.includes("@unq.gbl.spaces") &&
        !conversation.id.startsWith("48:"),
    );

    if (untitledOneOnOnes.length === 0) return;
    const currentUserIdentity = await this.resolveCurrentUserIdentity({
      allowSelfChatFallback: false,
    });
    const currentUserDisplayName = currentUserIdentity.displayName.toLowerCase();

    for (const chat of untitledOneOnOnes) {
      try {
        const members = await this.getMembers(chat.id);
        const namedMembers = members.filter((member) => member.displayName);
        const otherNamedMember = namedMembers.find(
          (member) =>
            member.displayName.toLowerCase() !== currentUserDisplayName,
        );
        if (otherNamedMember) {
          chat.topic = otherNamedMember.displayName;
          continue;
        }
      } catch {
        // Fall through to profile-based lookup.
      }

      if (!currentUserIdentity.userPrincipalName) {
        continue;
      }

      const conversationMatch = chat.id.match(
        /19:([a-f0-9-]+)_([a-f0-9-]+)@unq\.gbl\.spaces/i,
      );
      if (!conversationMatch) {
        continue;
      }

      const memberMris = [conversationMatch[1], conversationMatch[2]].map(
        (uuid) => `8:orgid:${uuid}`,
      );

      let profiles: UserProfile[];
      try {
        profiles = await fetchProfiles(this.token, memberMris);
      } catch {
        continue;
      }

      const otherProfile = profiles.find(
        (profile) =>
          profile.email.toLowerCase() !==
          currentUserIdentity.userPrincipalName?.toLowerCase(),
      );
      if (otherProfile?.displayName) {
        chat.topic = otherProfile.displayName;
      }
    }
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
          const textMessages = page.messages.filter((message) =>
            isTextMessageType(message.messageType),
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
   * Get messages from a conversation.
   *
   * Follows pagination links to fetch message history.
   * Use `limit` to cap the total number of messages returned.
   * Use `maxPages` and `pageSize` to fine-tune pagination behaviour.
   */
  async getMessages(
    conversationId: string,
    options?: GetMessagesOptions,
  ): Promise<Message[]> {
    return this.withTokenRefresh(async () => {
      const limit = options?.limit;
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

        if (limit !== undefined && allMessages.length >= limit) break;
        if (!result.backwardLink) break;
        backwardLink = result.backwardLink;
      }

      await this.enrichMessageParticipantDisplayNames(allMessages);

      if (limit !== undefined && allMessages.length > limit) {
        return allMessages.slice(0, limit);
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
      const messagesPage = await fetchMessagesPage(
        this.token,
        conversationId,
        pageSize,
        backwardLink,
      );
      await this.enrichMessageParticipantDisplayNames(messagesPage.messages);
      return messagesPage;
    });
  }

  private async enrichMessageParticipantDisplayNames(
    messages: Message[],
  ): Promise<void> {
    const unresolvedParticipantMris = new Set<string>();

    for (const message of messages) {
      for (const reaction of message.reactions) {
        for (const reactionUser of reaction.users) {
          if (!reactionUser.displayName) {
            unresolvedParticipantMris.add(reactionUser.mri);
          }
        }
      }

      for (const follower of message.followers) {
        if (!follower.displayName) {
          unresolvedParticipantMris.add(follower.mri);
        }
      }
    }

    if (unresolvedParticipantMris.size === 0) {
      return;
    }

    // Primary: resolve via profile API
    const displayNameByMri = new Map<string, string>();
    try {
      const profiles = await fetchProfiles(
        this.token,
        [...unresolvedParticipantMris],
      );
      for (const profile of profiles) {
        if (profile.displayName) {
          displayNameByMri.set(profile.mri, profile.displayName);
        }
      }
    } catch {
      // Profile API unavailable — fall through to sender-name fallback
    }

    // Fallback: use senderDisplayName from messages for any MRIs the profile
    // API couldn't resolve (e.g. re-provisioned / renamed accounts whose old
    // MRI is no longer in the directory but still appears in chat history).
    const stillUnresolved = new Set<string>();
    for (const mri of unresolvedParticipantMris) {
      if (!displayNameByMri.has(mri)) {
        stillUnresolved.add(mri);
      }
    }

    if (stillUnresolved.size > 0) {
      for (const message of messages) {
        if (stillUnresolved.size === 0) break;
        if (!message.senderDisplayName) continue;
        const senderMri = extractMriSuffix(message.senderMri);
        if (stillUnresolved.has(senderMri)) {
          displayNameByMri.set(senderMri, message.senderDisplayName);
          stillUnresolved.delete(senderMri);
        }
      }
    }

    // Apply resolved names
    for (const message of messages) {
      for (const reaction of message.reactions) {
        for (const reactionUser of reaction.users) {
          if (!reactionUser.displayName) {
            reactionUser.displayName =
              displayNameByMri.get(reactionUser.mri) ?? "";
          }
        }
      }

      for (const follower of message.followers) {
        if (!follower.displayName) {
          follower.displayName = displayNameByMri.get(follower.mri) ?? "";
        }
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
  async sendMessage(
    conversationId: string,
    content: string,
    format: MessageFormat = "markdown",
    amsReferences: string[] = [],
  ): Promise<SentMessage> {
    return this.withTokenRefresh(async () => {
      const displayName = await this.getCurrentUserDisplayName();
      return postMessage(
        this.token,
        conversationId,
        content,
        displayName,
        format,
        amsReferences,
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
   * Schedule a message to be sent at a future time.
   *
   * The `format` parameter controls how `content` is interpreted:
   * - `"text"` — plain text, sent as-is
   * - `"markdown"` (default) — converted from Markdown to HTML
   * - `"html"` — raw HTML, sent as-is
   */
  async scheduleMessage(
    conversationId: string,
    content: string,
    scheduleAt: Date,
    format: MessageFormat = "markdown",
    amsReferences: string[] = [],
  ): Promise<ScheduledMessage> {
    return this.withTokenRefresh(async () => {
      const displayName = await this.getCurrentUserDisplayName();
      return postScheduledMessage(
        this.token,
        conversationId,
        content,
        displayName,
        scheduleAt,
        format,
        amsReferences,
      );
    });
  }

  /**
   * Download an image attachment from the AMS (Async Media Service).
   *
   * @param amsObjectId - The AMS object ID from an ImageAttachment
   * @param fullSize - If true, download the full-size image; otherwise compressed (default: false)
   * @returns Binary image data with content type and size
   */
  async downloadImage(
    amsObjectId: string,
    fullSize = false,
  ): Promise<{ data: Buffer; contentType: string; size: number }> {
    return this.withTokenRefresh(async () => {
      const view = fullSize ? "imgpsh_fullsize_anim" : "imgo";
      return fetchAmsImage(this.token, amsObjectId, view);
    });
  }

  /**
   * Download a file attachment from SharePoint.
   *
   * @param fileUrl - The direct SharePoint file URL from a FileAttachment
   * @param itemId - The SharePoint item ID (unique GUID) from a FileAttachment
   * @returns Binary file data with content type, size, and file name
   */
  async downloadFile(
    fileUrl: string,
    itemId: string,
  ): Promise<{
    data: Buffer;
    contentType: string;
    size: number;
    fileName: string;
  }> {
    return this.withTokenRefresh(async () => {
      return fetchSharePointFile(this.token, fileUrl, itemId);
    });
  }

  /**
   * Send a message with inline images.
   *
   * Uploads each image to AMS, sets permissions for the conversation,
   * and constructs the message HTML with embedded `<img>` tags.
   *
   * Content sections and images are interleaved in order. Use `contentParts`
   * to specify the sequence of text and image content.
   *
   * @example
   *   await client.sendMessageWithImages(conversationId, [
   *     { type: "text", text: "Check out this screenshot:" },
   *     { type: "image", data: imageBuffer, fileName: "screenshot.png", contentType: "image/png" },
   *     { type: "text", text: "And another one:" },
   *     { type: "image", data: image2Buffer, fileName: "screenshot2.png", contentType: "image/png" },
   *   ]);
   */
  async sendMessageWithImages(
    conversationId: string,
    contentParts: MessageContentPart[],
  ): Promise<SentMessage> {
    return this.withTokenRefresh(async () => {
      const amsReferences: string[] = [];
      const htmlParts: string[] = [];

      for (const part of contentParts) {
        if (part.type === "text") {
          htmlParts.push(`<p>${part.text}</p>`);
        } else if (part.type === "image") {
          const { amsObjectId } = await uploadAmsImage(
            this.token,
            part.data,
            part.fileName,
            conversationId,
          );
          amsReferences.push(amsObjectId);
          htmlParts.push(
            buildAmsImageTag(amsObjectId, part.width, part.height),
          );
        }
      }

      const content = `<div>${htmlParts.join("")}</div>`;
      const displayName = await this.getCurrentUserDisplayName();
      return postMessage(
        this.token,
        conversationId,
        content,
        displayName,
        "html",
        amsReferences,
      );
    });
  }

  /**
   * Send a message with file attachments (uploaded to SharePoint).
   *
   * Uploads each file to the sender's OneDrive "Microsoft Teams Chat Files"
   * folder, then sends a message referencing the uploaded files via
   * `properties.files` JSON.
   *
   * Content sections, images, and files are handled together. Images are
   * uploaded to AMS (inline), files to SharePoint (as attachments).
   *
   * Requires a SharePoint token (captured during authentication) and the
   * user's email (set via `setEmail()` or automatically during `create()`).
   *
   * @example
   *   await client.sendMessageWithFiles(conversationId, [
   *     { type: "text", text: "Here's the document:" },
   *     { type: "file", data: fileBuffer, fileName: "report.md" },
   *   ]);
   */
  async sendMessageWithFiles(
    conversationId: string,
    contentParts: MessageContentPart[],
  ): Promise<SentMessage> {
    return this.withTokenRefresh(async () => {
      if (!this.userEmail) {
        throw new Error(
          "User email is required for file upload but is not set. " +
            "Call setEmail() on the client or use TeamsClient.create() which sets it automatically.",
        );
      }

      const amsReferences: string[] = [];
      const htmlParts: string[] = [];
      const uploadResults: SharePointUploadResult[] = [];

      for (const part of contentParts) {
        if (part.type === "text") {
          htmlParts.push(`<p>${part.text}</p>`);
        } else if (part.type === "image") {
          const { amsObjectId } = await uploadAmsImage(
            this.token,
            part.data,
            part.fileName,
            conversationId,
          );
          amsReferences.push(amsObjectId);
          htmlParts.push(
            buildAmsImageTag(amsObjectId, part.width, part.height),
          );
        } else if (part.type === "file") {
          const result = await uploadSharePointFile(
            this.token,
            part.data,
            part.fileName,
            this.userEmail,
          );
          uploadResults.push(result);
        }
      }

      const content =
        htmlParts.length > 0 ? `<div>${htmlParts.join("")}</div>` : "";
      const filesJson = buildFilesPropertyJson(uploadResults);
      const displayName = await this.getCurrentUserDisplayName();
      return postMessage(
        this.token,
        conversationId,
        content,
        displayName,
        "html",
        amsReferences,
        filesJson,
      );
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
   */
  async getCurrentUserDisplayName(): Promise<string> {
    return this.withTokenRefresh(async () => {
      return (await this.resolveCurrentUserIdentity()).displayName;
    });
  }

  private async resolveCurrentUserIdentity(options?: {
    allowSelfChatFallback?: boolean;
  }): Promise<CurrentUserIdentity> {
    if (this.cachedDisplayName) {
      return {
        displayName: this.cachedDisplayName,
        userPrincipalName: this.cachedUserPrincipalName,
      };
    }

    try {
      const properties = await fetchUserProperties(this.token);
      const currentUserIdentity = parseCurrentUserIdentity(properties);
      if (currentUserIdentity) {
        this.cachedDisplayName = currentUserIdentity.displayName;
        this.cachedUserPrincipalName = currentUserIdentity.userPrincipalName;
        return currentUserIdentity;
      }
    } catch {
      // Fall through to self-chat fallback.
    }

    if (options?.allowSelfChatFallback === false) {
      return {
        displayName: "Unknown User",
        userPrincipalName: null,
      };
    }

    const conversations = await fetchConversations(this.token, 10);

    for (const conversation of conversations) {
      try {
        if (!conversation.id.startsWith("48:notes")) continue;

        const page = await fetchMessagesPage(this.token, conversation.id, 10);
        const textMessage = page.messages.find(
          (message) =>
            isTextMessageType(message.messageType) &&
            message.senderDisplayName.length > 0,
        );

        if (textMessage) {
          this.cachedDisplayName = textMessage.senderDisplayName;
          this.cachedUserPrincipalName = null;
          return {
            displayName: textMessage.senderDisplayName,
            userPrincipalName: null,
          };
        }
      } catch {
        continue;
      }
    }

    return {
      displayName: "Unknown User",
      userPrincipalName: null,
    };
  }

  /**
   * Re-acquire the token and update the cached version.
   *
   * Called automatically by `withTokenRefresh` on 401 errors.
   * Supports both auto-login and smart login paths.
   */
  private async refreshToken(): Promise<void> {
    if (this.smartLoginOptions) {
      if (this.smartLoginOptions.email) {
        clearToken(this.smartLoginOptions.email);
      }
      const freshToken = await acquireTokenViaSmartLogin(
        this.smartLoginOptions,
      );
      if (this.smartLoginOptions.email) {
        saveToken(this.smartLoginOptions.email, freshToken);
      }
      this.token = freshToken;
      this.cachedDisplayName = null;
      this.cachedUserPrincipalName = null;
      return;
    }

    if (!this.autoLoginOptions) {
      throw new Error(
        "Cannot refresh token: no auto-login or smart-login options configured",
      );
    }

    clearToken(this.autoLoginOptions.email);
    const freshToken = await acquireTokenViaAutoLogin(this.autoLoginOptions);
    saveToken(this.autoLoginOptions.email, freshToken);
    this.token = freshToken;
    this.cachedDisplayName = null;
    this.cachedUserPrincipalName = null;
  }

  /**
   * Wrap an operation with automatic token refresh on 401.
   *
   * If the operation throws an `ApiAuthError` and auto-login or
   * smart-login options are configured, the token is re-acquired
   * and the operation retried exactly once.
   */
  private async withTokenRefresh<T>(operation: () => Promise<T>): Promise<T> {
    try {
      return await operation();
    } catch (error) {
      if (
        error instanceof ApiAuthError &&
        (this.autoLoginOptions || this.smartLoginOptions)
      ) {
        await this.refreshToken();
        return await operation();
      }
      throw error;
    }
  }
}
