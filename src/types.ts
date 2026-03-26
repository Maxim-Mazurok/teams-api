/**
 * Public type definitions for the Teams API client.
 *
 * All types used by the public API surface are defined here.
 */

/** Authentication token for Teams Chat Service REST API. */
export interface TeamsToken {
  /** The skype token used in the Authentication header. */
  skypeToken: string;
  /** API region (e.g. "apac", "emea", "amer"). */
  region: string;
  /**
   * OAuth2 Bearer token for the Teams middle-tier API (api.spaces.skype.com audience).
   * Used for profile resolution. Optional — only available when captured during auth.
   */
  bearerToken?: string;
  /**
   * OAuth2 Bearer token for the Substrate search API (substrate.office.com audience).
   * Used for people/chat/channel search. Optional — only available when captured during auth.
   */
  substrateToken?: string;
  /**
   * OAuth2 Bearer token for the AMS (Async Media Service) (ic3.teams.office.com audience).
   * Used for image upload. Optional — only available when captured during auth.
   */
  amsToken?: string;
  /**
   * OAuth2 Bearer token for SharePoint (*.sharepoint.com audience).
   * Used for downloading file attachments. Optional — only available when captured during auth.
   */
  sharePointToken?: string;
  /**
   * The SharePoint host for the user's OneDrive for Business.
   * Example: `contoso-my.sharepoint.com`.
   * Used for file upload. Captured from the MSAL token cache during auth.
   */
  sharePointHost?: string;
}

/** Options for automatic token acquisition via FIDO2 passkey. */
export interface AutoLoginOptions {
  /** Corporate email for Microsoft Entra ID login. */
  email: string;
  /** Explicit API region override. Omit to auto-detect when possible. */
  region?: string;
  /** Path to Chrome executable (defaults to system Chrome on macOS). */
  chromePath?: string;
  /** Directory for temporary browser profile (cleaned automatically). */
  profileDirectory?: string;
  /** Run browser in headless mode (default: true). */
  headless?: boolean;
  /** Emit progress messages to console (default: false). */
  verbose?: boolean;
}

/** Options for interactive browser login (all platforms). */
export interface InteractiveLoginOptions {
  /** Explicit API region override. Omit to auto-detect when possible. */
  region?: string;
  /** Corporate email to pre-fill on the login page (optional). */
  email?: string;
  /** Emit progress messages to console (default: false). */
  verbose?: boolean;
}

/** Options for manual token capture from a running Chrome debug session. */
export interface ManualTokenOptions {
  /** Chrome DevTools Protocol debug port (default: 9222). */
  debugPort?: number;
  /** Explicit API region override. Omit to auto-detect when possible. */
  region?: string;
}

/** A Teams conversation (chat, group chat, meeting, or channel). */
export interface Conversation {
  /** Unique conversation thread ID. */
  id: string;
  /** Display name or topic of the conversation. */
  topic: string;
  /** Type of conversation thread (e.g. "chat", "topic", "meeting"). */
  threadType: string;
  /** Server version number for the conversation. */
  version: number;
  /** ISO timestamp of the last received message, or null. */
  lastMessageTime: string | null;
  /** Number of members, or null if unknown. */
  memberCount: number | null;
}

/**
 * A single message in a conversation.
 *
 * Includes text/rich-text messages, system events, and control messages.
 * Use the `messageType` field to filter to the desired category.
 */
export interface Message {
  /** Server-assigned message ID. */
  id: string;
  /**
   * Message type identifier.
   * Common values: "RichText/Html", "Text", "ThreadActivity/AddMember",
   * "ThreadActivity/MemberJoined", "Event/Call", "MessageDelete".
   */
  messageType: string;
  /** Full MRI (Microsoft Resource Identifier) of the sender. */
  senderMri: string;
  /** Display name of the sender. */
  senderDisplayName: string;
  /** Message content (HTML for RichText/Html, plain text for Text). */
  content: string;
  /** ISO timestamp when the message originally arrived at the server. */
  originalArrivalTime: string;
  /** ISO timestamp when the message was composed. */
  composeTime: string;
  /** ISO timestamp of the last edit, or null if never edited. */
  editTime: string | null;
  /** Message subject line, or null. */
  subject: string | null;
  /** Whether the message has been deleted. */
  isDeleted: boolean;
  /** Reactions on this message (excludes follow subscriptions). */
  reactions: Reaction[];
  /** Users who followed (subscribed to) this message thread. */
  followers: Follower[];
  /** Users mentioned in this message. */
  mentions: Mention[];
  /** ID of the quoted/replied-to message, or null. */
  quotedMessageId: string | null;
  /** Inline images embedded in the message content (from AMS). */
  images: ImageAttachment[];
  /** File attachments (SharePoint-hosted documents, videos, etc.). */
  files: FileAttachment[];
}

/**
 * An inline image embedded in a message via the Async Media Service (AMS).
 *
 * Images are stored as `<img>` tags in the message HTML content. The AMS object
 * can be fetched with `Authorization: skype_token <token>`. Multiple image views
 * are available at different resolutions.
 */
export interface ImageAttachment {
  /** AMS object ID (e.g. "0-eaua-d2-877b82634f4e978692f2243d445a6650"). */
  amsObjectId: string;
  /** URL to the compressed/optimized image view. */
  url: string;
  /** Full-size image URL (imgpsh_fullsize_anim view). */
  fullSizeUrl: string;
  /** Image width in pixels, or null if not specified. */
  width: number | null;
  /** Image height in pixels, or null if not specified. */
  height: number | null;
  /** Position of this image in the content HTML (character offset of the `<img>` tag). */
  contentPosition: number;
}

/**
 * A file attachment hosted on SharePoint.
 *
 * File attachments include documents, videos, and other non-image files shared
 * through Teams. These are stored on SharePoint and referenced in `properties.files`.
 */
export interface FileAttachment {
  /** SharePoint item ID. */
  itemId: string;
  /** Original file name (e.g. "report.pdf"). */
  fileName: string;
  /** File extension without dot (e.g. "pdf", "docx", "mov"). */
  fileType: string;
  /** Direct SharePoint file URL. */
  fileUrl: string;
  /** Shareable link URL (accessible to conversation members). */
  shareUrl: string;
}

/** A reaction (emotion) on a message. */
export interface Reaction {
  /** Reaction key (e.g. "like", "heart", "laugh"). */
  key: string;
  /** Users who reacted with this emotion. */
  users: Array<{ mri: string; time: number; displayName?: string }>;
}

/**
 * A user who followed (subscribed to) a message thread.
 *
 * Teams stores "follow" as an emotion entry in `properties.emotions`.
 * Unlike regular reactions, the `value` field is "0" (following) or "1" (unfollowed),
 * not a message timestamp.
 */
export interface Follower {
  /** Full MRI of the follower. */
  mri: string;
  /** Epoch milliseconds when the follow action was recorded. */
  time: number;
  /** Display name when resolved from the profile API. */
  displayName?: string;
}

/** A user mention in a message. */
export interface Mention {
  /** MRI or tag ID of the mentioned user. */
  id: string;
  /** Display name of the mentioned user. */
  displayName: string;
}

/** A member of a conversation. */
export interface Member {
  /** Full MRI of the member. */
  id: string;
  /** Display name of the member. */
  displayName: string;
  /** Role in the conversation (e.g. "Admin", "User"). */
  role: string;
  /** Whether this member is a person or a bot/app (detected from MRI prefix). */
  memberType: "person" | "bot";
}

/** A user profile resolved from the Teams middle-tier API. */
export interface UserProfile {
  /** Full MRI of the user. */
  mri: string;
  /** Display name. */
  displayName: string;
  /** Email address. */
  email: string;
  /** Job title, or empty string. */
  jobTitle: string;
  /** User type (e.g. "Member", "Guest"). */
  userType: string;
}

/** Format for sending messages. */
export type MessageFormat = "text" | "markdown" | "html";

/**
 * Controls who gets access to uploaded SharePoint files.
 *
 * - `"chat"` — share with chat participants only (default).
 * - `"organization"` — anyone in the organization with the link.
 * - `"none"` — no sharing link; only the uploader can access the file.
 */
export type FileSharingScope = "chat" | "organization" | "none";

/** A part of a message body when sending messages with inline images or file attachments. */
export type MessageContentPart =
  | { type: "text"; text: string }
  | {
      type: "image";
      data: Buffer;
      fileName: string;
      contentType: string;
      width?: number;
      height?: number;
    }
  | {
      type: "file";
      data: Buffer;
      fileName: string;
      /** Per-file sharing scope override. When omitted, the scope from sendMessageWithFiles is used. */
      sharingScope?: FileSharingScope | { scope: "users"; emails: string[] };
    };

/** Result of sending a message. */
export interface SentMessage {
  /** Server-assigned message ID (OriginalArrivalTime). Use this for edit/delete operations. */
  messageId: string;
  /** Server-reported arrival timestamp (epoch milliseconds). Same value as messageId in numeric form. */
  arrivalTime: number;
}

/** Result of editing a message. */
export interface EditedMessage {
  /** The ID of the edited message. */
  messageId: string;
  /** ISO timestamp when the edit was applied. */
  editTime: string;
}

/** Result of deleting a message. */
export interface DeletedMessage {
  /** The ID of the deleted message. */
  messageId: string;
}

/** Result of adding or removing a reaction on a message. */
export interface ReactionResult {
  /** The ID of the message that was reacted to. */
  messageId: string;
  /** The reaction key (e.g. "like", "heart", "laugh", "surprised"). */
  reactionKey: string;
}

/** Result of scheduling a message (creating a scheduled draft). */
export interface ScheduledMessage {
  /** Server-assigned message ID (OriginalArrivalTime). */
  messageId: string;
  /** Server-reported arrival timestamp (epoch milliseconds). */
  arrivalTime: number;
  /** ISO timestamp when the message is scheduled to be sent. */
  scheduledTime: string;
}

/** A page of messages with pagination metadata. */
export interface MessagesPage {
  /** Messages in this page. */
  messages: Message[];
  /** URL for fetching the previous (older) page, or null if at the beginning. */
  backwardLink: string | null;
  /** Sync state token for incremental updates, or null. */
  syncState: string | null;
}

/** Options for fetching messages from a conversation. */
export interface GetMessagesOptions {
  /** Maximum total number of messages to return. When set, pagination stops
   *  as soon as at least this many messages have been collected, and the
   *  result is trimmed to exactly `limit` items. */
  limit?: number;
  /** Maximum number of pagination pages to fetch (default: 100). */
  maxPages?: number;
  /** Number of messages per page (default: 200). */
  pageSize?: number;
  /** Callback invoked with the running total after each page is fetched. */
  onProgress?: (totalFetched: number) => void;
}

/** Options for listing conversations. */
export interface ListConversationsOptions {
  /** Maximum number of conversations to return (default: 50). */
  pageSize?: number;
  /** If true, exclude system streams (annotations, threads, notifications, etc). */
  excludeSystemStreams?: boolean;
}

/**
 * Result of searching for a 1:1 conversation.
 */
export interface OneOnOneSearchResult {
  /** The conversation thread ID. */
  conversationId: string;
  /** Display name of the matched member. */
  memberDisplayName: string;
}

/** A single entry in a parsed meeting transcript. */
export interface TranscriptEntry {
  /** Speaker name as identified by Teams. */
  speaker: string;
  /** Start time in the recording (ISO 8601 duration or HH:MM:SS.mmm). */
  startTime: string;
  /** End time in the recording. */
  endTime: string;
  /** Spoken text content. */
  text: string;
}

/** Result of fetching a meeting transcript. */
export interface TranscriptResult {
  /** Meeting title extracted from the recording message. */
  meetingTitle: string;
  /** The raw VTT content. */
  rawVtt: string;
  /** Parsed transcript entries with speaker, timestamps, and text. */
  entries: TranscriptEntry[];
}

/** A person found via the Substrate people search API. */
export interface PersonSearchResult {
  /** Display name. */
  displayName: string;
  /** MRI (e.g. "8:orgid:uuid"). */
  mri: string;
  /** Email address. */
  email: string;
  /** Job title, or empty string. */
  jobTitle: string;
  /** Department, or empty string. */
  department: string;
  /** AAD object ID (without tenant). */
  objectId: string;
}

/** A chat found via the Substrate chat search API. */
export interface ChatSearchResult {
  /** Chat display name or topic. */
  name: string;
  /** Thread ID (conversation ID). */
  threadId: string;
  /** Thread type (e.g. "Chat", "Meeting"). */
  threadType: string;
  /** Members whose names matched the query. */
  matchingMembers: Array<{ displayName: string; mri: string }>;
  /** Other members in the chat. */
  chatMembers: Array<{ displayName: string; mri: string }>;
  /** Total number of members. */
  totalMemberCount: number;
}

/** Thread types that represent system streams, not user conversations. */
export const SYSTEM_STREAM_TYPES = [
  "streamofannotations",
  "streamofthreads",
  "streamofnotifications",
  "streamofmentions",
  "streamofnotes",
] as const;
