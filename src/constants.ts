/**
 * Shared constants used across the codebase.
 */

/** Message types that represent user-authored text content. */
export const TEXT_MESSAGE_TYPES = ["RichText/Html", "Text"] as const;

export type TextMessageType = (typeof TEXT_MESSAGE_TYPES)[number];

/** Check whether a message type represents user-authored text content. */
export function isTextMessageType(
  messageType: string,
): messageType is TextMessageType {
  return (TEXT_MESSAGE_TYPES as readonly string[]).includes(messageType);
}
