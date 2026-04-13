/**
 * Conversation protection — restricts edit/delete in sensitive conversations.
 *
 * Reads `TEAMS_PROTECTED_CONVERSATIONS` (comma-separated glob patterns).
 * When a conversation label matches any pattern, edit and delete actions
 * are blocked with an informative error.
 *
 * Glob wildcards: `*` matches any characters. Matching is case-insensitive.
 */

/**
 * Parse the `TEAMS_PROTECTED_CONVERSATIONS` env var (or a raw string)
 * into a trimmed array of non-empty patterns.
 */
export function resolveProtectedPatterns(
  raw?: string | null,
): readonly string[] {
  const value = raw ?? process.env.TEAMS_PROTECTED_CONVERSATIONS ?? "";
  return value
    .split(",")
    .map((pattern) => pattern.trim())
    .filter(Boolean);
}

/**
 * Convert a glob pattern (supporting `*`) into a case-insensitive RegExp.
 *
 * Every character except `*` is escaped; `*` becomes `.*`.
 */
function globToRegExp(pattern: string): RegExp {
  const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, (character) =>
    character === "*" ? ".*" : `\\${character}`,
  );
  return new RegExp(`^${escaped}$`, "i");
}

/**
 * Check whether a conversation label matches any of the protected patterns.
 *
 * Returns the first matching pattern, or `undefined` if no match.
 */
export function matchProtectedConversation(
  conversationLabel: string,
  patterns: readonly string[],
): string | undefined {
  for (const pattern of patterns) {
    if (globToRegExp(pattern).test(conversationLabel)) {
      return pattern;
    }
  }
  return undefined;
}
