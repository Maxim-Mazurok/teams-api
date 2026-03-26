/**
 * Teams emoji shortcut-to-ID resolution.
 *
 * Fetches the emoji catalog dynamically from the Teams CDN at runtime
 * and builds a shortcut→ID map in memory. This avoids hardcoding 1100+
 * entries and keeps the map up-to-date as Teams adds new emojis.
 *
 * The CDN URL pattern is:
 *   https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v1/metadata/{version}/{locale}.json
 *
 * The {version} is a content hash embedded in the Teams client config.
 * Multiple known versions are tried in order (newest first).
 * If all fetches fail, resolveReactionKey falls back to returning the
 * lowercased input (standard reactions like "like", "heart", "laugh"
 * have id === shortcut and work without any map).
 */

const EMOJI_CDN_BASE =
  "https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v1/metadata/";

/**
 * Teams ECS (Experimentation and Configuration Service) endpoint.
 * Returns a large JSON config (~1 MB) that includes `emoticonAssetVersion`,
 * the current emoji catalog version hash. No authentication required.
 */
const TEAMS_ECS_CONFIG_URL =
  "https://config.teams.microsoft.com/config/v1/MicrosoftTeams/1415_1.0.0.0?environment=prod&audienceGroup=general&teamsRing=general";

/**
 * Fallback version hashes if the ECS endpoint is unreachable.
 * These are known-good versions, newest first.
 */
const FALLBACK_VERSIONS = [
  "ec4576179210cde40ce5494513213583",
  "0f52465a47bf42f299c74a639443f33e",
];

const EMOJI_LOCALE = "en-gb";

interface EmojiEntry {
  id: string;
  shortcuts: string[];
}

interface EmojiCategory {
  emoticons: EmojiEntry[];
}

interface EmojiCatalog {
  categories: EmojiCategory[];
}

/** Module-level cache: populated by initializeEmojiMap(). */
let shortcutToEmojiId: Record<string, string> | null = null;
let initializationPromise: Promise<void> | null = null;

/**
 * Build a shortcut→ID map from the CDN emoji catalog JSON.
 *
 * For each emoji, maps each shortcut (stripped of parentheses) to the
 * emoji ID. Skips entries where the shortcut already equals the ID
 * (e.g. "like" → "like") since those need no mapping.
 */
function buildShortcutMap(catalog: EmojiCatalog): Record<string, string> {
  const map: Record<string, string> = {};
  for (const category of catalog.categories) {
    for (const emoji of category.emoticons) {
      for (const shortcut of emoji.shortcuts) {
        // Shortcuts are wrapped in parentheses like "(horse)" — strip them
        const cleaned =
          shortcut.startsWith("(") && shortcut.endsWith(")")
            ? shortcut.slice(1, -1)
            : shortcut;
        // Only add entries where shortcut differs from ID
        if (cleaned !== emoji.id) {
          map[cleaned] = emoji.id;
        }
      }
    }
  }
  return map;
}

/**
 * Fetch the current emoji catalog version hash from the Teams ECS config.
 *
 * Returns null if the ECS endpoint is unreachable or doesn't contain the version.
 */
async function fetchCurrentVersion(): Promise<string | null> {
  try {
    const response = await fetch(TEAMS_ECS_CONFIG_URL);
    if (!response.ok) {
      console.warn(
        `[emoji-map] ECS config returned ${String(response.status)}, will use fallback versions`,
      );
      return null;
    }
    const text = await response.text();
    const match = /"emoticonAssetVersion":"([^"]+)"/.exec(text);
    if (!match) {
      console.warn(
        "[emoji-map] emoticonAssetVersion not found in ECS config, will use fallback versions",
      );
      return null;
    }
    return match[1];
  } catch (error) {
    console.warn(
      "[emoji-map] Failed to fetch ECS config:",
      error instanceof Error ? error.message : error,
    );
    return null;
  }
}

/**
 * Fetch the emoji catalog from the Teams CDN.
 *
 * Tries the current version from ECS first, then falls back to known
 * version hashes. Returns null if all attempts fail.
 */
async function fetchEmojiCatalog(): Promise<EmojiCatalog | null> {
  const currentVersion = await fetchCurrentVersion();
  const versionsToTry = currentVersion
    ? [currentVersion, ...FALLBACK_VERSIONS.filter((v) => v !== currentVersion)]
    : FALLBACK_VERSIONS;

  for (const version of versionsToTry) {
    const url = `${EMOJI_CDN_BASE}${version}/${EMOJI_LOCALE}.json`;
    try {
      const response = await fetch(url);
      if (response.ok) {
        return (await response.json()) as EmojiCatalog;
      }
      console.warn(
        `[emoji-map] CDN returned ${String(response.status)} for version ${version}`,
      );
    } catch (error) {
      console.warn(
        `[emoji-map] Network error fetching version ${version}:`,
        error instanceof Error ? error.message : error,
      );
    }
  }
  console.warn(
    "[emoji-map] Failed to fetch emoji catalog from all versions. " +
      "Non-standard emoji shortcuts (e.g. 'horse') will not resolve to Teams emoji IDs.",
  );
  return null;
}

/**
 * Reset internal state — exposed for test isolation only.
 */
export function resetEmojiMap(): void {
  shortcutToEmojiId = null;
  initializationPromise = null;
}

/**
 * Initialize the emoji shortcut→ID map by fetching from the Teams CDN.
 *
 * Safe to call multiple times — only the first call triggers a fetch.
 * If the fetch fails, resolveReactionKey will fall back to returning
 * the lowercased input (which works for standard reactions).
 */
export async function initializeEmojiMap(): Promise<void> {
  if (shortcutToEmojiId !== null) return;

  if (initializationPromise === null) {
    initializationPromise = fetchEmojiCatalog().then((catalog) => {
      if (catalog) {
        shortcutToEmojiId = buildShortcutMap(catalog);
      }
    });
  }

  return initializationPromise;
}

/**
 * Resolve a user-friendly reaction name to the Teams emoji ID.
 *
 * Accepts shortcuts (e.g. "horse"), emoji IDs (e.g. "1f40e_horse"),
 * or standard reaction names (e.g. "like"). Returns the emoji ID that
 * the Teams Chat Service API expects.
 *
 * Resolution order:
 * 1. If the emoji map has been fetched, looks up the input as a shortcut.
 * 2. Otherwise, returns the lowercased input (works for standard reactions
 *    where id === shortcut, e.g. "like", "heart", "laugh").
 */
export function resolveReactionKey(input: string): string {
  if (shortcutToEmojiId !== null) {
    // Try exact match first (handles emoticon shortcuts like ":D")
    if (input in shortcutToEmojiId) {
      return shortcutToEmojiId[input];
    }
    // Try case-insensitive match (handles "Horse" → "horse" → "1f40e_horse")
    const lowered = input.toLowerCase();
    return shortcutToEmojiId[lowered] ?? lowered;
  }
  // Map not loaded — fall back to lowercasing (standard reactions still work)
  return input.toLowerCase();
}

