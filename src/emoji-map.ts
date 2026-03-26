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
 * Teams ECS (Experimentation and Configuration Service) endpoint base URL.
 * We append an app-version segment and a fixed query string.
 */
const TEAMS_ECS_CONFIG_BASE_URL =
  "https://config.teams.microsoft.com/config/v1/MicrosoftTeams";

/**
 * Query params required by the public ECS endpoint.
 */
const TEAMS_ECS_CONFIG_QUERY =
  "environment=prod&audienceGroup=general&teamsRing=general";

/**
 * App-version candidates for ECS config lookup.
 *
 * `0_0.0.0.0` is treated as a version-agnostic request and currently returns
 * the same `emoticonAssetVersion` as concrete app versions. Keep a concrete
 * fallback in case ECS behavior changes.
 */
const TEAMS_ECS_APP_VERSION_CANDIDATES = ["0_0.0.0.0", "1415_1.0.0.0"];

/**
 * Fallback version hashes if the ECS endpoint is unreachable.
 * These are known-good versions, newest first.
 */
const FALLBACK_VERSIONS = [
  "ec4576179210cde40ce5494513213583",
  "0f52465a47bf42f299c74a639443f33e",
];

const EMOJI_LOCALE = "en-gb";

function buildTeamsEcsConfigUrl(appVersion: string): string {
  return `${TEAMS_ECS_CONFIG_BASE_URL}/${appVersion}?${TEAMS_ECS_CONFIG_QUERY}`;
}

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
  for (const appVersionCandidate of TEAMS_ECS_APP_VERSION_CANDIDATES) {
    const teamsEcsConfigUrl = buildTeamsEcsConfigUrl(appVersionCandidate);
    try {
      const response = await fetch(teamsEcsConfigUrl);
      if (!response.ok) {
        console.warn(
          `[emoji-map] ECS config returned ${String(response.status)} for app version ${appVersionCandidate}, trying next candidate`,
        );
        continue;
      }
      const responseText = await response.text();
      const emoticonAssetVersionMatch =
        /"emoticonAssetVersion":"([^"]+)"/.exec(responseText);
      if (!emoticonAssetVersionMatch) {
        console.warn(
          `[emoji-map] emoticonAssetVersion not found for app version ${appVersionCandidate}, trying next candidate`,
        );
        continue;
      }
      const resolvedVersion = emoticonAssetVersionMatch[1];
      // Validate: must be a 32-char lowercase hex hash to avoid malformed URLs
      if (!/^[0-9a-f]{32}$/.test(resolvedVersion)) {
        console.warn(
          `[emoji-map] Unexpected emoticonAssetVersion format "${resolvedVersion}" for app version ${appVersionCandidate}, trying next candidate`,
        );
        continue;
      }
      return resolvedVersion;
    } catch (error) {
      console.warn(
        `[emoji-map] Failed to fetch ECS config for app version ${appVersionCandidate}:`,
        error instanceof Error ? error.message : error,
      );
    }
  }
  return null;
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
    ? [
        currentVersion,
        ...FALLBACK_VERSIONS.filter(
          (fallbackVersion) => fallbackVersion !== currentVersion,
        ),
      ]
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

