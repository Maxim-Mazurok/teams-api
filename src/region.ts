const KNOWN_TEAMS_REGIONS = ["apac", "emea", "amer"] as const;
const CHAT_SERVICE_HOST_PATTERN =
  /^(?<region>[a-z]+)\.ng\.msg\.teams\.microsoft\.com$/i;
const MIDDLE_TIER_PATH_PATTERN = /^\/api\/mt\/(?<region>[a-z]+)(?:\/|$)/i;

export const DEFAULT_TEAMS_REGION = "apac";
export const teamsRegions = [...KNOWN_TEAMS_REGIONS];

export type KnownTeamsRegion = (typeof KNOWN_TEAMS_REGIONS)[number];

export function canonicalizeTeamsRegion(
  region: string | undefined | null,
): string | undefined {
  const normalized = region?.trim().toLowerCase();
  return normalized ? normalized : undefined;
}

export function isKnownTeamsRegion(
  region: string | undefined | null,
): region is KnownTeamsRegion {
  const normalized = canonicalizeTeamsRegion(region);
  return normalized !== undefined
    ? KNOWN_TEAMS_REGIONS.includes(normalized as KnownTeamsRegion)
    : false;
}

export function detectTeamsRegionFromUrl(
  requestUrl: string,
): KnownTeamsRegion | undefined {
  try {
    const url = new URL(requestUrl);

    const hostRegion = canonicalizeTeamsRegion(
      CHAT_SERVICE_HOST_PATTERN.exec(url.hostname)?.groups?.region,
    );
    if (isKnownTeamsRegion(hostRegion)) {
      return hostRegion;
    }

    const pathRegion = canonicalizeTeamsRegion(
      MIDDLE_TIER_PATH_PATTERN.exec(url.pathname)?.groups?.region,
    );
    if (isKnownTeamsRegion(pathRegion)) {
      return pathRegion;
    }
  } catch {
    // Ignore malformed URLs and report no detection.
  }

  return undefined;
}

export function resolveTeamsRegion(
  explicitRegion?: string,
  detectedRegion?: string,
  fallbackRegion = DEFAULT_TEAMS_REGION,
): string {
  return (
    canonicalizeTeamsRegion(explicitRegion) ??
    canonicalizeTeamsRegion(detectedRegion) ??
    canonicalizeTeamsRegion(fallbackRegion) ??
    DEFAULT_TEAMS_REGION
  );
}
