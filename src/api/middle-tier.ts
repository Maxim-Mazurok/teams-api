/**
 * Teams middle-tier API for profile resolution.
 *
 * HTTP calls to teams.cloud.microsoft/api/mt/{region} for user profiles.
 * Requires an OAuth2 Bearer token (api.spaces.skype.com audience).
 */

import type { TeamsToken, UserProfile } from "../types.js";
import { fetchWithRetry, ApiAuthError } from "./common.js";

const MIDDLE_TIER_BASE = "https://teams.cloud.microsoft/api/mt";

/**
 * Resolve display names for a batch of MRIs via the Teams middle-tier profile endpoint.
 *
 * Requires a Bearer token (api.spaces.skype.com audience). Throws `ApiAuthError`
 * if the token is unavailable so callers can trigger re-authentication.
 */
export async function fetchProfiles(
  token: TeamsToken,
  mris: string[],
): Promise<UserProfile[]> {
  if (mris.length === 0) {
    return [];
  }
  if (!token.bearerToken) {
    throw new ApiAuthError(
      "Bearer token is missing — re-authentication required for profile resolution",
    );
  }

  const url = `${MIDDLE_TIER_BASE}/${token.region}/beta/users/fetchShortProfile?isMailAddress=false&enableGuest=true&skypeTeamsInfo=true`;

  const response = await fetchWithRetry(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token.bearerToken}`,
    },
    body: JSON.stringify(mris),
  });

  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      throw new ApiAuthError(
        `Profile resolution authentication failed: ${response.status} ${response.statusText}`,
      );
    }
    if (response.status >= 500) {
      throw new Error(
        `Profile resolution server error: ${response.status} ${response.statusText}`,
      );
    }
    return [];
  }

  const data = (await response.json()) as {
    value?: Array<{
      mri?: string;
      displayName?: string;
      email?: string;
      jobTitle?: string;
      userType?: string;
    }>;
  };

  return (data.value ?? []).map((profile) => ({
    mri: profile.mri ?? "",
    displayName: profile.displayName ?? "",
    email: profile.email ?? "",
    jobTitle: profile.jobTitle ?? "",
    userType: profile.userType ?? "",
  }));
}
