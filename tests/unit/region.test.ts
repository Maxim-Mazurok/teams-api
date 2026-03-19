import { describe, expect, it } from "vitest";
import {
  DEFAULT_TEAMS_REGION,
  detectTeamsRegionFromUrl,
  resolveTeamsRegion,
} from "../../src/region.js";

describe("detectTeamsRegionFromUrl", () => {
  it("should detect the region from a chat service hostname", () => {
    expect(
      detectTeamsRegionFromUrl(
        "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations",
      ),
    ).toBe("emea");
  });

  it("should detect the region from a middle-tier path", () => {
    expect(
      detectTeamsRegionFromUrl(
        "https://teams.cloud.microsoft/api/mt/amer/beta/users/fetchShortProfile",
      ),
    ).toBe("amer");
  });

  it("should ignore unknown or malformed URLs", () => {
    expect(
      detectTeamsRegionFromUrl(
        "https://teams.cloud.microsoft/api/mt/unknown/beta/users/fetchShortProfile",
      ),
    ).toBeUndefined();
    expect(detectTeamsRegionFromUrl("not a url")).toBeUndefined();
  });
});

describe("resolveTeamsRegion", () => {
  it("should prefer an explicit region override", () => {
    expect(resolveTeamsRegion("emea", "amer")).toBe("emea");
  });

  it("should use the detected region when no override is provided", () => {
    expect(resolveTeamsRegion(undefined, "amer")).toBe("amer");
  });

  it("should fall back to the default region when detection is unavailable", () => {
    expect(resolveTeamsRegion()).toBe(DEFAULT_TEAMS_REGION);
  });
});
