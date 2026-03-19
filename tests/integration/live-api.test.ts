/**
 * Integration tests — hit the real Teams API.
 *
 * These tests are skipped by default.
 * Run them with: TEAMS_TOKEN=<token> TEAMS_REGION=<apac|emea|amer> npx vitest run tests/integration
 *
 * Prerequisites:
 *   - A valid skype token (acquire via `npx tsx src/cli.ts auth --auto --email you@company.com`)
 *   - A Teams account with at least one conversation
 */

import { describe, it, expect } from "vitest";
import { TeamsClient } from "../../src/teams-client.js";

const skypeToken = process.env["TEAMS_TOKEN"];
const region = process.env["TEAMS_REGION"];

const shouldRun = Boolean(skypeToken && region);

describe.skipIf(!shouldRun)("Live Teams API", () => {
  function createClient(): TeamsClient {
    if (!skypeToken) throw new Error("TEAMS_TOKEN is required");
    if (!region) throw new Error("TEAMS_REGION is required");
    return TeamsClient.fromToken(skypeToken, region);
  }

  it("should list conversations", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 10 });

    expect(conversations.length).toBeGreaterThan(0);
    for (const conversation of conversations) {
      expect(typeof conversation.id).toBe("string");
      expect(typeof conversation.threadType).toBe("string");
    }
  });

  it("should get messages from the first conversation", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 5 });
    expect(conversations.length).toBeGreaterThan(0);

    const messages = await client.getMessages(conversations[0].id, {
      maxPages: 1,
      pageSize: 10,
    });

    expect(messages.length).toBeGreaterThan(0);
    for (const message of messages) {
      expect(typeof message.id).toBe("string");
      expect(typeof message.originalArrivalTime).toBe("string");
    }
  });

  it("should resolve current user display name", async () => {
    const client = createClient();
    const displayName = await client.getCurrentUserDisplayName();

    expect(typeof displayName).toBe("string");
    expect(displayName).not.toBe("Unknown User");
  });

  it("should get members of a conversation", async () => {
    const client = createClient();
    const conversations = await client.listConversations({ pageSize: 5 });
    expect(conversations.length).toBeGreaterThan(0);

    const members = await client.getMembers(conversations[0].id);

    expect(members.length).toBeGreaterThan(0);
    for (const member of members) {
      expect(typeof member.id).toBe("string");
    }
  });
});
