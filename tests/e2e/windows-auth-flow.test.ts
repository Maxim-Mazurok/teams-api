/**
 * Windows E2E tests — interactive login, token caching, and API operations.
 *
 * Tests the full interactive login flow on Windows: browser launch with
 * persistent profile, token caching, cache reuse, and core API operations.
 *
 * Skipped by default. Requires a browser and user interaction on first run.
 * Run on Windows with:
 *   TEAMS_EMAIL=you@company.com TEAMS_TEST_WINDOWS=1 npx vitest run tests/e2e/windows-auth-flow.test.ts
 */

import { describe, it, expect } from "vitest";
import { TeamsClient } from "../../src/teams-client.js";
import { saveToken, loadToken, clearToken } from "../../src/token-store.js";
import { teamsRegions } from "../../src/region.js";

const email = process.env["TEAMS_EMAIL"];
const shouldRun =
  process.platform === "win32" &&
  Boolean(process.env["TEAMS_TEST_WINDOWS"]) &&
  Boolean(email);

describe.skipIf(!shouldRun)(
  "Windows auth flow",
  { timeout: 5 * 60_000 },
  () => {
    let client: TeamsClient;

    it("should connect via interactive login and cache the token", async () => {
      // Clear any existing cached token so we test the full flow
      await clearToken(email!);
      await clearToken("_default");

      client = await TeamsClient.connect({ email: email!, verbose: true });
      const token = client.getToken();

      expect(token.skypeToken.length).toBeGreaterThan(100);
      expect(teamsRegions).toContain(token.region);
      expect(token.bearerToken).toBeTruthy();

      // Verify token was cached
      const cachedDefault = await loadToken("_default");
      expect(cachedDefault).not.toBeNull();
      const cachedEmail = await loadToken(email!);
      expect(cachedEmail).not.toBeNull();
    });

    it("should reconnect from cache without opening a browser", async () => {
      const client2 = await TeamsClient.connect({
        email: email!,
        verbose: true,
      });
      const t2 = client2.getToken();

      expect(t2.skypeToken).toBe(client.getToken().skypeToken);
      expect(t2.region).toBe(client.getToken().region);
    });

    it("should cache token from fromInteractiveLogin", async () => {
      await clearToken(email!);
      await clearToken("_default");

      const client3 = await TeamsClient.fromInteractiveLogin({
        email: email!,
        verbose: true,
      });
      const t3 = client3.getToken();
      expect(t3.skypeToken.length).toBeGreaterThan(100);

      const cachedDefault = await loadToken("_default");
      expect(cachedDefault).not.toBeNull();
      const cachedEmail = await loadToken(email!);
      expect(cachedEmail).not.toBeNull();
    });
  },
);

describe.skipIf(!shouldRun)(
  "Windows API operations",
  { timeout: 2 * 60_000 },
  () => {
    let client: TeamsClient;

    it("should connect (uses cache from previous suite)", async () => {
      client = await TeamsClient.connect({ email: email!, verbose: true });
      expect(client.getToken().skypeToken.length).toBeGreaterThan(100);
    });

    it("should list conversations", async () => {
      const conversations = await client.listConversations();
      expect(conversations.length).toBeGreaterThan(0);
    });

    it("should find a conversation by name", async () => {
      const conversations = await client.listConversations({ pageSize: 5 });
      // Use the first conversation's topic as a known-good search term
      if (conversations[0]?.topic) {
        const found = await client.findConversation(conversations[0].topic);
        expect(found).not.toBeNull();
      }
    });

    it("should get messages from a conversation", async () => {
      const conversations = await client.listConversations({ pageSize: 5 });
      const messages = await client.getMessages(conversations[0].id, {
        maxPages: 1,
        pageSize: 5,
      });
      expect(messages.length).toBeGreaterThan(0);
    });

    it("should get members of a conversation", async () => {
      const conversations = await client.listConversations({ pageSize: 5 });
      const members = await client.getMembers(conversations[0].id);
      expect(members.length).toBeGreaterThan(0);
    });

    it("should resolve current user display name", async () => {
      const name = await client.getCurrentUserDisplayName();
      expect(name.length).toBeGreaterThan(0);
    });
  },
);
