/**
 * End-to-end tests — full login → read → write flow.
 *
 * These tests are skipped by default. They require full system-level access:
 *   - macOS with a FIDO2 platform authenticator
 *   - System Chrome installed
 *   - FIDO2 passkey enrolled
 *
 * Run with: TEAMS_EMAIL=you@company.com npx vitest run tests/e2e
 */

import { describe, it, expect } from "vitest";
import { TeamsClient } from "../../src/teams-client.js";
import { teamsRegions } from "../../src/region.js";

const email = process.env["TEAMS_EMAIL"];
const shouldRun = Boolean(email);

describe.skipIf(!shouldRun)("Full E2E flow", { timeout: 120_000 }, () => {
  it("should auto-login and list conversations", async () => {
    if (!email) throw new Error("TEAMS_EMAIL is required");

    console.log("[e2e] Starting auto-login...");
    const client = await TeamsClient.fromAutoLogin({ email, verbose: true });

    console.log("[e2e] Listing conversations...");
    const conversations = await client.listConversations({ pageSize: 5 });
    console.log(`[e2e] Got ${conversations.length} conversations`);
    expect(conversations.length).toBeGreaterThan(0);

    const token = client.getToken();
    expect(token.skypeToken.length).toBeGreaterThan(100);
    expect(teamsRegions).toContain(token.region);
    console.log("[e2e] Test 1 passed");
  });

  it("should send a message to self-chat and read it back", async () => {
    if (!email) throw new Error("TEAMS_EMAIL is required");

    console.log("[e2e] Starting auto-login for send/read test...");
    const client = await TeamsClient.fromAutoLogin({ email, verbose: true });
    console.log("[e2e] Getting display name...");
    const displayName = await client.getCurrentUserDisplayName();
    console.log(`[e2e] Display name: ${displayName}`);

    // Find self-chat
    const selfChat = await client.findOneOnOneConversation(displayName);
    expect(selfChat).not.toBeNull();

    // Send a test message with unique content
    const uniqueContent = `E2E test message ${Date.now()}`;
    const sentResult = await client.sendMessage(
      selfChat!.conversationId,
      uniqueContent,
    );
    expect(typeof sentResult.messageId).toBe("string");

    // Read messages and verify our message appears
    const messages = await client.getMessages(selfChat!.conversationId, {
      maxPages: 1,
      pageSize: 20,
    });

    const ourMessage = messages.find((message) =>
      message.content.includes(uniqueContent),
    );
    expect(ourMessage).toBeDefined();
    console.log("[e2e] Test 2 passed");
  });
});
