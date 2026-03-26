/**
 * Server-level instructions for the Teams API MCP server.
 *
 * This is the single source of truth for workflow guidance, authentication
 * notes, and general tips that help AI agents use the tools effectively.
 * It is served via the MCP `instructions` field on server initialization
 * and also printed by the CLI `guide` command.
 *
 * Per-tool documentation (descriptions, parameters) lives in the action
 * definitions — this file only covers cross-cutting concerns.
 */

export const serverInstructions = `
teams-api: AI-native Microsoft Teams integration.

All tools share a unified interface. Every tool accepts an optional "format" parameter ("concise" or "detailed"). Default format is "concise".

- **concise** (default): Light Markdown with actionable IDs and key decision fields. Use this for most operations — it's human-readable and keeps follow-up actions possible (message IDs for reply/react/edit/delete, conversation IDs for subsequent calls). Nested collections may be summarized.
- **detailed**: Full JSON output. Use when you need exact field values, programmatic processing, or want to inspect the complete data structure.

## Typical workflows

### Read and summarize a conversation

Call teams_get_messages with chat: "Chat Name" for group chats, or to: "Person Name" for 1:1 chats. Conversation resolution happens automatically — no need to look up the ID first.

### Reply to someone

Call teams_send_message with to: "Person Name" and content: "Your reply" for 1:1 chats, or chat: "Chat Name" and content: "Your message" for group chats.

### Monitor a conversation for new messages

Call teams_get_messages with chat: "Chat Name" and order: "newest-first" and limit: 10 to check recent activity.

### Find out who is in a chat

Call teams_get_members with chat: "Chat Name" or to: "Person Name".

### Look up a conversation ID

Call teams_find_conversation with the topic name. For 1:1 chats (which have no topic), use teams_find_one_on_one with the person's name.

### Search for people or chats

Call teams_find_people to search the organization directory by name. Call teams_find_chats to search chats by name or member.

### Get a meeting transcript

Call teams_get_transcript with chat: "Meeting Name" to retrieve the parsed transcript, or pass rawVtt: true for the original VTT file.

## Important notes

- Most tools accept chat, to, or conversationId to identify a conversation. You only need one — the server resolves the rest automatically.
- Conversation IDs look like "19:abc123@thread.v2" for group chats or "48:notes" for self-chat.
- 1:1 chats have no topic name — use teams_find_one_on_one or the "to" parameter to find them by person name.
- Messages include reactions, mentions, followers (thread subscribers), and quoted message references.
- Content for teams_send_message and teams_edit_message is interpreted as Markdown by default and converted to rich HTML. Use messageFormat: "html" for raw HTML or "text" for plain text.
- Token lifetime is ~24 hours. The server re-authenticates automatically when using auto-login (TEAMS_AUTO=true).
- All tools share the same parameters and behavior as the CLI commands and programmatic API.
`.trim();
