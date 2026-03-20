---
name: teams-api
description: Interact with Microsoft Teams — read messages, send replies, list conversations, search people, and look up team members. Use when an AI agent needs to participate in Teams workflows.
---

# teams-api MCP Skill

Use this skill when you need to interact with Microsoft Teams — reading messages, sending replies, listing conversations, or looking up team members.

All tools share a unified interface with the CLI and programmatic API. Tool definitions, parameters, and descriptions are generated from a single source (`src/actions.ts`).\n\nAll tools accept an optional `format` parameter (`json`, `text`, `md`, or `toon`). Default format is `toon`.

## When to use

- Reading unread messages from Teams chats
- Sending messages or replies to conversations
- Finding a specific conversation or 1:1 chat
- Looking up who is in a conversation
- Monitoring conversations for updates
- Acting on behalf of the user in Teams workflows

## Available MCP tools

### teams_list_conversations

List conversations (chats, group chats, meetings, channels). Returns conversation ID, topic, type, member count, and last message time.

Parameters:

- `limit` (optional, number): Maximum number of conversations to return. Default: 50.

### teams_find_conversation

Find a conversation by topic name (case-insensitive partial match). For 1:1 chats (which have no topic), use `teams_find_one_on_one` instead.

Parameters:

- `query` (required, string): Partial topic name to search for.

### teams_find_one_on_one

Find a 1:1 conversation with a person by name. Searches untitled chats by scanning recent message sender names. Also finds the self-chat if the name matches the current user.

Parameters:

- `personName` (required, string): Name of the person to find (case-insensitive partial match).

### teams_get_messages

Get messages from a conversation. Identify the conversation by topic name, person name for 1:1 chats, or direct ID. At least one identifier is required.

Parameters:

- `chat` (optional, string): Find conversation by topic name (partial match).
- `to` (optional, string): Find 1:1 conversation by person name.
- `conversationId` (optional, string): Direct conversation thread ID.
- `maxPages` (optional, number): Maximum pagination pages to fetch. Default: 100.
- `pageSize` (optional, number): Messages per page. Default: 200.
- `textOnly` (optional, boolean): Only return text messages, excluding system events. Default: true.
- `order` (optional, string): Message order — `oldest-first` (chronological, default) or `newest-first`.

### teams_send_message

Send a plain-text message to a conversation. Identify the conversation by topic name, person name for 1:1 chats, or direct ID. At least one identifier is required.

Parameters:

- `chat` (optional, string): Find conversation by topic name (partial match).
- `to` (optional, string): Find 1:1 conversation by person name.
- `conversationId` (optional, string): Direct conversation thread ID.
- `content` (required, string): Message text to send.

### teams_get_members

List members of a conversation. Identify the conversation by topic name, person name for 1:1 chats, or direct ID. At least one identifier is required. Note: 1:1 chat members may have empty display names.

Parameters:

- `chat` (optional, string): Find conversation by topic name (partial match).
- `to` (optional, string): Find 1:1 conversation by person name.
- `conversationId` (optional, string): Direct conversation thread ID.

### teams_whoami

Get the display name and region of the currently authenticated user.

No parameters.

## Typical workflows

### Read and summarize a conversation

1. Call `teams_get_messages` with `chat: "Chat Name"` — conversation resolution happens automatically
2. Summarize the messages

Or for 1:1 chats:

1. Call `teams_get_messages` with `to: "Person Name"`
2. Summarize the messages

### Reply to someone

1. Call `teams_send_message` with `to: "Person Name"` and `content: "Your reply"`

Or for group chats:

1. Call `teams_send_message` with `chat: "Chat Name"` and `content: "Your message"`

### Monitor a conversation

1. Call `teams_get_messages` with `chat: "Chat Name"` and `maxPages: 1` to check for new messages

### Find out who is in a chat

1. Call `teams_get_members` with `chat: "Chat Name"`

### Look up a conversation ID

1. Call `teams_find_conversation` with the topic name
2. Use the returned conversation ID for subsequent operations

## Authentication

The MCP server handles authentication automatically via environment variables. Common configurations:

- `TEAMS_TOKEN=<token>` — pre-acquired skype token (all platforms)
- `TEAMS_AUTO=true` and `TEAMS_EMAIL=user@company.com` — unattended FIDO2 login (macOS only)
- `TEAMS_LOGIN=true` — interactive browser login (all platforms, opens a visible browser window)

## Important notes

- Conversation IDs look like `19:abc123@thread.v2` for group chats or `48:notes` for self-chat
- 1:1 chats have no topic name — use `teams_find_one_on_one` or the `to` parameter to find them by person name
- Messages include reactions, mentions, and quoted message references
- Token lifetime is ~24 hours — the server re-authenticates automatically with auto-login
- All tools share the same parameters and behavior as the CLI commands and programmatic API
