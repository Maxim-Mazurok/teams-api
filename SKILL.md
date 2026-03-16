# teams-api MCP Skill

Use this skill when you need to interact with Microsoft Teams — reading messages, sending replies, listing conversations, or looking up team members.

## When to use

- Reading unread messages from Teams chats
- Sending messages or replies to conversations
- Finding a specific conversation or 1:1 chat
- Looking up who is in a conversation
- Monitoring conversations for updates
- Acting on behalf of the user in Teams workflows

## Available MCP tools

### teams_list_conversations

List the user's available Teams conversations (chats, group chats, meetings).

Parameters:

- `pageSize` (optional, number): How many conversations to fetch. Default: 50.

### teams_find_conversation

Find a conversation by topic name. Uses case-insensitive partial matching.

Parameters:

- `query` (required, string): Search term to match against conversation topics.

### teams_find_one_on_one

Find a 1:1 chat with a specific person by their display name.

Parameters:

- `personName` (required, string): Full or partial name of the person.

### teams_get_messages

Get messages from a conversation. Automatically paginates for complete history.

Parameters:

- `conversationId` (required, string): The conversation ID (e.g. `19:abc@thread.v2`).
- `maxPages` (optional, number): Maximum pages to fetch. Default: 5.
- `pageSize` (optional, number): Messages per page. Default: 50.

### teams_send_message

Send a plain-text message to a conversation.

Parameters:

- `conversationId` (required, string): Target conversation ID.
- `content` (required, string): Message text to send.

### teams_get_members

List members of a conversation.

Parameters:

- `conversationId` (required, string): The conversation ID.

### teams_whoami

Get the display name of the currently authenticated user.

No parameters.

## Typical workflows

### Read and summarize a conversation

1. Call `teams_find_conversation` with the chat name
2. Call `teams_get_messages` with the conversation ID
3. Summarize the messages

### Reply to someone

1. Call `teams_find_one_on_one` with the person's name (or `teams_find_conversation` for group chats)
2. Call `teams_send_message` with the conversation ID and your reply

### Monitor a conversation

1. Call `teams_find_conversation` to get the conversation ID
2. Periodically call `teams_get_messages` with `maxPages: 1` to check for new messages

### Find out who is in a chat

1. Call `teams_find_conversation` or `teams_list_conversations` to get the conversation ID
2. Call `teams_get_members` with that ID

## Authentication

The MCP server handles authentication automatically via environment variables. The user configures `TEAMS_AUTO=true` and `TEAMS_EMAIL=user@company.com` for unattended login, or `TEAMS_TOKEN=<token>` for a pre-acquired token.

## Important notes

- Conversation IDs look like `19:abc123@thread.v2` for group chats or `48:notes` for self-chat
- 1:1 chats have no topic name — use `teams_find_one_on_one` to find them by person name
- Messages include reactions, mentions, and quoted message references
- Token lifetime is ~24 hours — the server re-authenticates automatically with auto-login
