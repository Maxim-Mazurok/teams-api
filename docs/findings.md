# API findings

Technical notes discovered while building and testing the Teams Chat Service REST API integration.

## Teams Chat Service REST API

The API base URL follows the pattern `https://{region}.ng.msg.teams.microsoft.com/v1`.

Known regions: `apac`, `emea`, `amer`.

### Authentication

All endpoints require the header:

```
Authentication: skypetoken=<token>
```

Note: it's `Authentication`, not `Authorization`. The skype token value is embedded directly in the header.

Token lifetime is approximately 24 hours due to the `skypetoken_asm` cookie setting `Max-Age=86399`.

### Token acquisition

The skype token is acquired from the Teams web application at `https://teams.microsoft.com`. There are two proven methods:

1. **Playwright auto-login** — Launch system Chrome via Playwright's persistent context, navigate to Teams, and intercept the `skypetoken_asm` cookie after the authentication flow completes. On macOS with a FIDO2 platform authenticator (e.g. Intune Company Portal) and passkey enrolled, the entire login flow is fully automated with zero user interaction.

2. **CDP debug session** — Connect to a running Chrome instance (started with `--remote-debugging-port=9222`) where Teams is already open. Evaluate `document.cookie` in the Teams page context to extract `skypetoken_asm`.

### Key endpoints

**List conversations:**

```
GET /users/ME/conversations?view=mychats&pageSize={n}
```

Returns `{ conversations: [...] }` with fields: `id`, `version`, `threadProperties.topic`, `threadProperties.memberCount`, `threadProperties.threadType`, `threadProperties.lastMessage.composetime`.

**Get messages (one page):**

```
GET /users/ME/conversations/{conversationId}/messages?pageSize={n}&view=superchat
```

Returns `{ messages: [...], _metadata: { backwardLink, syncState } }`.

Pagination: follow `_metadata.backwardLink` for older messages. When `backwardLink` is `null`, all history has been fetched.

**Send message:**

```
POST /users/ME/conversations/{conversationId}/messages
Content-Type: application/json

{
  "content": "<p>message text</p>",
  "messagetype": "RichText/Html",
  "contenttype": "text",
  "imdisplayname": "Sender Display Name",
  "clientmessageid": "<unique-id>",
  "properties": {
    "importance": "",
    "subject": ""
  }
}
```

Returns `{ OriginalArrivalTime: <timestamp>, id: <messageId> }`.

The `imdisplayname` field is required — it determines how the sender name appears in the chat. If omitted, the message sender shows as blank.

**Get members:**

```
GET /threads/{conversationId}?view=msnp24Equivalent
```

Returns `{ members: [{ id, friendlyName, role }] }`.

Note: For 1:1 chats, `friendlyName` is typically empty. To resolve display names, scan recent message senders.

**Get user properties:**

```
GET /users/ME/properties
```

Returns user metadata including `displayname`.

### Message structure

Each raw message includes:

| Field                 | Description                                               |
| --------------------- | --------------------------------------------------------- |
| `id`                  | Unique message ID                                         |
| `messagetype`         | `Text`, `RichText/Html`, `ThreadActivity/AddMember`, etc. |
| `from`                | Sender MRI, e.g. `https://...;messagingv2/8:orgid:uuid`   |
| `imdisplayname`       | Sender display name                                       |
| `content`             | Message body (plain text or HTML)                         |
| `originalarrivaltime` | ISO timestamp                                             |
| `composetime`         | ISO timestamp                                             |
| `edittime`            | ISO timestamp or empty                                    |
| `properties.emotions` | Reactions — can be a JSON string or an array              |
| `properties.mentions` | Mentions in `<at>` tags                                   |
| `amsreferences`       | Inline image/file references                              |

### Reactions format

The `properties.emotions` field is inconsistent:

- Sometimes a **JSON string**: `"[{\"key\":\"like\",\"users\":[{\"mri\":\"8:orgid:uuid\",\"time\":1234}]}]"`
- Sometimes a **raw array**: `[{"key":"like","users":[{"mri":"8:orgid:uuid","time":1234}]}]`

Both formats must be handled. Each emotion has a `key` (reaction name) and `users` array with MRI and timestamp.

### Mentions format

The `properties.mentions` field is an array of objects:

```json
[
  {
    "id": 0,
    "mentionType": "person",
    "mri": "8:orgid:uuid",
    "displayName": "Alice Smith"
  }
]
```

### Conversation types

| Thread type             | Description                           |
| ----------------------- | ------------------------------------- |
| `chat`                  | 1:1 or group chat                     |
| `space`                 | Channel/space                         |
| `meeting`               | Meeting chat                          |
| `streamofnotes`         | Self-chat (ID starts with `48:notes`) |
| `streamofannotations`   | System stream                         |
| `streamofnotifications` | System stream                         |
| `streamofmentions`      | System stream                         |
| `streamofthreads`       | System stream                         |

System streams should be filtered out in most user-facing operations.

### Deleted messages

Deleted messages have `messagetype` set to `Text` but include `properties.deletetime` with an ISO timestamp. The `content` field is typically empty.

### Quoted messages

When a message is a reply, the `content` field contains a `<blockquote>` element wrapping the quoted message. The `itemtype="http://schema.skype.com/Reply"` attribute identifies these. The quoted message ID can be extracted from `data-cid` or `messageid` attributes on inner elements.

## Worker intercept findings (from browser extension research)

These findings are from the browser extension POC and may be useful for future hybrid approaches:

- Workers created by Teams (`/v2/worker/precompiled-web-worker-*.js`) carry the richest message data
- Main-thread `fetch` and `XMLHttpRequest` do not capture chat payloads — they go through the worker
- Worker traffic includes GraphQL operations like `ComponentsChatQueriesMessageListQuery`
- Worker responses include `content`, `quotedMessages`, `fromUser.displayName`, `emotionsSummary`, and `emotions[].users[].userId`
- When Teams serves from cache, the worker still returns structured data (from `indexedDB_NewGetRangeMethod`)
- In a probe run: `worker-create: 3`, `worker-request: 57`, `worker-response: 47`, with `fetch: 0`, `xhr: 0`, `ws: 0`
