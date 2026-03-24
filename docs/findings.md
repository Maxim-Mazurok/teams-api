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

Returns `{ OriginalArrivalTime: <timestamp>, id: <clientMessageId> }`.

**Important:** The `id` field in the response is the echoed-back `clientmessageid`, NOT the server-assigned message ID. The server-assigned ID is `OriginalArrivalTime` — this is the value that must be used for edit and delete operations. Using the `clientmessageid` for edit/delete causes a `ColdStoreNotSupportedForMessageException` error.

The `imdisplayname` field is required — it determines how the sender name appears in the chat. If omitted, the message sender shows as blank.

**Edit message:**

```
PUT /users/ME/conversations/{conversationId}/messages/{messageId}
Content-Type: application/json

{
  "content": "<p>updated text</p>",
  "messagetype": "RichText/Html",
  "contenttype": "text",
  "skypeeditedid": "<messageId>",
  "imdisplayname": "Sender Display Name"
}
```

Returns `{ edittime: "<ISO timestamp>" }`.

The `skypeeditedid` field must match the message ID in the URL. Only the message author can edit a message.

**Delete message:**

```
DELETE /users/ME/conversations/{conversationId}/messages/{messageId}
```

Returns empty body on success. Only the message author can delete a message. The message is soft-deleted (marked with `deletetime` in the message metadata) rather than physically removed.

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

### Follow subscriptions

The `properties.emotions` array may include an entry with `key: "follow"`. This represents users who subscribed to ("followed") a channel thread — it is **not** a visible reaction.

- `value: "0"` — the user is actively following
- `value: "1"` — the user unfollowed

Regular reactions store a message timestamp as their `value`; follow entries use `"0"` / `"1"` as boolean flags.

The API client separates these into a dedicated `followers` field on `Message` and excludes them from `reactions`.

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

## Meeting transcript retrieval

Meeting recordings and transcripts in Teams are stored on the recording initiator's OneDrive for Business (SharePoint). The transcript data can be accessed via two paths:

### Path 1: AMS (Async Media Service) — uses existing skype token

The simplest approach. The AMS transcript URL is embedded in the chat messages and can be fetched directly with the skype token.

**Auth**: `Authorization: skype_token <skypeToken>` (note: `skype_token` with underscore, not the Chat Service format)

**Endpoint**: `https://as-prod.asyncgw.teams.microsoft.com/v1/objects/<amsDocumentId>/views/transcript`

**Response**: VTT format (`text/vtt`), with speaker names in `<v>` tags and timestamps. HTML entities are used (e.g. `&#39;` for apostrophe).

### Path 2: SharePoint API — needs separate SharePoint token

Used by the Teams web client for the Recap/Transcript UI. Requires a token for the `{tenant}-my.sharepoint.com` audience, which is acquired via MSAL in the browser.

**Step 1**: Get transcript metadata from drive item:

```
GET https://<tenant>-my.sharepoint.com/_api/v2.1/drives/<driveId>/items/<itemId>?select=media/transcripts&$expand=media/transcripts
Authorization: Bearer <sharepointToken>
```

Response includes `media.transcripts[]` with `id`, `displayName`, `temporaryDownloadUrl`, `languageTag`, etc.

**Step 2**: Download transcript content:

```
GET https://<tenant>-my.sharepoint.com/_api/v2.1/drives/<driveId>/items/<itemId>/media/transcripts/<transcriptId>/streamContent?is=1&applymediaedits=false
Authorization: Bearer <sharepointToken>
```

Response: VTT format.

### How to find transcript URLs from chat messages

Transcript metadata is embedded in two message types in the Chat Service API:

#### `RichText/Media_CallTranscript` message

Content is JSON with fields:

- `scopeId` / `callId` — the call identifier
- `storageId` — `<userId>@<tenantId>` identifying the OneDrive storage
- `isExportedToOdsp` — whether the transcript has been exported to SharePoint

#### `RichText/Media_CallRecording` message (status="Success")

Content is XML (`<URIObject>`) containing `<RecordingContent>` with `<item>` elements:

- `type="amsTranscript"` → AMS URL: `https://as-prod.asyncgw.teams.microsoft.com/v1/objects/<id>/views/transcript`
- `type="onedriveForBusinessTranscript"` → SharePoint URL with `driveId`, `driveItemId`, and transcript `id`
- `type="onedriveForBusinessVideo"` → SharePoint sharing URL with `driveId` and `driveItemId`

The `properties.atp` field contains the SharePoint sharing URL with encoded access tokens.

### Recommended implementation approach

Use **Path 1 (AMS)** since it works with the existing skype token — no new token acquisition needed.

1. Fetch chat messages for the conversation
2. Find `RichText/Media_CallRecording` messages with `RecordingStatus status="Success"`
3. Parse the XML content to extract `<item type="amsTranscript" uri="...">` URL
4. Fetch the transcript VTT from the AMS URL with `Authorization: skype_token <skypeToken>`
5. Parse VTT to extract speaker names and text

### Additional API surfaces observed

| Host                                  | Auth                                 | Purpose                                              |
| ------------------------------------- | ------------------------------------ | ---------------------------------------------------- |
| `as-prod.asyncgw.teams.microsoft.com` | `Authorization: skype_token <token>` | AMS: transcript VTT, video, roster events            |
| `substrate.office.com`                | Bearer token (substrate audience)    | WorkingSetFiles API, search, signals                 |
| `<tenant>-my.sharepoint.com`          | Bearer token (SharePoint audience)   | Drive items, transcript metadata, stream content     |
| `graph.microsoft.com`                 | Bearer token (Graph audience)        | Drive items, shares resolution, user license details |
| `australiaeast1-mediap.svc.ms`        | URL-embedded auth params             | Video manifest/streaming                             |

## SharePoint file upload for Teams messages

Files shared in Teams conversations are uploaded to the sender's OneDrive for Business and referenced in the message via a `properties.files` JSON string.

### Upload flow

1. **PUT file content** to SharePoint:
   ```
   PUT https://{tenant}-my.sharepoint.com/personal/{user_email_underscored}/_api/v2.0/drive/root:/Microsoft%20Teams%20Chat%20Files/{fileName}:/content?@name.conflictBehavior=rename&$select=*,sharepointIds,webDavUrl
   Authorization: Bearer {sharePointToken}
   Content-Type: application/octet-stream

   <file bytes>
   ```

   - `{tenant}-my.sharepoint.com` — personal OneDrive site host (extracted from the SharePoint JWT `aud` claim)
   - `{user_email_underscored}` — user email with `.` and `@` replaced by `_` (e.g. `alice_smith_contoso_com`)
   - `@name.conflictBehavior=rename` — auto-renames on filename collision (appends `(1)`, etc.)
   - Returns **201** with item metadata including `sharepointIds.listItemUniqueId`, `sharepointIds.siteId`, `webDavUrl`, `webUrl`

2. **Send message** with `properties.files` JSON:
   ```
   POST /users/ME/conversations/{conversationId}/messages
   ```
   The message body includes `properties.files` as a JSON string containing an array of file descriptors.

### `properties.files` schema

Each entry in the array:

```json
{
  "@type": "http://schema.skype.com/File",
  "version": 2,
  "id": "{listItemUniqueId}",
  "itemid": "{listItemUniqueId}",
  "fileName": "report.pdf",
  "fileType": "pdf",
  "title": "report.pdf",
  "type": "pdf",
  "state": "active",
  "objectUrl": "https://{host}/personal/{user}/Documents/Microsoft%20Teams%20Chat%20Files/report.pdf",
  "baseUrl": "https://{host}/personal/{user}/",
  "permissionScope": "users",
  "sharepointIds": {
    "listItemUniqueId": "{listItemUniqueId}",
    "siteId": "{siteId}"
  },
  "fileInfo": {
    "itemId": null,
    "fileUrl": "https://{host}/personal/{user}/Documents/Microsoft%20Teams%20Chat%20Files/report.pdf",
    "siteUrl": "https://{host}/personal/{user}/",
    "serverRelativeUrl": "/personal/{user}/Documents/Microsoft Teams Chat Files/report.pdf",
    "shareUrl": null,
    "shareId": null
  },
  "fileChicletState": {
    "serviceName": "p2p",
    "state": "active"
  }
}
```

### Authentication

Uses the SharePoint Bearer token (audience: `*.sharepoint.com`), captured from the MSAL localStorage cache during authentication. This is the same token used for downloading file attachments.
