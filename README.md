# teams-api

AI-native Microsoft Teams integration — read conversations, send messages, and manage members via the Teams Chat Service REST API.

Designed for autonomous AI agents that need to interact with Teams: read messages, reply to people, monitor conversations, and participate in team workflows.

> [!NOTE]
> This project was AI-generated using Claude Opus 4.6 with human guidance and review.

## Getting Started

`teams-api` can be used in three ways:

1. As an MCP server for editors and AI tools. This is the recommended path for most users.
2. As a CLI for direct terminal use.
3. As a programmatic Node.js library. This is the advanced path and is documented near the end.

### Recommended: MCP via `npx @latest`

Prefer running from `npx` so you always pick up the latest published version while the project is evolving quickly.

If you are on macOS and already use a platform authenticator / FIDO2 passkey, start with auto-login:

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "teams-api@latest"],
      "env": {
        "TEAMS_AUTO": "true",
        "TEAMS_EMAIL": "you@example.com"
      }
    }
  }
}
```

If you are on Windows or Linux, or you want a visible browser login flow, use interactive login instead:

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "teams-api@latest"],
      "env": {
        "TEAMS_LOGIN": "true",
        "TEAMS_EMAIL": "you@example.com"
      }
    }
  }
}
```

> [!TIP]
> In both recommended setups, `teams-api` captures the full token bundle automatically during login. Most users never need to copy tokens manually.

### CLI

You can also use the CLI directly without installing anything:

```bash
npx -y -p teams-api@latest teams-api auth --login
npx -y -p teams-api@latest teams-api list-conversations --login --limit 20
```

If you use the CLI often, a global install is optional:

```bash
npm install -g teams-api
teams-api auth --login
```

### Advanced Topics

Manual token usage, debug-session auth, and programmatic Node.js usage are covered later in this README.

## Platform support

| Feature                | macOS          | Windows / Linux |
| ---------------------- | -------------- | --------------- |
| **Interactive login**  | Full support   | Full support    |
| **Auto-login (FIDO2)** | Full support   | Not supported   |
| **Debug session**      | Full support   | Full support    |
| **Direct token**       | Full support   | Full support    |
| **Token caching**      | macOS Keychain | Not available   |
| **CLI & MCP server**   | Full support   | Full support    |
| **Programmatic API**   | Full support   | Full support    |

## Authentication

Most users do not need to manage tokens manually.

If you use interactive login, auto-login, or a Chrome debug session, `teams-api` captures the full token bundle automatically from Teams web traffic. That includes the base `skypeToken` plus the extra bearer tokens used for profile resolution and reliable people/chat/channel search.

Those flows also detect the Teams chat region automatically from the intercepted request URLs, so most users do not need to set `--region` or `TEAMS_REGION`.

On macOS, prefer auto-login when you have a platform authenticator / FIDO2 passkey set up. On other platforms, use interactive login.

Direct token usage is the advanced/manual path.

| Method                | Description                                                                  | Automation       | Platform |
| --------------------- | ---------------------------------------------------------------------------- | ---------------- | -------- |
| **Auto-login**        | Playwright launches system Chrome and captures the full token bundle         | Fully unattended | macOS    |
| **Interactive login** | Opens a browser window and captures skype, middle-tier, and Substrate tokens | One-time manual  | All      |
| **Debug session**     | Connects to a running Chrome instance and captures the full token bundle     | Semi-manual      | All      |
| **Direct token**      | Provide a previously captured token or token bundle explicitly               | Manual           | All      |

### Auto-login (macOS only)

Requires macOS with a platform authenticator (e.g. Intune Company Portal) and a FIDO2 passkey enrolled. Fully unattended — no browser window appears.

### Interactive login (recommended for Windows / Linux)

The easiest cross-platform option. A browser window opens, you log in with any method your organization supports (password, MFA, passkey, etc.), and the token is captured automatically:

```bash
teams-api auth --login
```

Optionally pre-fill your email:

```bash
teams-api auth --login --email you@example.com
```

> [!NOTE]
> Interactive login uses Playwright's bundled Chromium. No system Chrome installation is required.

### Advanced / manual methods

**Debug session** — start Chrome with `--remote-debugging-port=9222`, navigate to Teams and log in, then run:

```bash
teams-api auth --debug-port 9222
```

**Direct token** — advanced/manual only. Extract `x-skypetoken` from browser DevTools (Network tab) and pass it directly:

```bash
teams-api list-conversations --token "<paste-token-here>" --region emea
```

If you want reliable people/chat/channel lookup and profile resolution on the direct-token path, also pass the extra bearer tokens captured from Teams requests:

```bash
teams-api find-people \
  --token "<paste-skype-token-here>" \
  --bearer-token "<paste-api-spaces-skype-bearer-token-here>" \
  --substrate-token "<paste-substrate-bearer-token-here>" \
  --region emea \
  --query "Jane Doe"
```

> [!TIP]
> Skip this section if you are using `--login`, `--auto`, `TEAMS_LOGIN`, or `TEAMS_AUTO`. Those modes capture the full token bundle automatically.

> [!TIP]
> Direct-token mode still needs an explicit region. See [API regions](#api-regions) below.

## CLI

Preferred without install:

```bash
npx -y -p teams-api@latest teams-api <command> [options]
```

Optional global install for frequent use:

```bash
npm install -g teams-api
teams-api <command> [options]
```

The examples below use `teams-api` for readability. If you are not installing globally, replace it with `npx -y -p teams-api@latest teams-api`.

### Auth flags (available on all commands)

| Flag                        | Description                                                                      |
| --------------------------- | -------------------------------------------------------------------------------- |
| `--login`                   | Interactive browser login (all platforms)                                        |
| `--auto`                    | Auto-acquire token via FIDO2 passkey (macOS)                                     |
| `--email <email>`           | Corporate email (required with `--auto`, optional otherwise)                     |
| `--token <token>`           | Use an existing skype token (advanced/manual)                                    |
| `--bearer-token <token>`    | Optional middle-tier bearer token (advanced/manual)                              |
| `--substrate-token <token>` | Optional Substrate bearer token (advanced/manual)                                |
| `--debug-port <port>`       | Chrome debug port (default: 9222)                                                |
| `--region <region>`         | API region override. Auto-detected for login/debug auth; required with `--token` |
| `--format <format>`         | Output format: json, text, md, toon                                              |
| `--output <file>`           | Export output to file (default format: md)                                       |

### Examples

```bash
# Acquire a token (interactive — all platforms)
teams-api auth --login

# Acquire a token (auto — macOS with FIDO2)
teams-api auth --auto --email you@example.com

# List conversations
teams-api list-conversations --login --limit 20 --format json

# Find a conversation by topic
teams-api find-conversation --auto --email you@example.com --query "Design Review"

# Find a 1:1 chat by person name
teams-api find-one-on-one --auto --email you@example.com --person-name "Jane Doe"

# Read messages (by topic name, person name, or direct ID)
teams-api get-messages --auto --email you@example.com --chat "Design Review"
teams-api get-messages --auto --email you@example.com --to "Jane Doe" --max-pages 5
teams-api get-messages --auto --email you@example.com --conversation-id "19:abc@thread.v2" --format json

# Newest-first order (API returns newest-first; default is oldest-first/chronological)
teams-api get-messages --auto --email you@example.com --chat "General" --order newest-first

# Send a message
teams-api send-message --auto --email you@example.com --to "Jane Doe" --content "Hello!"
teams-api send-message --auto --email you@example.com --chat "Design Review" --content "Status update"

# List members
teams-api get-members --auto --email you@example.com --chat "Design Review" --format md

# Get current user info
teams-api whoami --auto --email you@example.com

# Export messages to a file (default format: md)
teams-api get-messages --auto --email you@example.com --chat "General" --output exports/general.md

# Export as JSON to a file
teams-api get-messages --auto --email you@example.com --chat "General" --format json --output exports/general.json

# Toon format (fun ASCII art output)
teams-api list-conversations --auto --email you@example.com --format toon
```

## MCP server

The MCP server exposes Teams operations as tools for AI agents via stdio transport.

### Recommended setup

For most users, use one of these login-based configs. No manual token extraction is required.

**macOS (auto-login with FIDO2):**

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "teams-api@latest"],
      "env": {
        "TEAMS_AUTO": "true",
        "TEAMS_EMAIL": "you@example.com"
      }
    }
  }
}
```

**All platforms (interactive login):**

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "teams-api@latest"],
      "env": {
        "TEAMS_LOGIN": "true",
        "TEAMS_EMAIL": "you@example.com"
      }
    }
  }
}
```

> [!TIP]
> In the recommended configs above, `teams-api` captures the full token bundle automatically during login. You do not need to copy `skypeToken`, `bearerToken`, or `substrateToken` into the MCP config.

### Advanced: direct token configuration

Use this only if you already have tokens from another flow or need to avoid browser-based auth entirely.

**All platforms (direct token):**

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "teams-api@latest"],
      "env": {
        "TEAMS_TOKEN": "<paste-skype-token-here>",
        "TEAMS_BEARER_TOKEN": "<optional-api-spaces-skype-bearer-token>",
        "TEAMS_SUBSTRATE_TOKEN": "<optional-substrate-bearer-token>",
        "TEAMS_REGION": "emea"
      }
    }
  }
}
```

> [!TIP]
> If you do use direct tokens, `teams-api auth --login` prints the full token object as JSON. For basic chat operations, `skypeToken` is enough. For reliable people/chat/channel search and profile resolution, also pass `bearerToken` and `substrateToken`.

### Environment variables

| Variable                | Description                                                          |
| ----------------------- | -------------------------------------------------------------------- |
| `TEAMS_TOKEN`           | Pre-existing skype token                                             |
| `TEAMS_BEARER_TOKEN`    | Optional middle-tier bearer token                                    |
| `TEAMS_SUBSTRATE_TOKEN` | Optional Substrate bearer token                                      |
| `TEAMS_REGION`          | API region override. Required with `TEAMS_TOKEN`; optional otherwise |
| `TEAMS_EMAIL`           | Corporate email for auto-login or interactive login                  |
| `TEAMS_AUTO`            | Set to `true` to enable auto-login (macOS + FIDO2)                   |
| `TEAMS_LOGIN`           | Set to `true` to enable interactive browser login                    |
| `TEAMS_DEBUG_PORT`      | Chrome debug port (default: 9222)                                    |

### Available tools

All MCP tools accept an optional `format` parameter (`json`, `text`, `md`, or `toon`). Default format is `toon`.

| Tool                       | Description                                           |
| -------------------------- | ----------------------------------------------------- |
| `teams_list_conversations` | List available conversations                          |
| `teams_find_conversation`  | Find a conversation by topic or member name           |
| `teams_find_one_on_one`    | Find a 1:1 chat with a person                         |
| `teams_find_people`        | Search the organization directory                     |
| `teams_find_chats`         | Search chats by name or member                        |
| `teams_get_messages`       | Get messages from a conversation                      |
| `teams_send_message`       | Send a message to a conversation                      |
| `teams_get_members`        | List members of a conversation                        |
| `teams_get_transcript`     | Get a meeting transcript from a recorded conversation |
| `teams_whoami`             | Get the authenticated user's display name             |

See [SKILL.md](SKILL.md) for detailed tool descriptions and typical agent workflows.

## API regions

The Teams Chat Service URL varies by region. Login-based and debug-session auth detect it automatically. You only need to set `--region` or `TEAMS_REGION` when you are supplying tokens directly or want to force an override:

| Region | Base URL                                     |
| ------ | -------------------------------------------- |
| `apac` | `https://apac.ng.msg.teams.microsoft.com/v1` |
| `emea` | `https://emea.ng.msg.teams.microsoft.com/v1` |
| `amer` | `https://amer.ng.msg.teams.microsoft.com/v1` |

## Known limitations

- Token lifetime is ~24 hours. After expiry, you must re-acquire.
- The Teams Chat Service REST API is undocumented and may change without notice.
- Auto-login requires macOS, system Chrome, a platform authenticator, and a FIDO2 passkey. On other platforms, use interactive login (`--login`) instead.
- Token caching (macOS Keychain) is only available on macOS. On other platforms, pass the token directly or re-run interactive login each time.
- The members API returns empty display names for 1:1 chat participants. Use `findOneOnOneConversation()` to resolve names from message history.
- Reaction actor identities come from the `emotions` field in message payloads. Parsing handles both JSON-string and array formats.

## Programmatic API

This is the advanced integration path. Most users should start with MCP or CLI instead.

Install the package in your project:

```bash
npm install teams-api
```

Example:

```typescript
import { TeamsClient } from "teams-api";

// Interactive login — opens a browser, you log in manually (all platforms)
const client = await TeamsClient.fromInteractiveLogin();

// Or auto-login via platform authenticator (macOS + FIDO2 passkey)
const autoClient = await TeamsClient.fromAutoLogin({
  email: "you@example.com",
});

// Advanced/manual: create a client from previously captured tokens
const manualClient = TeamsClient.fromToken("skype-token-here", "emea");

const conversations = await client.listConversations();
const messages = await client.getMessages(conversations[0].id, {
  maxPages: 5,
  onProgress: (count) => console.log(`Fetched ${count} messages`),
});

await client.sendMessage(conversations[0].id, "Hello from the API!");

const oneOnOne = await client.findOneOnOneConversation("Jane Doe");
const members = await client.getMembers(conversations[0].id);
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, architecture, and implementation notes.

## License

MIT
