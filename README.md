# teams-api

AI-native Microsoft Teams integration — read conversations, send messages, and manage members via the Teams Chat Service REST API.

Designed for autonomous AI agents that need to interact with Teams: read messages, reply to people, monitor conversations, and participate in team workflows.

> [!NOTE]
> This project was AI-generated using Claude Opus 4.6 with human guidance and review.

## Installation

```bash
npm install -g teams-api-mcp
```

Or run directly via `npx` without installing:

```bash
npx -y teams-api-mcp
```

## Quick start

```typescript
import { TeamsClient } from "teams-api-mcp";

// Interactive login — opens a browser, you log in manually (all platforms)
const client = await TeamsClient.fromInteractiveLogin({ region: "emea" });

// Or from an existing token (~24h lifetime)
const client = TeamsClient.fromToken("skype-token-here", "apac");

// Or auto-login via platform authenticator (macOS + FIDO2 passkey)
const client = await TeamsClient.fromAutoLogin({
  email: "you@example.com",
});

// List conversations
const conversations = await client.listConversations();

// Read messages (auto-paginates)
const messages = await client.getMessages(conversations[0].id, {
  maxPages: 5,
  onProgress: (count) => console.log(`Fetched ${count} messages`),
});

// Send a message
await client.sendMessage(conversations[0].id, "Hello from the API!");

// Find a 1:1 chat by person name
const oneOnOne = await client.findOneOnOneConversation("Jane Doe");

// Get conversation members
const members = await client.getMembers(conversations[0].id);
```

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

All access requires a **skype token** from an authenticated Teams web session. Token lifetime is ~24 hours.

| Method                | Description                                                                   | Automation       | Platform |
| --------------------- | ----------------------------------------------------------------------------- | ---------------- | -------- |
| **Interactive login** | Opens a browser window — you log in manually, token is captured automatically | One-time manual  | All      |
| **Auto-login**        | Playwright launches system Chrome, completes FIDO2 passkey authentication     | Fully unattended | macOS    |
| **Debug session**     | Connects to a running Chrome instance via Chrome DevTools Protocol            | Semi-manual      | All      |
| **Direct token**      | Provide a previously captured skype token string                              | Manual           | All      |

### Interactive login (recommended for Windows / Linux)

The easiest cross-platform option. A browser window opens, you log in with any method your organization supports (password, MFA, passkey, etc.), and the token is captured automatically:

```bash
teams-api auth --login --region emea
```

Optionally pre-fill your email:

```bash
teams-api auth --login --email you@example.com --region emea
```

> [!NOTE]
> Interactive login uses Playwright's bundled Chromium. No system Chrome installation is required.

### Auto-login (macOS only)

Requires macOS with a platform authenticator (e.g. Intune Company Portal) and a FIDO2 passkey enrolled. Fully unattended — no browser window appears.

### Other methods

**Debug session** — start Chrome with `--remote-debugging-port=9222`, navigate to Teams and log in, then run:

```bash
teams-api auth --debug-port 9222 --region emea
```

**Direct token** — extract `x-skypetoken` from browser DevTools (Network tab) and pass it directly:

```bash
teams-api list-conversations --token "<paste-token-here>" --region emea
```

> [!TIP]
> Replace `emea` with your region (`apac`, `emea`, or `amer`). See [API regions](#api-regions) below.

## CLI

After installing globally (`npm install -g teams-api-mcp`), run commands with `teams-api`.

### Auth flags (available on all commands)

| Flag                  | Description                                                  |
| --------------------- | ------------------------------------------------------------ |
| `--login`             | Interactive browser login (all platforms)                    |
| `--auto`              | Auto-acquire token via FIDO2 passkey (macOS)                 |
| `--email <email>`     | Corporate email (required with `--auto`, optional otherwise) |
| `--token <token>`     | Use an existing skype token                                  |
| `--debug-port <port>` | Chrome debug port (default: 9222)                            |
| `--region <region>`   | API region (default: apac)                                   |
| `--format <format>`   | Output format: json, text, md, toon                          |
| `--output <file>`     | Export output to file (default format: md)                   |

### Examples

```bash
# Acquire a token (interactive — all platforms)
teams-api auth --login --region emea

# Acquire a token (auto — macOS with FIDO2)
teams-api auth --auto --email you@example.com

# List conversations
teams-api list-conversations --login --region emea --limit 20 --format json

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

### Configuration

**macOS (auto-login with FIDO2):**

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "teams-api-mcp"],
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
      "args": ["-y", "teams-api-mcp"],
      "env": {
        "TEAMS_LOGIN": "true",
        "TEAMS_EMAIL": "you@example.com",
        "TEAMS_REGION": "emea"
      }
    }
  }
}
```

**All platforms (direct token):**

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "teams-api-mcp"],
      "env": {
        "TEAMS_TOKEN": "<paste-skype-token-here>",
        "TEAMS_REGION": "emea"
      }
    }
  }
}
```

> [!TIP]
> To get a token for the MCP config, run `teams-api auth --login --region emea` and copy the `skypeToken` value from the output.

### Environment variables

| Variable           | Description                                         |
| ------------------ | --------------------------------------------------- |
| `TEAMS_TOKEN`      | Pre-existing skype token                            |
| `TEAMS_REGION`     | API region (default: apac)                          |
| `TEAMS_EMAIL`      | Corporate email for auto-login or interactive login |
| `TEAMS_AUTO`       | Set to `true` to enable auto-login (macOS + FIDO2)  |
| `TEAMS_LOGIN`      | Set to `true` to enable interactive browser login   |
| `TEAMS_DEBUG_PORT` | Chrome debug port (default: 9222)                   |

### Available tools

All MCP tools accept an optional `format` parameter (`json`, `text`, `md`, or `toon`). Default format is `toon`.

| Tool                       | Description                               |
| -------------------------- | ----------------------------------------- |
| `teams_list_conversations` | List available conversations              |
| `teams_find_conversation`  | Find a conversation by topic name         |
| `teams_find_one_on_one`    | Find a 1:1 chat with a person             |
| `teams_get_messages`       | Get messages from a conversation          |
| `teams_send_message`       | Send a message to a conversation          |
| `teams_get_members`        | List members of a conversation            |
| `teams_whoami`             | Get the authenticated user's display name |

See [SKILL.md](SKILL.md) for detailed tool descriptions and typical agent workflows.

## API regions

The Teams Chat Service URL varies by region. Use the `region` parameter or `TEAMS_REGION` env var:

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

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, architecture, and implementation notes.

## License

MIT
