# teams-api

AI-native Microsoft Teams integration — read conversations, send messages, and manage members via the Teams Chat Service REST API.

Designed for autonomous AI agents that need to interact with Teams: read messages, reply to people, monitor conversations, and participate in team workflows.

> [!NOTE]
> This project was AI-generated using Claude Opus 4.6 with human guidance and review.

## Quick start

```typescript
import { TeamsClient } from "./src/teams-client.js";

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
npx tsx src/cli.ts auth --login --region emea
```

Optionally pre-fill your email:

```bash
npx tsx src/cli.ts auth --login --email you@example.com --region emea
```

> [!NOTE]
> Interactive login uses Playwright's bundled Chromium. No system Chrome installation is required.

### Auto-login (macOS only)

Requires macOS with a platform authenticator (e.g. Intune Company Portal) and a FIDO2 passkey enrolled. Fully unattended — no browser window appears.

### Other methods

**Debug session** — start Chrome with `--remote-debugging-port=9222`, navigate to Teams and log in, then run:

```bash
npx tsx src/cli.ts auth --debug-port 9222 --region emea
```

**Direct token** — extract `x-skypetoken` from browser DevTools (Network tab) and pass it directly:

```bash
npx tsx src/cli.ts list-conversations --token "<paste-token-here>" --region emea
```

> [!TIP]
> Replace `emea` with your region (`apac`, `emea`, or `amer`). See [API regions](#api-regions) below.

## CLI

Run commands with `npx tsx src/cli.ts`.

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
npx tsx src/cli.ts auth --login --region emea

# Acquire a token (auto — macOS with FIDO2)
npx tsx src/cli.ts auth --auto --email you@example.com

# List conversations
npx tsx src/cli.ts list-conversations --login --region emea --limit 20 --format json

# Find a conversation by topic
npx tsx src/cli.ts find-conversation --auto --email you@example.com --query "Design Review"

# Find a 1:1 chat by person name
npx tsx src/cli.ts find-one-on-one --auto --email you@example.com --person-name "Jane Doe"

# Read messages (by topic name, person name, or direct ID)
npx tsx src/cli.ts get-messages --auto --email you@example.com --chat "Design Review"
npx tsx src/cli.ts get-messages --auto --email you@example.com --to "Jane Doe" --max-pages 5
npx tsx src/cli.ts get-messages --auto --email you@example.com --conversation-id "19:abc@thread.v2" --format json

# Send a message
npx tsx src/cli.ts send-message --auto --email you@example.com --to "Jane Doe" --content "Hello!"
npx tsx src/cli.ts send-message --auto --email you@example.com --chat "Design Review" --content "Status update"

# List members
npx tsx src/cli.ts get-members --auto --email you@example.com --chat "Design Review" --format md

# Get current user info
npx tsx src/cli.ts whoami --auto --email you@example.com

# Export messages to a file (default format: md)
npx tsx src/cli.ts get-messages --auto --email you@example.com --chat "General" --output exports/general.md

# Export as JSON to a file
npx tsx src/cli.ts get-messages --auto --email you@example.com --chat "General" --format json --output exports/general.json

# Toon format (fun ASCII art output)
npx tsx src/cli.ts list-conversations --auto --email you@example.com --format toon
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
      "args": ["-y", "tsx", "/absolute/path/to/teams-api/src/mcp-server.ts"],
      "env": {
        "TEAMS_AUTO": "true",
        "TEAMS_EMAIL": "you@example.com"
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
      "args": ["-y", "tsx", "/path/to/teams-api/src/mcp-server.ts"],
      "env": {
        "TEAMS_TOKEN": "<paste-skype-token-here>",
        "TEAMS_REGION": "emea"
      }
    }
  }
}
```

> [!TIP]
> To get a token for the MCP config, run `npx tsx src/cli.ts auth --login --region emea` and copy the `skypeToken` value from the output.

````

### Environment variables

| Variable           | Description                                          |
| ------------------ | ---------------------------------------------------- |
| `TEAMS_TOKEN`      | Pre-existing skype token                             |
| `TEAMS_REGION`     | API region (default: apac)                           |
| `TEAMS_EMAIL`      | Corporate email for auto-login or interactive login  |
| `TEAMS_AUTO`       | Set to `true` to enable auto-login (macOS + FIDO2)   |
| `TEAMS_LOGIN`      | Set to `true` to enable interactive browser login    |
| `TEAMS_DEBUG_PORT` | Chrome debug port (default: 9222)                    |

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

## Testing

```bash
# Unit tests (mocked fetch, no network)
npm test

# Integration tests (requires a live token)
TEAMS_TOKEN=<token> TEAMS_REGION=apac npm run test:integration

# E2E tests (requires macOS + FIDO2 passkey)
TEAMS_EMAIL=you@example.com npm run test:e2e

# Watch mode
npm run test:watch
````

## Known limitations

- Token lifetime is ~24 hours. After expiry, you must re-acquire.
- The Teams Chat Service REST API is undocumented and may change without notice.
- Auto-login requires macOS, system Chrome, a platform authenticator, and a FIDO2 passkey. On other platforms, use interactive login (`--login`) instead.
- Token caching (macOS Keychain) is only available on macOS. On other platforms, pass the token directly or re-run interactive login each time.
- The members API returns empty display names for 1:1 chat participants. Use `findOneOnOneConversation()` to resolve names from message history.
- Reaction actor identities come from the `emotions` field in message payloads. Parsing handles both JSON-string and array formats.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for architecture details, implementation notes, and development guidelines.

## License

MIT
