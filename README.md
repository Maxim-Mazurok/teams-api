# teams-api

AI-native Microsoft Teams integration — read conversations, send messages, and manage members via the Teams Chat Service REST API.

Designed for autonomous AI agents that need to interact with Teams: read messages, reply to people, monitor conversations, and participate in team workflows.

> [!NOTE]
> This project was AI-generated using Claude Opus 4.6 with human guidance and review.

## Quick start

```typescript
import { TeamsClient } from "./src/teams-client.js";

// From an existing token (~24h lifetime)
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

## Authentication

All access requires a **skype token** from an authenticated Teams web session. Token lifetime is ~24 hours.

| Method            | Description                                                               | Automation       |
| ----------------- | ------------------------------------------------------------------------- | ---------------- |
| **Auto-login**    | Playwright launches system Chrome, completes FIDO2 passkey authentication | Fully unattended |
| **Debug session** | Connects to a running Chrome instance via Chrome DevTools Protocol        | Semi-manual      |
| **Direct token**  | Provide a previously captured skype token string                          | Manual           |

Auto-login requires macOS with a platform authenticator (e.g. Intune Company Portal) and a FIDO2 passkey enrolled.

## CLI

Run commands with `npx tsx src/cli.ts`.

### Auth flags (available on all commands)

| Flag                  | Description                              |
| --------------------- | ---------------------------------------- |
| `--auto`              | Auto-acquire token via FIDO2 passkey     |
| `--email <email>`     | Corporate email (required with `--auto`) |
| `--token <token>`     | Use an existing skype token              |
| `--debug-port <port>` | Chrome debug port (default: 9222)        |
| `--region <region>`   | API region (default: apac)               |

### Examples

```bash
# Acquire a token
npx tsx src/cli.ts auth --auto --email you@example.com

# List conversations
npx tsx src/cli.ts list-conversations --token "$TOKEN" --limit 20 --json

# Find a conversation by topic
npx tsx src/cli.ts find-conversation --token "$TOKEN" --query "Design Review"

# Find a 1:1 chat by person name
npx tsx src/cli.ts find-one-on-one --token "$TOKEN" --person-name "Jane Doe"

# Read messages (by topic name, person name, or direct ID)
npx tsx src/cli.ts get-messages --token "$TOKEN" --chat "Design Review"
npx tsx src/cli.ts get-messages --token "$TOKEN" --to "Jane Doe" --max-pages 5
npx tsx src/cli.ts get-messages --token "$TOKEN" --conversation-id "19:abc@thread.v2" --json

# Send a message
npx tsx src/cli.ts send-message --token "$TOKEN" --to "Jane Doe" --content "Hello!"
npx tsx src/cli.ts send-message --token "$TOKEN" --chat "Design Review" --content "Status update"

# List members
npx tsx src/cli.ts get-members --token "$TOKEN" --chat "Design Review" --json

# Get current user info
npx tsx src/cli.ts whoami --token "$TOKEN"
```

## MCP server

The MCP server exposes Teams operations as tools for AI agents via stdio transport.

### Configuration

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

### Environment variables

| Variable           | Description                        |
| ------------------ | ---------------------------------- |
| `TEAMS_TOKEN`      | Pre-existing skype token           |
| `TEAMS_REGION`     | API region (default: apac)         |
| `TEAMS_EMAIL`      | Corporate email for auto-login     |
| `TEAMS_AUTO`       | Set to `true` to enable auto-login |
| `TEAMS_DEBUG_PORT` | Chrome debug port (default: 9222)  |

### Available tools

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
```

## Known limitations

- Token lifetime is ~24 hours. After expiry, you must re-acquire.
- The Teams Chat Service REST API is undocumented and may change without notice.
- Auto-login requires macOS, system Chrome, a platform authenticator, and a FIDO2 passkey.
- The members API returns empty display names for 1:1 chat participants. Use `findOneOnOneConversation()` to resolve names from message history.
- Reaction actor identities come from the `emotions` field in message payloads. Parsing handles both JSON-string and array formats.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for architecture details, implementation notes, and development guidelines.

## License

MIT
