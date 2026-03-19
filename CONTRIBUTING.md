# Contributing

Development guide for the teams-api project. For user-facing documentation (installation, CLI usage, MCP setup), see [README.md](README.md).

## Getting started

```bash
git clone https://github.com/Maxim-Mazurok/teams-api.git
cd teams-api
npm install
```

To run commands from source, use `npx -y tsx src/cli.ts` instead of `teams-api`:

```bash
npx -y tsx src/cli.ts auth --login --region emea
npx -y tsx src/cli.ts list-conversations --login --region emea
```

### MCP server from source

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

## Architecture

```
src/types.ts          All public interfaces and types
src/api.ts            Low-level REST calls to Teams Chat Service
src/auth.ts           Token acquisition (Playwright auto-login + CDP debug session)
src/teams-client.ts   Public API class (TeamsClient) — the main entry point
src/cli.ts            Commander-based CLI
src/mcp-server.ts     MCP server with stdio transport
```

### Data flow

```
TeamsClient (public API)
  ├── auth.ts       — acquires a TeamsToken via one of two strategies
  └── api.ts        — stateless HTTP calls using TeamsToken
        └── Teams Chat Service REST API
              https://{region}.ng.msg.teams.microsoft.com/v1
```

`TeamsClient` is the only public-facing class. It accepts a `TeamsToken` (from any auth strategy) and delegates to the stateless `api.ts` functions. The CLI and MCP server both consume `TeamsClient`.

### Authentication strategies

1. **Interactive login** (`acquireTokenViaInteractiveLogin`): Opens a visible Chromium browser window (Playwright's bundled browser) and navigates to Teams. The user completes the login manually using any method their organization supports (password, MFA, passkey, etc.). Once Teams loads, the skype token is captured via CDP Fetch interception during a page reload. Works on all platforms without requiring system Chrome or FIDO2 passkeys.

2. **Auto-login** (`acquireTokenViaAutoLogin`): Launches system Chrome via Playwright persistent context, navigates to the Teams web app, fills the email on the Microsoft Entra ID login page, waits for FIDO2 passkey authentication to complete, then intercepts the `x-skypetoken` header via CDP Fetch interception during a page reload. macOS only.

3. **Debug session** (`acquireTokenViaDebugSession`): Connects to a running Chrome instance via puppeteer-core CDP, finds the Teams tab, enables Fetch interception, triggers a page reload, and captures the `x-skypetoken` header.

All three strategies use the same CDP Fetch interception pattern to extract the token from live network requests.

### Token lifecycle

- The skype token is obtained from the Teams web application at `https://teams.cloud.microsoft/`
- Token lifetime is approximately 24 hours (the `skypetoken_asm` cookie has `Max-Age=86399`)
- Authentication header format: `Authentication: skypetoken=<token>` (note: `Authentication`, not `Authorization`)
- When the token expires, any API call will return `401` and a new token must be acquired

## API documentation

See [docs/findings.md](docs/findings.md) for detailed REST API endpoint documentation, including:

- Request/response formats for all endpoints
- Message structure and field descriptions
- Reactions format (JSON string vs. array inconsistency)
- Mentions format
- Conversation types and system stream filtering
- Deleted message detection
- Quoted message (reply) parsing
- Worker intercept findings from earlier browser extension research

## Development

### Prerequisites

- Node.js 20+
- npm

### Scripts

| Command                    | Description                                 |
| -------------------------- | ------------------------------------------- |
| `npm test`                 | Run unit tests                              |
| `npm run test:unit`        | Run unit tests only                         |
| `npm run test:integration` | Run integration tests (needs `TEAMS_TOKEN`) |
| `npm run test:e2e`         | Run E2E tests (needs `TEAMS_EMAIL`)         |
| `npm run test:watch`       | Run tests in watch mode                     |
| `npm run type-check`       | TypeScript type checking                    |
| `npm run lint`             | Check formatting with Prettier              |
| `npm run format`           | Auto-format with Prettier                   |
| `npm run mcp`              | Start MCP server                            |

### Testing

**Unit tests** (`tests/unit/`): Mock `fetch` globally and test the API and client layers in isolation. No network access required.

```bash
npm test
```

**Integration tests** (`tests/integration/`): Hit the real Teams API. Skipped by default — set `TEAMS_TOKEN` and `TEAMS_REGION` env vars to run.

```bash
TEAMS_TOKEN=<token> TEAMS_REGION=apac npm run test:integration
```

**E2E tests** (`tests/e2e/`): Full auto-login → read → write flow. Skipped by default — set `TEAMS_EMAIL` to run. Requires macOS with a platform authenticator and FIDO2 passkey.

```bash
TEAMS_EMAIL=you@example.com npm run test:e2e
```

### Code style

- TypeScript strict mode
- Prettier for formatting
- No default exports
- Named exports only
- ESM syntax in `.ts` files, CommonJS in `package.json`

### Releases

Releases are fully automated on every push to `main`. After the CI matrix passes, GitHub Actions runs `semantic-release` to:

- determine the next version
- create the Git tag and GitHub release
- publish the package to npm
- commit the updated `package.json`, `package-lock.json`, and `CHANGELOG.md` back to `main`

Versioning follows conventional commit prefixes:

- `feat` triggers a minor release
- `BREAKING CHANGE` in the commit footer triggers a major release
- `fix`, `docs`, `chore`, `ci`, `refactor`, `test`, `style`, `build`, `perf`, and `revert` trigger a patch release

Use conventional commit messages on PR titles and squash merges so every push to `main` produces the expected automated release.

If you use VS Code's Generate Commit Message action, this repo includes workspace settings and a dedicated Copilot instruction file so generated commit messages follow the expected Conventional Commit format by default.

## Implementation notes

### 1:1 chat name resolution

The Teams members API returns empty `friendlyName` / `displayName` for 1:1 chat participants. The `findOneOnOneConversation` method works around this by scanning recent message senders in untitled chats for a name match. It also checks the self-chat (`48:notes` conversation) for the current user's own name.

### Reactions parsing

The `properties.emotions` field in message payloads has inconsistent formatting — sometimes it's a JSON string, sometimes a raw array. The `parseReactions` helper in `api.ts` handles both formats and fails gracefully on malformed data.

### System stream filtering

Teams returns several system streams alongside real conversations (annotations, notifications, mentions, threads, notes). `listConversations` filters these out by default. The full list of filtered types is in the `SYSTEM_STREAMS` constant in `teams-client.ts`.
