# Contributing

Development guide for the teams-api project. For user-facing documentation (installation, CLI usage, MCP setup), see [README.md](README.md).

### Documentation hierarchy

- **[README.md](README.md)** — user-facing: installation, CLI usage, MCP setup, API quick start.
- **CONTRIBUTING.md** (this file) — developer-facing: architecture, testing, code style, release process.
- **[.github/copilot-instructions.md](.github/copilot-instructions.md)** — AI agent behavior: tool calling, MCP workflows, reverse engineering, domain knowledge that only AI agents need.

Keep each file focused on its audience. Do not duplicate content across files.

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
src/
  types.ts              All public interfaces and types
  constants.ts          Message type constants and type guards
  html-utils.ts         HTML entity decoding utilities
  region.ts             Region resolution and validation
  teams-client.ts       Public API class (TeamsClient) — the main entry point
  token-store.ts        Cross-platform token persistence (via credential-store)
  credential-store.ts   Platform credential stores (Keychain, DPAPI, secret-tool)
  platform.ts           Platform detection and auto-login eligibility
  smart-login.ts        Smart login — auto-login with interactive fallback
  browser-runtime.ts    Installed-browser detection and Playwright launch helpers
  cli.ts                Commander-based CLI (driven by actions/definitions)
  mcp-server.ts         MCP server with stdio transport
  server-instructions.ts  MCP server instructions and CLI guide content
  api/
    common.ts           Shared HTTP utilities and headers
    chat-service.ts     Chat Service REST calls (conversations, messages, members)
    middle-tier.ts      Middle-tier profile and presence lookups
    substrate.ts        Substrate search API (people, chats)
    transcripts.ts      VTT transcript fetching and parsing
  auth/
    auto-login.ts       Auto-login via system Chrome + FIDO2 passkey
    debug-session.ts    CDP debug session token capture
    interactive.ts      Interactive browser login flow
    page-diagnostics.ts   Page state analysis and error detection
    token-capture.ts      CDP Fetch interception for token extraction
  actions/
    definitions.ts      Action registry — imports and assembles all actions
    conversation-actions.ts  List, find, and 1:1 conversation actions
    message-actions.ts       Get, send, edit, delete message actions
    search-actions.ts        People and chat search actions
    utility-actions.ts       Whoami, get-members, get-transcript actions
    formatters.ts       Output formatting utilities and type definitions
    conversation-resolution.ts  Conversation ID resolution logic
```

### Data flow

```
TeamsClient (public API)
  ├── auth/*          — acquires a TeamsToken via one of three strategies
  └── api/*           — stateless HTTP calls using TeamsToken
        ├── chat-service.ts  — conversations, messages, members
        ├── middle-tier.ts   — profiles, presence
        ├── substrate.ts     — people/chat search
        └── transcripts.ts   — meeting transcript VTT
```

TeamsClient is the only public-facing class. It accepts a TeamsToken (from any auth strategy) and delegates to the stateless `api/*` functions. The CLI and MCP server both consume TeamsClient via the unified action definitions in `actions/definitions.ts`.

### Authentication strategies

1. **Smart login** (`acquireTokenViaSmartLogin` in `src/smart-login.ts`): The default auth strategy. Attempts auto-login first (macOS), falls back to interactive login on other platforms or when auto-login fails. Cached tokens are reused automatically via the platform credential store.

2. **Interactive login** (`acquireTokenViaInteractiveLogin` in `src/auth/interactive.ts`): Opens a visible browser window and navigates to Teams. Prefers an installed browser (Edge, Chrome) when available, falling back to Playwright's bundled Chromium. The user completes the login manually. Works on all platforms.

3. **Auto-login** (`acquireTokenViaAutoLogin` in `src/auth/auto-login.ts`): Launches system Chrome via Playwright persistent context and completes FIDO2 passkey authentication automatically. macOS only. Usually invoked via smart login rather than directly.

4. **Debug session** (`acquireTokenViaDebugSession` in `src/auth/debug-session.ts`): Connects to a running Chrome instance via puppeteer-core CDP, finds the Teams tab, enables Fetch interception, triggers a page reload, and captures the `x-skypetoken` header.

All strategies use the same CDP Fetch interception pattern in `src/auth/token-capture.ts` to extract the token from live network requests.

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

### Output format system

All actions support two output formats via the `format` parameter:

| Format       | Default | Description |
| ------------ | ------- | ----------- |
| `concise`    | Yes     | Light Markdown optimized for actionability. Includes identifiers and decision-critical fields. Nested collections may be summarized with counts/previews. |
| `detailed`   | No      | Full JSON — the raw result object via `JSON.stringify`. |

Each `ActionDefinition` implements a single `formatConcise(result)` method that renders the result as Markdown. The `detailed` format is handled generically by `formatOutput()` — no per-action code needed.

**Concise format rules:**
- Preserve next-action capability without requiring `detailed`.
- Include stable identifiers that enable follow-up operations (message IDs, conversation IDs, MRI/object IDs when relevant).
- Include decision-critical fields for the operation, but avoid low-value noise.
- Nested collections may be summarized (counts, matched subsets, attachment summaries) unless full expansion is required for the immediate next action.
- Avoid arbitrary truncation of fields you choose to display.
- Omit empty/null fields rather than showing placeholders.
- Use Markdown tables for list data and bullet lists for single-item results.

**Detailed format rule:**
- Always return full JSON with complete structure and values.

### Concise output contracts by action

Use this table as the review checklist for formatter changes.

| Action | Concise must include | Concise may summarize |
| ------ | -------------------- | --------------------- |
| `list-conversations` | Conversation ID, topic, thread type, last-message indicator | Ancillary conversation metadata not needed for selecting a target conversation |
| `find-conversation` | Conversation ID, thread type, topic | Optional timestamps/nullable fields |
| `find-one-on-one` | Conversation ID, matched member display name | Search diagnostics |
| `find-people` | Display name, email, MRI, object ID (if present) | Optional profile attributes |
| `find-chats` | Thread ID, thread type, member count, match-relevant members | Full member roster |
| `get-messages` | Message ID, sender, timestamp, message body, quoted message reference IDs | Reactions/followers/mentions as counts or compact lists; attachment internals |
| `send-message` | Target conversation label, message ID, send/schedule timestamp | Transport-level details |
| `edit-message` | Conversation label, message ID, edit timestamp | Transport-level details |
| `delete-message` | Conversation label, message ID | Transport-level details |
| `add-reaction` | Conversation label, message ID, resolved reaction key | Transport-level details |
| `remove-reaction` | Conversation label, message ID, resolved reaction key | Transport-level details |
| `get-members` | Member ID, name, role, member type | Non-actionable profile enrichment |
| `whoami` | Display name, region | Internal auth metadata |
| `get-transcript` | Meeting title and readable transcript content | Raw VTT structure unless `rawVtt` is requested |
| `download-file` | File name, file type, size, content type, saved location | Binary payload internals in Markdown text (binary is returned in content blocks) |

This design follows [Anthropic's tool design guidance](https://docs.anthropic.com/en/docs/build-with-claude/tool-use/best-practices) — a `response_format` enum with concise (~1/3 tokens) and detailed (full data) modes.

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

The `properties.emotions` field in message payloads has inconsistent formatting — sometimes it's a JSON string, sometimes a raw array. The `parseReactions` helper in `api/chat-service.ts` handles both formats and fails gracefully on malformed data.

### System stream filtering

Teams returns several system streams alongside real conversations (annotations, notifications, mentions, threads, notes). `listConversations` filters these out by default. The full list of filtered types is in the `SYSTEM_STREAM_TYPES` constant in `types.ts`.

## Debug telemetry

`src/telemetry.ts` provides a local debug logging facility for contributors investigating tool behavior.

**Off by default.** Enable with `TEAMS_TELEMETRY=true`.

When enabled, every tool call is appended as a JSON line to a local JSONL file:

| Platform | Path |
| --- | --- |
| macOS | `~/Library/Application Support/teams-api/telemetry.jsonl` |
| Linux | `$XDG_DATA_HOME/teams-api/telemetry.jsonl` (default: `~/.local/share/teams-api/telemetry.jsonl`) |
| Windows | `%APPDATA%\teams-api\telemetry.jsonl` |

Each record contains: tool name, format, full input parameters, raw API result, formatted output string, duration (ms), and timestamp. Auth events and errors (with full stack traces) are also recorded.

**This is internal tooling only.** It is not mentioned in the user-facing README and should not be presented as a feature to end users. The telemetry file grows unboundedly — clear it manually with `> /path/to/telemetry.jsonl` when no longer needed.

To enable for your local MCP server in VS Code, add `"TEAMS_TELEMETRY": "true"` to the `env` block of the `teams` server in your MCP config.
