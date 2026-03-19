# AI Agent Instructions

Instructions for AI agents working on this codebase. For human-readable architecture, setup, and development guidance, see [CONTRIBUTING.md](../CONTRIBUTING.md).

## Project context

- **What**: AI-native Microsoft Teams integration — CLI, MCP server, and programmatic TypeScript API
- **API source**: All Teams REST endpoints were reverse-engineered from the Teams web client (no public docs exist)
- **Single source of truth**: `src/actions.ts` defines all actions consumed by the CLI, MCP server, and tests
- **Entry point**: `TeamsClient` in `src/teams-client.ts` — the only public-facing class

## Code conventions

- TypeScript strict mode, Prettier for formatting
- Named exports only (no default exports)
- ESM syntax in `.ts` files, `"type": "commonjs"` in `package.json`
- Stateless API layer (`src/api.ts`) — all functions accept a token and return data
- `TeamsClient` delegates to `api.ts` for HTTP and `auth.ts` for token acquisition

## Build and test

```bash
npm install             # setup
npm test                # unit tests (mocked fetch, no network)
npm run type-check      # TypeScript checking
npm run lint            # Prettier check
npm run format          # auto-format
```

Integration tests need `TEAMS_TOKEN` + `TEAMS_REGION`. E2E tests need `TEAMS_EMAIL` (macOS + FIDO2 passkey).

## Dual-token authentication

Teams uses **two independent tokens** captured during a single CDP Fetch interception flow in `src/auth.ts`:

| Token        | Header                               | Audience                                             | Used for                                                                 |
| ------------ | ------------------------------------ | ---------------------------------------------------- | ------------------------------------------------------------------------ |
| Skype token  | `Authentication: skypetoken=<token>` | Chat Service (`{region}.ng.msg.teams.microsoft.com`) | Messages, threads, members, conversations                                |
| Bearer token | `Authorization: Bearer <token>`      | `api.spaces.skype.com` (MSAL)                        | Middle-tier APIs on `teams.cloud.microsoft` (profiles, presence, search) |

These tokens **cannot** be derived from each other. Both are captured in all three auth strategies (auto-login, interactive, debug session) and persisted together in the macOS Keychain via `src/token-store.ts`.

If you discover a new endpoint that requires a different token or audience, you must update the CDP Fetch interception in `src/auth.ts` to also capture it.

## Known API surfaces

| Host                                      | Auth         | Examples                                   |
| ----------------------------------------- | ------------ | ------------------------------------------ |
| `{region}.ng.msg.teams.microsoft.com/v1/` | Skype token  | Conversations, messages, thread members    |
| `teams.cloud.microsoft/api/mt/{region}/`  | Bearer token | `fetchShortProfile`, user search, presence |
| `presence.teams.microsoft.com/`           | Bearer token | Availability status                        |

See [docs/findings.md](../docs/findings.md) for detailed endpoint documentation.

## Reverse engineering Teams APIs

When you need to discover new endpoints or understand undocumented behavior, use the Playwright MCP browser tools to observe the real Teams web client.

### When to reverse-engineer

- An existing endpoint returns incomplete data (e.g. empty `displayName` in the members API)
- The Teams web client clearly has a feature but no known endpoint covers it
- You suspect a separate API host or auth token is needed

### Playwright MCP tools

| Tool                                    | Purpose                                                                                                 |
| --------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| `browser_navigate`                      | Open a URL                                                                                              |
| `browser_snapshot`                      | Accessibility snapshot of the current page (**use this, not page title** — Teams titles are misleading) |
| `browser_network_requests`              | Capture all network requests since page load                                                            |
| `browser_evaluate` / `browser_run_code` | Run JS in page context (must be a **function expression**: `async (page) => { ... }`)                   |
| `browser_click`                         | Click UI elements to trigger API calls                                                                  |
| `browser_close`                         | Clean up                                                                                                |

### Workflow

#### 1. Start fresh

Close any existing browser, then open Teams in a clean session. Teams aggressively caches in IndexedDB and service workers — a stale session won't show the network requests you need.

```
browser_close
browser_navigate → https://teams.cloud.microsoft/v2/
```

On the corporate machine, Teams auto-authenticates via Entra ID SSO — no manual login needed. Verify with `browser_snapshot` that the main UI loaded (don't trust the page title).

#### 2. Capture a baseline

Call `browser_network_requests` to snapshot current requests before triggering anything.

#### 3. Trigger the target feature

Navigate to the Teams feature you're investigating (e.g. open a group chat, click the member list). Then capture requests again and diff against the baseline.

```
browser_click → (element that triggers the behavior)
browser_snapshot → (verify the data appeared in the UI)
browser_network_requests → (find new API calls)
```

#### 4. Analyze requests

For each interesting request, note:

- **Method and full URL** (including query parameters)
- **Auth header** — `Authentication: skypetoken=...` (Chat Service) vs `Authorization: Bearer ...` (middle-tier)
- **Request body** (for POST)
- **Response structure**

#### 5. Identify auth requirements

Check the request headers to determine which token is needed:

- `Authentication:` header → skype token (Chat Service API)
- `Authorization: Bearer` header → MSAL bearer token (middle-tier API)
- New/unknown audience → inspect the MSAL token cache in browser IndexedDB or sessionStorage

#### 6. Test the endpoint

Create a `.tmp_` prefixed test script to verify the endpoint works:

```typescript
import { TeamsClient } from "./src/teams-client.js";

async function main() {
  const client = await TeamsClient.create({
    email: process.env.TEAMS_EMAIL!,
    verbose: true,
  });

  const token = client.getToken();
  // token.skypeToken for Chat Service, token.bearerToken for middle-tier

  const response = await fetch("https://teams.cloud.microsoft/api/mt/...", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token.bearerToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(/* request body */),
  });

  console.log(await response.json());
}

main().catch(console.error);
```

Run: `TEAMS_EMAIL=you@example.com npx -y tsx .tmp_test-endpoint.ts`

Delete the temp script when done.

#### 7. Implement

1. Add the API function to `src/api.ts` (stateless, token as parameter)
2. Wire it into the relevant `TeamsClient` method in `src/teams-client.ts`
3. Add unit tests with mocked `fetch` in `tests/unit/`
4. Document the endpoint in `docs/findings.md`

### Gotchas

- **Page titles lie**: Teams shows "Calendar" even on the Chat view. Always use `browser_snapshot`.
- **Aggressive caching**: If you don't see expected network requests, the data is cached in IndexedDB. Use a fresh browser session.
- **IndexedDB as a fallback data source**: When you can't find the API call, read cached data via `browser_evaluate`. Stores like `Teams:profiles:{tenantId}` reveal response shapes and can help identify the originating endpoint.
- **MSAL token audiences**: Teams acquires tokens for multiple audiences (`api.spaces.skype.com`, `chatsvcagg.teams.microsoft.com`, etc.). The `authsvc/v1.0/authz` endpoint exchanges the Bearer token for a skype token — they're chained but neither reconstructs the other.
- **MRI prefixes**: People use `8:orgid:{uuid}`, bots use `28:{uuid}`.
- **Request body format varies**: Some endpoints accept raw JSON arrays (e.g. `fetchShortProfile`), others expect wrapped objects. Always check the actual body in network captures.
