---
name: teams-api
description: "Use Microsoft Teams from a terminal or Node.js with teams-api. Use when an agent needs the CLI to read, search, download, send, edit, or delete Teams content, or needs the SDK for ad hoc analysis, custom automation, or application integration."
---

# Teams API

Use the CLI for straightforward Teams operations and shell automation. Use the SDK for ad hoc analysis scripts, custom control flow, or application integration when the CLI is not flexible enough.

## CLI

Run without installing globally:

```bash
npx -y -p teams-api@latest teams-api <command> [options]
```

1. Discover current operations with `npx -y -p teams-api@latest teams-api --help`.
2. Read cross-command guidance with `npx -y -p teams-api@latest teams-api guide`.
3. Inspect exact arguments with `npx -y -p teams-api@latest teams-api <command> --help`.
4. Run the command. Authentication is automatic by default; use `--login` when an interactive browser login is needed.
5. Keep the default concise output for reading. Use `--format detailed` when exact structured data is needed.

Do not ask users to paste tokens unless they explicitly choose manual token authentication. Before write operations, verify the target conversation and content. Never expose authentication output.

## SDK

Use the SDK for ad hoc scripts, repeated calls, typed results, custom filtering or aggregation, and application code:

```bash
npm install teams-api
```

Import the named `TeamsClient` export from `teams-api`. Use the package README and exported TypeScript declarations for current constructors, methods, options, and return types instead of reproducing the API surface here.
