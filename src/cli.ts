#!/usr/bin/env npx tsx
/**
 * teams-api CLI
 *
 * Thin adapter that maps unified action definitions to CLI commands.
 * All commands, parameters, descriptions, and execution logic come
 * from `src/actions.ts` — the single source of truth.
 *
 * Special commands (auth, logout) are defined here directly since
 * they handle authentication rather than Teams data operations.
 */

import { writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { Command } from "commander";
import { TeamsClient } from "./teams-client.js";
import { actions, formatOutput } from "./actions.js";
import type { ActionParameter, OutputFormat } from "./actions.js";
import type { AutoLoginOptions, ManualTokenOptions } from "./types.js";

const VALID_FORMATS: OutputFormat[] = ["json", "text", "md", "toon"];

const program = new Command();

program
  .name("teams-api")
  .description("AI-native Microsoft Teams integration CLI")
  .version("0.1.0");

// ── Shared auth options ────────────────────────────────────────────────

interface AuthFlags {
  auto?: boolean;
  email?: string;
  token?: string;
  debugPort: string;
  region: string;
}

function addAuthOptions(command: Command): Command {
  return command
    .option("--auto", "Auto-acquire token via FIDO2 passkey")
    .option("--email <email>", "Corporate email for auto login")
    .option("--token <token>", "Use an existing skype token directly")
    .option(
      "--debug-port <port>",
      "Chrome debug port for manual token capture",
      "9222",
    )
    .option("--region <region>", "API region", "apac");
}

async function createClient(flags: AuthFlags): Promise<TeamsClient> {
  if (flags.token) {
    return TeamsClient.fromToken(flags.token, flags.region);
  }

  if (flags.auto) {
    if (!flags.email) {
      console.error("Error: --email is required when using --auto");
      process.exit(1);
    }
    const autoLoginOptions: AutoLoginOptions = {
      email: flags.email,
      headless: true,
      verbose: true,
    };
    return TeamsClient.create(autoLoginOptions);
  }

  const manualOptions: ManualTokenOptions = {
    debugPort: Number(flags.debugPort),
  };
  return TeamsClient.fromDebugSession(manualOptions);
}

// ── Helpers ────────────────────────────────────────────────────────────

function camelToKebab(name: string): string {
  return name.replace(/([A-Z])/g, "-$1").toLowerCase();
}

function coerceParameter(
  value: string | boolean | undefined,
  type: ActionParameter["type"],
): unknown {
  if (value === undefined) return undefined;
  switch (type) {
    case "number":
      return Number(value);
    case "boolean":
      return value === true || value === "true";
    default:
      return value;
  }
}

// ── Register actions as CLI commands ──────────────────────────────────

for (const action of actions) {
  const command = new Command(action.name).description(action.description);

  addAuthOptions(command);

  for (const parameter of action.parameters) {
    const flag = camelToKebab(parameter.name);

    if (parameter.type === "boolean") {
      if (parameter.default === true) {
        // Default-true boolean: define --no-* to allow opting out.
        // Commander auto-sets the value to true when --no-* is absent.
        command.option(`--no-${flag}`, `Disable: ${parameter.description}`);
      } else {
        command.option(`--${flag}`, parameter.description);
      }
    } else if (parameter.required) {
      command.requiredOption(`--${flag} <value>`, parameter.description);
    } else {
      command.option(
        `--${flag} <value>`,
        parameter.description,
        parameter.default !== undefined ? String(parameter.default) : undefined,
      );
    }
  }

  command.option("--format <format>", "Output format (json, text, md, toon)");
  command.option(
    "--output <file>",
    "Export output to file (default format: md)",
  );

  command.action(async (flags: Record<string, unknown>) => {
    const client = await createClient(flags as unknown as AuthFlags);

    const actionParameters: Record<string, unknown> = {};
    for (const parameter of action.parameters) {
      actionParameters[parameter.name] = coerceParameter(
        flags[parameter.name] as string | boolean | undefined,
        parameter.type,
      );
    }

    // Inject progress callback for message fetching
    if (action.name === "get-messages") {
      actionParameters.onProgress = (count: number) =>
        process.stderr.write(`\r  ${count} messages fetched...`);
    }

    // Determine output format
    const rawFormat = flags.format as string | undefined;
    if (rawFormat && !VALID_FORMATS.includes(rawFormat as OutputFormat)) {
      console.error(
        `Error: Invalid format "${rawFormat}". Valid formats: ${VALID_FORMATS.join(", ")}`,
      );
      process.exit(1);
    }
    let outputFormat: OutputFormat;
    if (rawFormat) {
      outputFormat = rawFormat as OutputFormat;
    } else if (flags.output) {
      outputFormat = "md";
    } else {
      outputFormat = "text";
    }

    try {
      const result = await action.execute(client, actionParameters);

      if (action.name === "get-messages") {
        process.stderr.write("\n");
      }

      const output = formatOutput(action, result, outputFormat);

      if (flags.output) {
        const outputPath = resolve(flags.output as string);
        writeFileSync(outputPath, output, "utf-8");
        console.log(`Output written to ${outputPath}`);
      } else {
        console.log(output);
      }
    } catch (error) {
      if (action.name === "get-messages") {
        process.stderr.write("\n");
      }
      console.error(`Error: ${(error as Error).message}`);
      process.exit(1);
    }
  });

  program.addCommand(command);
}

// ── auth command (special — creates a client, not a data operation) ───

addAuthOptions(
  program
    .command("auth")
    .description("Acquire a Teams token and print it to stdout"),
).action(async (flags: AuthFlags) => {
  const client = await createClient(flags);
  const token = client.getToken();
  console.log(JSON.stringify(token, null, 2));
});

// ── logout command (special — clears cached token) ────────────────────

program
  .command("logout")
  .description("Clear cached token from the macOS Keychain")
  .requiredOption("--email <email>", "Email whose cached token to clear")
  .action((flags: { email: string }) => {
    TeamsClient.clearCachedToken(flags.email);
    console.log(`Cached token for ${flags.email} cleared.`);
  });

program.parseAsync().catch((error: Error) => {
  console.error("Fatal error:", error.message);
  process.exit(1);
});
