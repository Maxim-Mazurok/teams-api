/**
 * Debug telemetry for teams-api.
 *
 * Intended for developer use only — not a user-facing feature.
 * Logs complete tool inputs, raw API results, formatted outputs, and errors
 * to a local JSONL file so you can replay, inspect, and debug any interaction.
 *
 * OFF by default. Enable with the environment variable:
 *   TEAMS_TELEMETRY=true
 *
 * Storage paths (directory created automatically):
 *   macOS:   ~/Library/Application Support/teams-api/telemetry.jsonl
 *   Linux:   $XDG_DATA_HOME/teams-api/telemetry.jsonl  (default: ~/.local/share/teams-api/telemetry.jsonl)
 *   Windows: %APPDATA%\teams-api\telemetry.jsonl
 *
 * Override the output path entirely with TEAMS_TELEMETRY_PATH (useful for testing).
 *
 * Each line in the file is a self-contained JSON object (JSONL).
 * Clear the file manually to start fresh.
 */

import { appendFileSync, existsSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { homedir } from "node:os";
import { randomUUID } from "node:crypto";
import type { Platform } from "./platform.js";
import { detectPlatform } from "./platform.js";

// ── Session ID ───────────────────────────────────────────────────────

let _sessionId: string | null = null;

/** Returns a stable session ID for this process lifetime (random UUID). */
export function getSessionId(): string {
  if (!_sessionId) {
    _sessionId = randomUUID();
  }
  return _sessionId;
}

// ── Storage ──────────────────────────────────────────────────────────

/** Returns the platform-appropriate directory for telemetry data. */
export function getTelemetryDir(platform?: Platform): string {
  const p = platform ?? detectPlatform();
  switch (p) {
    case "macos":
      return join(homedir(), "Library", "Application Support", "teams-api");
    case "windows": {
      const appData =
        process.env.APPDATA ?? join(homedir(), "AppData", "Roaming");
      return join(appData, "teams-api");
    }
    case "linux": {
      const xdgData =
        process.env.XDG_DATA_HOME ?? join(homedir(), ".local", "share");
      return join(xdgData, "teams-api");
    }
  }
}

/** Returns the full path to the telemetry JSONL file. */
export function getTelemetryPath(platform?: Platform): string {
  if (process.env.TEAMS_TELEMETRY_PATH) {
    return process.env.TEAMS_TELEMETRY_PATH;
  }
  return join(getTelemetryDir(platform), "telemetry.jsonl");
}

// ── Enable / disable ─────────────────────────────────────────────────

/** Returns true when telemetry is enabled via TEAMS_TELEMETRY=true. */
export function isTelemetryEnabled(): boolean {
  return process.env.TEAMS_TELEMETRY === "true";
}

// ── Write ─────────────────────────────────────────────────────────────

/**
 * Append a single JSON record to the telemetry file.
 * Silently swallowed on failure — telemetry must never break the tool.
 */
function write(record: Record<string, unknown>): void {
  if (!isTelemetryEnabled()) return;
  try {
    const filePath = getTelemetryPath();
    const dir = process.env.TEAMS_TELEMETRY_PATH
      ? dirname(filePath)
      : getTelemetryDir();
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true });
    }
    appendFileSync(
      filePath,
      JSON.stringify({ session: getSessionId(), ts: Date.now(), ...record }) +
        "\n",
      "utf-8",
    );
  } catch {
    // Telemetry failures must never surface to the user
  }
}

// ── Public API ────────────────────────────────────────────────────────

/**
 * Record a completed tool call with its full input, raw API result,
 * formatted output string, and timing.
 */
export function recordToolCall(opts: {
  tool: string;
  format: string;
  parameters: Record<string, unknown>;
  result: unknown;
  output: string;
  durationMs: number;
}): void {
  write({
    event: "tool_call",
    tool: opts.tool,
    format: opts.format,
    durationMs: opts.durationMs,
    parameters: opts.parameters,
    result: opts.result,
    output: opts.output,
  });
}

/**
 * Record a failed tool call with the full error.
 */
export function recordToolError(opts: {
  tool: string;
  format: string;
  parameters: Record<string, unknown>;
  error: unknown;
  durationMs: number;
}): void {
  write({
    event: "tool_error",
    tool: opts.tool,
    format: opts.format,
    durationMs: opts.durationMs,
    parameters: opts.parameters,
    error:
      opts.error instanceof Error
        ? { name: opts.error.name, message: opts.error.message, stack: opts.error.stack }
        : String(opts.error),
  });
}

/**
 * Record an authentication attempt.
 */
export function recordAuth(opts: {
  strategy: "auto" | "login" | "debug" | "token";
  success: boolean;
  error?: unknown;
}): void {
  write({
    event: "auth",
    strategy: opts.strategy,
    success: opts.success,
    ...(opts.error != null && {
      error:
        opts.error instanceof Error
          ? { name: opts.error.name, message: opts.error.message }
          : String(opts.error),
    }),
  });
}
