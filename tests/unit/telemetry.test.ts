/**
 * Unit tests for debug telemetry (src/telemetry.ts).
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { existsSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  isTelemetryEnabled,
  getTelemetryDir,
  getTelemetryPath,
  getSessionId,
  recordToolCall,
  recordToolError,
  recordAuth,
} from "../../src/telemetry.js";

// ── isTelemetryEnabled ───────────────────────────────────────────────

describe("isTelemetryEnabled", () => {
  afterEach(() => vi.unstubAllEnvs());

  it("returns false when TEAMS_TELEMETRY is not set", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "");
    expect(isTelemetryEnabled()).toBe(false);
  });

  it("returns true when TEAMS_TELEMETRY=true", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "true");
    expect(isTelemetryEnabled()).toBe(true);
  });

  it("returns false when TEAMS_TELEMETRY=1 (must be exact string 'true')", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "1");
    expect(isTelemetryEnabled()).toBe(false);
  });
});

// ── getTelemetryDir ──────────────────────────────────────────────────

describe("getTelemetryDir", () => {
  afterEach(() => vi.unstubAllEnvs());

  it("returns macOS Application Support path", () => {
    const dir = getTelemetryDir("macos");
    expect(dir).toMatch(/Library[/\\]Application Support[/\\]teams-api$/);
  });

  it("returns XDG path on linux when XDG_DATA_HOME is set", () => {
    vi.stubEnv("XDG_DATA_HOME", "/custom/data");
    expect(getTelemetryDir("linux")).toBe("/custom/data/teams-api");
  });

  it("returns ~/.local/share on linux when XDG_DATA_HOME is not set", () => {
    delete process.env.XDG_DATA_HOME;
    expect(getTelemetryDir("linux")).toMatch(/\.local[/\\]share[/\\]teams-api$/);
  });

  it("returns APPDATA path on windows when APPDATA is set", () => {
    vi.stubEnv("APPDATA", "C:\\Users\\alice\\AppData\\Roaming");
    expect(getTelemetryDir("windows")).toBe(
      join("C:\\Users\\alice\\AppData\\Roaming", "teams-api"),
    );
  });
});

// ── getTelemetryPath ─────────────────────────────────────────────────

describe("getTelemetryPath", () => {
  afterEach(() => vi.unstubAllEnvs());

  it("ends with telemetry.jsonl", () => {
    vi.stubEnv("TEAMS_TELEMETRY_PATH", "");
    expect(getTelemetryPath("macos")).toMatch(/telemetry\.jsonl$/);
  });

  it("returns TEAMS_TELEMETRY_PATH when set", () => {
    vi.stubEnv("TEAMS_TELEMETRY_PATH", "/custom/path/debug.jsonl");
    expect(getTelemetryPath()).toBe("/custom/path/debug.jsonl");
  });
});

// ── getSessionId ─────────────────────────────────────────────────────

describe("getSessionId", () => {
  it("returns a UUID-shaped string", () => {
    expect(getSessionId()).toMatch(
      /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/,
    );
  });

  it("is stable within the same process", () => {
    expect(getSessionId()).toBe(getSessionId());
  });
});

// ── write functions ───────────────────────────────────────────────────
//
// These tests write to a real temp directory by setting TEAMS_TELEMETRY_PATH
// to a temp file path so writes land there instead of the real data directory.

describe("write functions", () => {
  let tmpDir: string;
  let telemetryFile: string;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), "teams-telemetry-test-"));
    telemetryFile = join(tmpDir, "telemetry.jsonl");
    vi.stubEnv("TEAMS_TELEMETRY_PATH", telemetryFile);
  });

  afterEach(() => {
    vi.unstubAllEnvs();
    rmSync(tmpDir, { recursive: true, force: true });
  });

  function readRecords(): Record<string, unknown>[] {
    if (!existsSync(telemetryFile)) return [];
    return readFileSync(telemetryFile, "utf-8")
      .split("\n")
      .filter((l) => l.trim())
      .map((l) => JSON.parse(l) as Record<string, unknown>);
  }

  it("does not write anything when telemetry is disabled", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "");
    recordToolCall({
      tool: "get-messages",
      format: "concise",
      parameters: { conversationId: "thread1" },
      result: [],
      output: "No messages",
      durationMs: 10,
    });
    expect(existsSync(telemetryFile)).toBe(false);
  });

  it("writes a tool_call record with full inputs and output", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "true");
    recordToolCall({
      tool: "send-message",
      format: "concise",
      parameters: { conversationId: "thread1", content: "Hello" },
      result: { messageId: "123" },
      output: "Message sent",
      durationMs: 42,
    });
    const records = readRecords();
    expect(records).toHaveLength(1);
    const r = records[0];
    expect(r.event).toBe("tool_call");
    expect(r.tool).toBe("send-message");
    expect(r.format).toBe("concise");
    expect(r.durationMs).toBe(42);
    expect((r.parameters as Record<string, unknown>).content).toBe("Hello");
    expect((r.result as Record<string, unknown>).messageId).toBe("123");
    expect(r.output).toBe("Message sent");
    expect(typeof r.ts).toBe("number");
    expect(typeof r.session).toBe("string");
  });

  it("writes a tool_error record with full error details", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "true");
    const err = new Error("Request failed with 401");
    err.name = "AuthError";
    recordToolError({
      tool: "get-messages",
      format: "detailed",
      parameters: { conversationId: "thread1" },
      error: err,
      durationMs: 15,
    });
    const records = readRecords();
    expect(records).toHaveLength(1);
    const r = records[0];
    expect(r.event).toBe("tool_error");
    expect((r.error as Record<string, unknown>).name).toBe("AuthError");
    expect((r.error as Record<string, unknown>).message).toBe(
      "Request failed with 401",
    );
    expect(
      typeof (r.error as Record<string, unknown>).stack,
    ).toBe("string");
  });

  it("writes an auth record", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "true");
    recordAuth({ strategy: "auto", success: true });
    const records = readRecords();
    expect(records).toHaveLength(1);
    const r = records[0];
    expect(r.event).toBe("auth");
    expect(r.strategy).toBe("auto");
    expect(r.success).toBe(true);
    expect(r.error).toBeUndefined();
  });

  it("includes error details in auth failure record", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "true");
    recordAuth({
      strategy: "login",
      success: false,
      error: new Error("Timeout"),
    });
    const records = readRecords();
    expect(records[0].success).toBe(false);
    expect(
      ((records[0].error) as Record<string, unknown>).message,
    ).toBe("Timeout");
  });

  it("appends multiple records sequentially", () => {
    vi.stubEnv("TEAMS_TELEMETRY", "true");
    recordAuth({ strategy: "token", success: true });
    recordToolCall({
      tool: "get-messages",
      format: "concise",
      parameters: {},
      result: [],
      output: "",
      durationMs: 5,
    });
    const records = readRecords();
    expect(records).toHaveLength(2);
    expect(records[0].event).toBe("auth");
    expect(records[1].event).toBe("tool_call");
  });
});
