/**
 * Audit logging for state-modifying Teams actions.
 *
 * Emits structured JSON Lines to stderr or a file for compliance and audit trails.
 * Controlled by the TEAMS_AUDIT_LOG environment variable:
 *   - "off" or unset — no audit logging (default)
 *   - "stderr" — write to stderr
 *   - "file:/path/to/audit.jsonl" — append to file
 *
 * Audit failures are silent — they must never break the tool.
 */

import { appendFileSync, existsSync, mkdirSync } from "node:fs";
import { dirname } from "node:path";
import type { AuditDestination, AuditEvent } from "./types.js";

/** Resolve the audit destination from the TEAMS_AUDIT_LOG environment variable. */
export function resolveAuditDestination(): AuditDestination {
  const value = process.env.TEAMS_AUDIT_LOG?.trim();
  if (!value || value === "off") return "off";
  if (value === "stderr") return "stderr";
  if (value.startsWith("file:")) return value as AuditDestination;
  return "off";
}

/** Emit a structured audit event to the configured destination. */
export function emitAuditEvent(event: AuditEvent): void {
  const destination = resolveAuditDestination();
  if (destination === "off") return;

  try {
    const line = JSON.stringify(event) + "\n";

    if (destination === "stderr") {
      process.stderr.write(line);
      return;
    }

    if (destination.startsWith("file:")) {
      const filePath = destination.slice("file:".length);
      const directory = dirname(filePath);
      if (!existsSync(directory)) {
        mkdirSync(directory, { recursive: true });
      }
      appendFileSync(filePath, line, "utf-8");
    }
  } catch {
    // Audit failures must never surface to the user
  }
}
