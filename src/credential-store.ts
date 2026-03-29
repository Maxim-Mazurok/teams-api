/**
 * Cross-platform credential storage.
 *
 * Provides a unified interface for securely storing and retrieving
 * credentials across macOS, Windows, and Linux.
 *
 * | Platform | Mechanism                                              |
 * |----------|--------------------------------------------------------|
 * | macOS    | System Keychain via `security` CLI                     |
 * | Windows  | Windows Credential Manager via keytar (native wincred) |
 * | Linux    | `secret-tool` (libsecret) if available, else file with |
 * |          | 0o600 perms at ~/.config/teams-api/                    |
 */

import { execFileSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  writeFileSync,
  unlinkSync,
} from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";

export interface CredentialStore {
  save(account: string, data: string): Promise<void>;
  load(account: string): Promise<string | null>;
  clear(account: string): Promise<void>;
}

// ── macOS: System Keychain ────────────────────────────────────────────

const CREDENTIAL_SERVICE = "teams-api";

class KeychainStore implements CredentialStore {
  async save(account: string, data: string): Promise<void> {
    execFileSync("security", [
      "add-generic-password",
      "-a",
      account,
      "-s",
      CREDENTIAL_SERVICE,
      "-w",
      data,
      "-U",
    ]);
  }

  async load(account: string): Promise<string | null> {
    try {
      return execFileSync(
        "security",
        ["find-generic-password", "-a", account, "-s", CREDENTIAL_SERVICE, "-w"],
        { encoding: "utf-8" },
      ).trim();
    } catch {
      return null;
    }
  }

  async clear(account: string): Promise<void> {
    try {
      execFileSync("security", [
        "delete-generic-password",
        "-a",
        account,
        "-s",
        CREDENTIAL_SERVICE,
      ]);
    } catch {
      // Entry may not exist, that's fine
    }
  }
}

// ── Windows: Windows Credential Manager via keytar ───────────────────
//
// The previous implementation spawned PowerShell with an inline script that
// loaded System.Security.Cryptography.ProtectedData to call DPAPI.
// That pattern (PowerShell + Add-Type + Protect/Unprotect on binary data) is
// a textbook ransomware behavioral signature and is reliably flagged by
// Windows Defender. keytar calls the native wincred API directly from C++,
// with no PowerShell or inline scripting involved.

function getKeytar(): typeof import("keytar") {
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    return require("keytar") as typeof import("keytar");
  } catch {
    throw new Error(
      "keytar is required on Windows for credential storage. " +
        "Install it with: npm install keytar",
    );
  }
}

function accountKey(account: string): string {
  return createHash("sha256").update(account).digest("hex");
}

function accountToFileName(account: string): string {
  return createHash("sha256").update(account).digest("hex") + ".dat";
}

// Windows Credential Manager limits each credential blob to 2560 bytes.
// Teams tokens easily exceed this (~4500+ bytes for 5 JWTs).  We chunk
// large payloads across multiple credential entries.
const WINCRED_MAX_BLOB = 2560;
const CHUNK_MARKER = "CHUNKED:";

class WinCredStore implements CredentialStore {
  async save(account: string, data: string): Promise<void> {
    const key = accountKey(account);
    const kt = getKeytar();

    // Clear any previous chunks before saving
    await this.clearChunks(key, kt);

    if (data.length <= WINCRED_MAX_BLOB) {
      await kt.setPassword(CREDENTIAL_SERVICE, key, data);
      return;
    }

    const chunks: string[] = [];
    for (let i = 0; i < data.length; i += WINCRED_MAX_BLOB) {
      chunks.push(data.slice(i, i + WINCRED_MAX_BLOB));
    }

    // Primary entry stores the chunk count
    await kt.setPassword(CREDENTIAL_SERVICE, key, `${CHUNK_MARKER}${chunks.length}`);
    for (let i = 0; i < chunks.length; i++) {
      await kt.setPassword(CREDENTIAL_SERVICE, `${key}:${i}`, chunks[i]);
    }
  }

  async load(account: string): Promise<string | null> {
    const key = accountKey(account);
    const kt = getKeytar();

    const value = await kt.getPassword(CREDENTIAL_SERVICE, key);
    if (!value) return null;

    if (!value.startsWith(CHUNK_MARKER)) {
      return value; // Single-entry credential (backward compatible)
    }

    const chunkCount = parseInt(value.slice(CHUNK_MARKER.length), 10);
    const parts: string[] = [];
    for (let i = 0; i < chunkCount; i++) {
      const chunk = await kt.getPassword(CREDENTIAL_SERVICE, `${key}:${i}`);
      if (chunk === null) return null; // Corrupted — missing chunk
      parts.push(chunk);
    }
    return parts.join("");
  }

  async clear(account: string): Promise<void> {
    const key = accountKey(account);
    const kt = getKeytar();
    await this.clearChunks(key, kt);
    try {
      await kt.deletePassword(CREDENTIAL_SERVICE, key);
    } catch {
      // Entry may not exist
    }
  }

  private async clearChunks(
    key: string,
    kt: typeof import("keytar"),
  ): Promise<void> {
    try {
      const value = await kt.getPassword(CREDENTIAL_SERVICE, key);
      if (value?.startsWith(CHUNK_MARKER)) {
        const count = parseInt(value.slice(CHUNK_MARKER.length), 10);
        for (let i = 0; i < count; i++) {
          try {
            await kt.deletePassword(CREDENTIAL_SERVICE, `${key}:${i}`);
          } catch {
            // Chunk may not exist
          }
        }
      }
    } catch {
      // Primary entry may not exist
    }
  }
}

// ── Linux: secret-tool or file-based fallback ────────────────────────

function getLinuxStorePath(): string {
  const configDir =
    process.env.XDG_CONFIG_HOME ?? join(process.env.HOME ?? homedir(), ".config");
  return join(configDir, "teams-api");
}

function hasSecretTool(): boolean {
  try {
    execFileSync("which", ["secret-tool"], { stdio: "pipe" });
    return true;
  } catch {
    return false;
  }
}

class LinuxStore implements CredentialStore {
  private useSecretTool: boolean;

  constructor() {
    this.useSecretTool = hasSecretTool();
  }

  private getFilePath(account: string): string {
    const dir = getLinuxStorePath();
    return join(dir, accountToFileName(account));
  }

  private ensureDir(): void {
    const dir = getLinuxStorePath();
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true, mode: 0o700 });
    }
  }

  async save(account: string, data: string): Promise<void> {
    if (this.useSecretTool) {
      try {
        execFileSync(
          "secret-tool",
          [
            "store",
            "--label",
            "teams-api",
            "service",
            CREDENTIAL_SERVICE,
            "account",
            account,
          ],
          { input: data, encoding: "utf-8" },
        );
        return;
      } catch {
        // Fall through to file-based storage
        this.useSecretTool = false;
      }
    }

    this.ensureDir();
    writeFileSync(this.getFilePath(account), data, {
      encoding: "utf-8",
      mode: 0o600,
    });
  }

  async load(account: string): Promise<string | null> {
    if (this.useSecretTool) {
      try {
        return execFileSync(
          "secret-tool",
          ["lookup", "service", CREDENTIAL_SERVICE, "account", account],
          { encoding: "utf-8" },
        ).trim();
      } catch {
        // Secret not found or secret-tool failed
        return null;
      }
    }

    const filePath = this.getFilePath(account);
    if (!existsSync(filePath)) {
      return null;
    }

    try {
      return readFileSync(filePath, { encoding: "utf-8" }).trim();
    } catch {
      return null;
    }
  }

  async clear(account: string): Promise<void> {
    if (this.useSecretTool) {
      try {
        execFileSync("secret-tool", [
          "clear",
          "service",
          CREDENTIAL_SERVICE,
          "account",
          account,
        ]);
      } catch {
        // Entry may not exist
      }
    }

    // Also clean up the file-based fallback
    const filePath = this.getFilePath(account);
    try {
      unlinkSync(filePath);
    } catch {
      // File may not exist
    }
  }
}

// ── Factory ──────────────────────────────────────────────────────────

export function createCredentialStore(): CredentialStore {
  switch (process.platform) {
    case "darwin":
      return new KeychainStore();
    case "win32":
      return new WinCredStore();
    default:
      return new LinuxStore();
  }
}
