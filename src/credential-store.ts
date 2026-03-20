/**
 * Cross-platform credential storage.
 *
 * Provides a unified interface for securely storing and retrieving
 * credentials across macOS, Windows, and Linux.
 *
 * | Platform | Mechanism                                              |
 * |----------|--------------------------------------------------------|
 * | macOS    | System Keychain via `security` CLI                     |
 * | Windows  | DPAPI encryption → file in %APPDATA%/teams-api/        |
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
import { join } from "node:path";
import { createHash } from "node:crypto";

export interface CredentialStore {
  save(account: string, data: string): void;
  load(account: string): string | null;
  clear(account: string): void;
}

// ── macOS: System Keychain ────────────────────────────────────────────

const KEYCHAIN_SERVICE = "teams-api";

class KeychainStore implements CredentialStore {
  save(account: string, data: string): void {
    execFileSync("security", [
      "add-generic-password",
      "-a",
      account,
      "-s",
      KEYCHAIN_SERVICE,
      "-w",
      data,
      "-U",
    ]);
  }

  load(account: string): string | null {
    try {
      return execFileSync(
        "security",
        ["find-generic-password", "-a", account, "-s", KEYCHAIN_SERVICE, "-w"],
        { encoding: "utf-8" },
      ).trim();
    } catch {
      return null;
    }
  }

  clear(account: string): void {
    try {
      execFileSync("security", [
        "delete-generic-password",
        "-a",
        account,
        "-s",
        KEYCHAIN_SERVICE,
      ]);
    } catch {
      // Entry may not exist, that's fine
    }
  }
}

// ── Windows: DPAPI encryption → file ─────────────────────────────────

function getWindowsStorePath(): string {
  const appData =
    process.env.APPDATA ??
    join(process.env.USERPROFILE ?? "", "AppData", "Roaming");
  return join(appData, "teams-api");
}

function accountToFileName(account: string): string {
  return createHash("sha256").update(account).digest("hex") + ".dat";
}

class WinCredStore implements CredentialStore {
  private getFilePath(account: string): string {
    const dir = getWindowsStorePath();
    return join(dir, accountToFileName(account));
  }

  private ensureDir(): void {
    const dir = getWindowsStorePath();
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true });
    }
  }

  save(account: string, data: string): void {
    this.ensureDir();
    const base64Input = Buffer.from(data, "utf-8").toString("base64");

    // Use DPAPI to encrypt data tied to the current Windows user account
    const encrypted = execFileSync(
      "powershell",
      [
        "-NoProfile",
        "-Command",
        `Add-Type -AssemblyName System.Security; ` +
          `$bytes = [System.Convert]::FromBase64String('${base64Input}'); ` +
          `$encrypted = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser); ` +
          `[System.Convert]::ToBase64String($encrypted)`,
      ],
      { encoding: "utf-8" },
    ).trim();

    writeFileSync(this.getFilePath(account), encrypted, { encoding: "utf-8" });
  }

  load(account: string): string | null {
    const filePath = this.getFilePath(account);
    if (!existsSync(filePath)) {
      return null;
    }

    try {
      const encrypted = readFileSync(filePath, { encoding: "utf-8" }).trim();

      const decrypted = execFileSync(
        "powershell",
        [
          "-NoProfile",
          "-Command",
          `Add-Type -AssemblyName System.Security; ` +
            `$encrypted = [System.Convert]::FromBase64String('${encrypted}'); ` +
            `$bytes = [System.Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser); ` +
            `[System.Text.Encoding]::UTF8.GetString($bytes)`,
        ],
        { encoding: "utf-8" },
      ).trim();

      return decrypted;
    } catch {
      // Encrypted data may be corrupted or from a different user
      this.clear(account);
      return null;
    }
  }

  clear(account: string): void {
    const filePath = this.getFilePath(account);
    try {
      unlinkSync(filePath);
    } catch {
      // File may not exist, that's fine
    }
  }
}

// ── Linux: secret-tool or file-based fallback ────────────────────────

function getLinuxStorePath(): string {
  const configDir =
    process.env.XDG_CONFIG_HOME ?? join(process.env.HOME ?? "", ".config");
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

  save(account: string, data: string): void {
    if (this.useSecretTool) {
      try {
        execFileSync(
          "secret-tool",
          [
            "store",
            "--label",
            "teams-api",
            "service",
            KEYCHAIN_SERVICE,
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

  load(account: string): string | null {
    if (this.useSecretTool) {
      try {
        return execFileSync(
          "secret-tool",
          ["lookup", "service", KEYCHAIN_SERVICE, "account", account],
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

  clear(account: string): void {
    if (this.useSecretTool) {
      try {
        execFileSync("secret-tool", [
          "clear",
          "service",
          KEYCHAIN_SERVICE,
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
