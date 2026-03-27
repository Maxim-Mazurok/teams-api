# Security

## Reporting a vulnerability

If you discover a security vulnerability in this project, please open a GitHub issue with the label `security`. Do not include exploit code or sensitive credentials in the report.

---

## Windows Defender false positive

### What triggers it

On Windows, the previous token storage implementation spawned a PowerShell process with an inline `-Command` script that loaded `System.Security.Cryptography.ProtectedData` and called `Protect`/`Unprotect` on binary data:

```
powershell -Command "Add-Type -AssemblyName System.Security; ... ProtectedData::Protect(...) ..."
```

This is one of the most reliable behavioral signatures in Windows Defender's heuristics for ransomware and credential-theft malware — spawning PowerShell, loading a crypto assembly, and encrypting arbitrary binary data inline. Defender flags the *behavior*, not the intent.

### What we do instead

Tokens are now stored using [keytar](https://github.com/atom/node-keytar), which calls the native **Windows Credential Manager** (`wincred`) API directly from C++. No PowerShell is spawned, no inline scripts are executed, no crypto assembly is loaded. Stored credentials appear in **Control Panel → Credential Manager → Windows Credentials** under the service name `teams-api`.

### Token storage by platform

| Platform | Storage |
|----------|---------|
| macOS    | Keychain (via `security` CLI) |
| Windows  | Windows Credential Manager (via keytar / wincred) |
| Linux    | `secret-tool` (libsecret) or `~/.config/teams-api/` with 0o600 perms |

### Migrating from the old DPAPI storage

If you used an older version of `teams-api` on Windows, your tokens were stored as DPAPI-encrypted `.dat` files in `%APPDATA%\teams-api\`. These are no longer read. Run `teams-api auth --login` once to re-authenticate and store tokens in the new location. You can safely delete the old `%APPDATA%\teams-api\` directory.
