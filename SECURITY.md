# Security

## Reporting a vulnerability

If you discover a security vulnerability in this project, please open a GitHub issue with the label `security`. Do not include exploit code or sensitive credentials in the report.

---

## Why endpoint security tools may flag this package

`teams-api` uses several techniques during authentication that can match behavioral heuristics used by endpoint security tools (e.g. Windows Defender, CrowdStrike, SentinelOne). This page explains each behavior and why it is safe.

### Chrome DevTools Protocol (CDP) network interception

During the login flow, the package opens a browser and uses the [Chrome DevTools Protocol `Fetch` domain](https://chromedevtools.github.io/devtools-protocol/tot/Fetch/) to intercept outbound requests to specific known Microsoft API hosts and extract authentication headers (`x-skypetoken`, `Authorization: Bearer`).

**Why it is safe:** Interception is scoped to a controlled set of Teams/Microsoft API hostnames. Requests are always continued unmodified — no data is altered or redirected. The browser is launched only during the explicit `teams-api auth --login` command.

**Why it may be flagged:** Network request interception via CDP resembles proxy/sniffer tooling.

### Browser automation

The package launches an Edge or Chrome browser process via [Playwright](https://playwright.dev/) to present the Teams login page.

**Why it is safe:** The browser is launched only when you run `teams-api auth --login` or `teams-api auth --auto-login`. No browser is launched during normal CLI/MCP operation after tokens are cached.

**Why it may be flagged:** Unattended browser spawning is a known indicator for fraud automation and credential theft tools.

### `localStorage` access

After the login page completes, the package reads the browser's `localStorage` to extract MSAL-cached tokens for additional Microsoft API scopes (SharePoint, image upload service).

**Why it is safe:** Access is scoped to the Teams web app origin. Only bearer tokens for Microsoft 365 services are read — no passwords, secrets, or personal data.

**Why it may be flagged:** Direct `localStorage` access via `page.evaluate()` matches patterns used by some in-browser credential theft scripts.

### Token storage

Captured tokens are stored in the operating-system credential store:

| Platform | Storage |
|----------|---------|
| macOS    | Keychain |
| Windows  | DPAPI-encrypted file at `%APPDATA%\teams-api\` |
| Linux    | `secret-tool` (libsecret) |

Tokens are short-lived (≤ 24 hours) and scoped only to the Microsoft 365 tenant you authenticated against.

---

## Windows Defender exclusion guidance

If Windows Defender or Defender for Endpoint quarantines or blocks `teams-api`, you can add exclusions to allow it.

### Option 1 — Exclusion via Windows Security settings

1. Open **Windows Security → Virus & threat protection**.
2. Under **Virus & threat protection settings**, click **Manage settings**.
3. Scroll to **Exclusions** and click **Add or remove exclusions**.
4. Add a **Folder** exclusion for:
   - `%APPDATA%\teams-api\`
5. Add a **Process** exclusion for `node.exe` (if you are using a global npm install and Defender is blocking the Node.js process itself).

### Option 2 — Submit as false positive

If you believe the detection is incorrect, submit the file to Microsoft for review:

- [Microsoft Security Intelligence — Submit a file](https://www.microsoft.com/en-us/wdsi/filesubmission)

Reference the open-source repository URL when submitting so the analyst can verify the code.

### Option 3 — Enterprise policy (Defender for Endpoint)

In a managed environment, your IT/security team can add the npm package scope or process path to the allow-list via Intune or Group Policy. Reference this `SECURITY.md` and the repository URL when requesting approval.
