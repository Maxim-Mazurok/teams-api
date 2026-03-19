/**
 * Post-build script: ensure compiled entry-point files have the correct
 * `#!/usr/bin/env node` shebang so they are executable via npx / npm bin.
 *
 * The source .ts files use `#!/usr/bin/env npx tsx` which gets compiled
 * into the JS output. This script replaces it with the correct node shebang.
 */

const { readFileSync, writeFileSync, chmodSync } = require("node:fs");
const { resolve } = require("node:path");

const NODE_SHEBANG = "#!/usr/bin/env node\n";
const entryPoints = ["dist/cli.js", "dist/mcp-server.js"];

for (const relativePath of entryPoints) {
  const filePath = resolve(__dirname, "..", relativePath);
  let content = readFileSync(filePath, "utf-8");

  // Remove any existing shebang line (e.g. #!/usr/bin/env npx tsx)
  if (content.startsWith("#!")) {
    content = content.slice(content.indexOf("\n") + 1);
  }

  writeFileSync(filePath, NODE_SHEBANG + content);
  chmodSync(filePath, 0o755);
}
