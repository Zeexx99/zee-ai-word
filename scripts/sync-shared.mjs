// Keeps the host-agnostic modules in src/shared identical across the Zee AI add-ins.
//
//   node scripts/sync-shared.mjs D:\ADD_IN_CLAUDE --check   -> lists files that differ (exit code 1 if any)
//   node scripts/sync-shared.mjs D:\ADD_IN_CLAUDE           -> copies src/shared/* into the other project
//
// A fix in the Gemini client, Google sign-in, SSE parser or hash then needs to be made once and synced.
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = resolve(fileURLToPath(new URL("..", import.meta.url)), "src", "shared");
const [target, flag] = process.argv.slice(2);
if (!target) {
  console.error("Usage: node scripts/sync-shared.mjs <other-project-folder> [--check]");
  process.exit(2);
}
const dest = resolve(target, "src", "shared");
const files = readdirSync(here).filter((f) => /\.(ts|tsx|md)$/.test(f));
const norm = (s) => s.replace(/\r\n/g, "\n");
const differ = files.filter((f) => !existsSync(join(dest, f)) || norm(readFileSync(join(dest, f), "utf8")) !== norm(readFileSync(join(here, f), "utf8")));

if (flag === "--check") {
  if (!differ.length) console.log(`src/shared is in sync with ${dest}`);
  else console.log(`Out of sync with ${dest}:\n${differ.map((f) => `  - ${f}`).join("\n")}`);
  process.exit(differ.length ? 1 : 0);
}
mkdirSync(dest, { recursive: true });
for (const f of differ) writeFileSync(join(dest, f), readFileSync(join(here, f)));
console.log(differ.length ? `Copied ${differ.length} file(s) to ${dest}: ${differ.join(", ")}` : "Nothing to copy – already in sync.");
