// The folder a Source Control button offers, shortened from the middle - tested by calling it.
//
// The owner's screenshot (2026-09-10) showed "Use f:\github\xlide\xlide_vbide\artifacts\fixtures\
// scmfixture" stretching the button across the pane. The drive and the folder's own name are what
// a developer reads, so these pin that the middle gives way, on folder boundaries, and that a
// path that already fits is left exactly as it was.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-pathlabel-"));
const compiled = path.join(scratch, "pathlabel.mjs");

await build({
  entryPoints: [path.join(root, "src", "pathlabel.ts")],
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
});

const { shortenPath } = await import(pathToFileURL(compiled).href);

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

const SCREENSHOT = "f:\\github\\xlide\\xlide_vbide\\artifacts\\fixtures\\scmfixture";

check("the screenshot's path keeps its drive and first folder, and its last two", () => {
  assert.equal(shortenPath(SCREENSHOT), "f:\\github\\...\\fixtures\\scmfixture");
});

check("a path that fits comes back exactly as it was", () => {
  assert.equal(shortenPath("C:\\Repos\\Book1"), "C:\\Repos\\Book1");
  const exactly = `C:\\${"a".repeat(37)}`;
  assert.equal(exactly.length, 40);
  assert.equal(shortenPath(exactly), exactly);
});

check("the cuts land on folder separators, so only whole folder names are dropped", () => {
  const [start, end] = shortenPath(SCREENSHOT).split("...");
  assert.ok(SCREENSHOT.startsWith(start) && start.endsWith("\\"), start);
  assert.ok(SCREENSHOT.endsWith(end) && end.startsWith("\\"), end);
});

check("the end, which names the folder, has the larger share", () => {
  const [start, end] = shortenPath(SCREENSHOT).split("...");
  assert.ok(end.length > start.length, `${start} | ${end}`);
});

check("forward slashes are separators too", () => {
  assert.equal(
    shortenPath("/home/someone/projects/xlide/artifacts/fixtures/scmfixture"),
    "/home/someone/.../fixtures/scmfixture");
});

check("a name with no separator in it is cut by characters, keeping both ends", () => {
  const long = "Quarterly-close-model-for-the-finance-team-2026";
  assert.equal(shortenPath(long), `${long.slice(0, 14)}...${long.slice(-23)}`);
});

check("a last folder longer than its share keeps its own end", () => {
  const shortened = shortenPath("C:\\x\\a-very-long-final-folder-name-that-goes-on-and-on");
  assert.ok(shortened.startsWith("C:\\x\\...") && shortened.endsWith("that-goes-on-and-on"), shortened);
});

check("nothing it shortens comes back longer than was asked for", () => {
  const segments = ["C:", "Users", "someone", "OneDrive - Contoso", "Documents", "Finance", "2026", "Q3 close", "models"];
  for (let count = 2; count <= segments.length; count += 1) {
    const whole = segments.slice(0, count).join("\\");
    for (const most of [20, 30, 40, 60]) {
      const shortened = shortenPath(whole, most);
      assert.ok(shortened.length <= most, `${most}: ${shortened} (${shortened.length})`);
      assert.ok(whole.length <= most ? shortened === whole : shortened.includes("..."), shortened);
    }
  }
});

let passed = 0;
for (const { name, fn } of checks) {
  try {
    fn();
    passed += 1;
    console.log(`ok   ${name}`);
  } catch (error) {
    console.log(`FAIL ${name}\n     ${error.message}`);
  }
}

await rm(scratch, { recursive: true, force: true });
console.log(`${passed}/${checks.length} passed`);
process.exit(passed === checks.length ? 0 : 1);
