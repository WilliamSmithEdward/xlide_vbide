// Structural check on dist/. It does not launch a browser; it asserts that the build produced
// a self-contained, relatively linked set of files.

import assert from "node:assert/strict";
import { readFile, stat } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

async function fileExists(file) {
  try {
    return (await stat(file)).isFile();
  } catch {
    return false;
  }
}

check("dist/index.html exists", async () => {
  assert.ok(await fileExists(path.join(dist, "index.html")), "dist/index.html is missing");
});

check("every src= and href= in index.html resolves inside dist", async () => {
  const html = await readFile(path.join(dist, "index.html"), "utf8");
  const refs = [...html.matchAll(/(?:src|href)\s*=\s*"([^"]+)"/g)].map((m) => m[1]);
  assert.ok(refs.length > 0, "index.html references no assets");
  for (const ref of refs) {
    assert.ok(
      !/^[a-z]+:|^\/\//i.test(ref) && !ref.startsWith("/"),
      `reference is not relative: ${ref}`,
    );
    const target = path.join(dist, ref);
    assert.ok(await fileExists(target), `referenced asset missing from dist: ${ref}`);
  }
  console.log(`    references: ${refs.join(", ")}`);
});

check("bundle contains the xlide-dark theme and the vba language id", async () => {
  const bundle = await readFile(path.join(dist, "editor.js"), "utf8");
  assert.ok(bundle.includes("xlide-dark"), 'bundle does not contain "xlide-dark"');
  assert.ok(bundle.includes("xlide-light"), 'bundle does not contain "xlide-light"');
  assert.ok(bundle.includes("vba"), 'bundle does not contain "vba"');
});

check("the semantic legend carries every type the engine emits, append-only", async () => {
  // The provider drops any token whose type is not in this list, silently - so an engine that
  // starts emitting a type the page never learned paints nothing and fails nowhere. The legend
  // is APPEND-ONLY so existing indices keep their meaning: `function` is the call colouring
  // (xlide_vscode#20) and `enumMember` is the host-constant tint (xlide_vscode#35), each added
  // at the end in its day. The theme must map every entry, in both palettes, or that entry
  // paints the default foreground.
  const bundle = await readFile(path.join(dist, "editor.js"), "utf8");
  assert.ok(
    bundle.includes('["class","enum","struct","type","variable","function","enumMember"]')
    || bundle.includes('["class", "enum", "struct", "type", "variable", "function", "enumMember"]'),
    "the bundled legend is not the six known types plus enumMember, in that order");
  assert.match(bundle, /token:\s*"function",\s*foreground:\s*"dcdcaa"/i,
    'the theme does not map semantic "function" to the call yellow');
  assert.match(bundle, /token:\s*"enumMember",\s*foreground:\s*"4fc1ff"/i,
    'the dark theme does not map semantic "enumMember" to the host-known blue');
  assert.match(bundle, /token:\s*"enumMember",\s*foreground:\s*"0070c1"/i,
    'the light theme does not map semantic "enumMember" to the host-known blue');
});

check("codicon font shipped and referenced relatively from the stylesheet", async () => {
  assert.ok(await fileExists(path.join(dist, "codicon.ttf")), "dist/codicon.ttf is missing");
  const css = await readFile(path.join(dist, "editor.css"), "utf8");
  assert.match(css, /url\(["']?\.\/codicon\.ttf["']?\)/, "editor.css does not reference ./codicon.ttf");
});

check("editor context menu ships opaque, contrasting surfaces in both themes (#51)", async () => {
  const css = await readFile(path.join(dist, "editor.css"), "utf8");
  const surfaces = [...css.matchAll(/\.context-view\.monaco-component\.monaco-menu-container\s*>\s*\.monaco-scrollable-element\s*\{([^}]+)\}/g)];
  assert.equal(surfaces.length, 2, "the editor context menu needs dark and light surface rules");
  assert.match(surfaces[0][1], /background-color:\s*#252526\s*!important/i);
  assert.match(surfaces[0][1], /color:\s*#f0f0f0\s*!important/i);
  assert.match(css, /@media\s*\(prefers-color-scheme:\s*light\)\s*\{/i);
  assert.match(surfaces[1][1], /background-color:\s*(?:#fff|#ffffff)\s*!important/i);
  assert.match(surfaces[1][1], /color:\s*#1b1b1f\s*!important/i);
});

check("worker is bundled and wired up with a relative url", async () => {
  const worker = path.join(dist, "editor.worker.js");
  assert.ok(await fileExists(worker), "dist/editor.worker.js is missing");
  const size = (await stat(worker)).size;
  assert.ok(size > 50_000, `worker bundle looks like a stub: ${size} bytes`);
  const bundle = await readFile(path.join(dist, "editor.js"), "utf8");
  assert.ok(
    bundle.includes("./editor.worker.js"),
    "editor.js does not reference ./editor.worker.js",
  );
});

let failed = 0;
for (const { name, fn } of checks) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (error) {
    failed += 1;
    console.log(`FAIL ${name}`);
    console.log(`     ${error.message}`);
  }
}

console.log(`${checks.length - failed}/${checks.length} passed`);
process.exitCode = failed === 0 ? 0 : 1;
