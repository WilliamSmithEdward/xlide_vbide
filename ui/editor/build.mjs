// Bundles the editor surface into dist/ with only relative asset references, so the output can
// be served from a WebView2 virtual host mapping, a static server, or a subdirectory without
// any rewriting.

import { build } from "esbuild";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { access, copyFile, mkdir, readdir, rm, stat, writeFile } from "node:fs/promises";

const require = createRequire(import.meta.url);
const root = path.dirname(fileURLToPath(import.meta.url));
const dist = path.join(root, "dist");

const INDEX_HTML = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'self' blob:; style-src 'self' 'unsafe-inline'; font-src 'self'; img-src 'self' data:; worker-src 'self' blob:; connect-src 'self'">
<title>xlide editor</title>
<link rel="stylesheet" href="./editor.css">
<style>
  html, body { height: 100%; margin: 0; padding: 0; overflow: hidden; }
  body { background: #1e1e1e; }
  @media (prefers-color-scheme: light) { body { background: #ffffff; } }
  #container { position: absolute; inset: 0; }
</style>
</head>
<body>
<div id="container"></div>
<script src="./editor.js"></script>
</body>
</html>
`;

async function exists(file) {
  try {
    await access(file);
    return true;
  } catch {
    return false;
  }
}

async function ensureCodicon() {
  // The font normally lands in dist via the css url() rewrite. If the codicon feature import
  // is ever dropped the deliverable would silently vanish, so fall back to a direct copy.
  const target = path.join(dist, "codicon.ttf");
  if (await exists(target)) {
    return;
  }
  const api = require.resolve("monaco-editor/editor/editor.api.js");
  const source = path.resolve(
    path.dirname(api),
    "../base/browser/ui/codicons/codicon/codicon.ttf",
  );
  await copyFile(source, target);
}

async function main() {
  await rm(dist, { recursive: true, force: true });
  await mkdir(dist, { recursive: true });

  const common = {
    absWorkingDir: root,
    outdir: "dist",
    bundle: true,
    target: ["chrome120"],
    platform: "browser",
    minify: true,
    sourcemap: false,
    legalComments: "none",
    // Flat names so index.html can reference ./editor.js, ./editor.css and ./codicon.ttf.
    entryNames: "[name]",
    assetNames: "[name]",
    loader: {
      ".ttf": "file",
      ".svg": "file",
    },
    logLevel: "warning",
  };

  const results = await Promise.all([
    // The page bundle is a classic script: no module loader, no import.meta, so index.html can
    // load it with a plain relative <script src>.
    build({ ...common, entryPoints: { editor: "src/main.ts" }, format: "iife" }),
    // The worker bundle is a module worker, matching how Monaco constructs its own workers.
    // Monaco lazily dynamic-imports an optional diff package inside the worker; a module
    // worker parses that unconditionally, a classic worker would not be guaranteed to.
    build({ ...common, entryPoints: { "editor.worker": "src/editor.worker.ts" }, format: "esm" }),
  ]);

  await writeFile(path.join(dist, "index.html"), INDEX_HTML, "utf8");
  await ensureCodicon();

  const names = (await readdir(dist)).sort();
  const rows = [];
  for (const name of names) {
    const info = await stat(path.join(dist, name));
    rows.push({ name, bytes: info.size });
  }

  const width = Math.max(...rows.map((row) => row.name.length));
  console.log("dist:");
  for (const row of rows) {
    console.log(`  ${row.name.padEnd(width)}  ${String(row.bytes).padStart(9)} bytes`);
  }
  console.log(`  total${" ".repeat(Math.max(0, width - 5))}  ${String(rows.reduce((sum, row) => sum + row.bytes, 0)).padStart(9)} bytes`);

  const warnings = results.flatMap((result) => result.warnings);
  console.log(`warnings: ${warnings.length}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
