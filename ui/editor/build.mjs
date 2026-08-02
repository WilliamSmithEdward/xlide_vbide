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
</head>
<body>
<div id="shell">
  <div id="toolbar" role="toolbar" aria-label="Editor commands"></div>
  <div id="main">
    <div id="sidebar">
      <div id="sidebar-head">Project</div>
      <div id="sidebar-tree" role="tree" aria-label="Project explorer"></div>
    </div>
    <div id="sidebar-splitter" role="separator" aria-orientation="vertical" aria-label="Resize the project explorer" tabindex="0"></div>
    <div id="workspace">
      <div id="tabs" role="tablist" aria-label="Open modules"></div>
      <div id="container"></div>
      <div id="panel-splitter" role="separator" aria-orientation="horizontal" aria-label="Resize the panel" tabindex="0"></div>
      <div id="panel">
        <div id="panel-head">
          <button id="panel-toggle" type="button" aria-expanded="true" aria-controls="panel-body" aria-label="Show or hide the panel"></button>
          <div id="panel-tabs" role="tablist" aria-label="Panels">
            <button class="panel-tab active" data-panel="problems" role="tab" aria-selected="true" type="button">Problems</button>
            <button class="panel-tab" data-panel="immediate" role="tab" aria-selected="false" type="button">Immediate</button>
          </div>
          <span id="panel-count">no problems</span>
        </div>
        <div id="panel-body">
          <div id="panel-list" role="list"></div>
          <div id="immediate" hidden>
            <div id="immediate-log" role="log" aria-live="polite" aria-label="Immediate window output"></div>
            <div id="immediate-entry">
              <span id="immediate-prompt" aria-hidden="true">&gt;</span>
              <input id="immediate-input" type="text" spellcheck="false" autocomplete="off"
                     aria-label="Evaluate an expression or run a statement" placeholder="? Range(&quot;A1&quot;).Value">
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  <div id="status"><span id="status-position">Ln 1, Col 1</span><span id="status-module"></span><span id="status-notice" role="status" aria-live="polite"></span></div>
</div>
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
