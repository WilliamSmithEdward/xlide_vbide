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
<!-- Before the stylesheet arrives, the page is already the editor's dark ground: the host holds
     the browser hidden until ready, and this keeps even a reload from ever flashing white. -->
<style>html,body{margin:0;height:100%;background:#1e1e1e}</style>
<link rel="stylesheet" href="./editor.css">
</head>
<body>
<div id="shell">
  <div id="menubar" role="menubar" aria-label="Menus"></div>
  <div id="toolbar" role="toolbar" aria-label="Editor commands"></div>
  <div id="main">
    <div id="sidebar">
      <div id="sidebar-tree" role="tree" aria-label="Project explorer"></div>
      <div id="properties-splitter" role="separator" aria-orientation="horizontal" aria-label="Resize the properties panel" tabindex="0" hidden></div>
      <div id="properties" hidden>
        <button id="properties-head" type="button" aria-expanded="true" aria-controls="properties-list">Properties</button>
        <div id="properties-object" aria-label="Selected object"></div>
        <div id="properties-list" role="list" aria-label="Properties of the selected component"></div>
      </div>
    </div>
    <div id="sidebar-splitter" role="separator" aria-orientation="vertical" aria-label="Resize the project explorer" tabindex="0"></div>
    <div id="workspace">
      <div id="tabs" role="tablist" aria-label="Open modules"></div>
      <div id="container"></div>
      <div id="empty-view" aria-hidden="true">
        <div id="empty-view-message">
          <span class="codicon codicon-files"></span>
          <p>No module is open.</p>
          <p class="empty-hint">Pick one in the explorer, or right-click a workbook to add one.</p>
        </div>
      </div>
      <div id="panel-splitter" role="separator" aria-orientation="horizontal" aria-label="Resize the panel" tabindex="0"></div>
      <div id="panel">
        <div id="panel-head">
          <button id="panel-toggle" type="button" aria-expanded="true" aria-controls="panel-body" aria-label="Show or hide the panel"></button>
          <div id="panel-tabs" role="tablist" aria-label="Panels">
            <button class="panel-tab active" data-panel="problems" role="tab" aria-selected="true" type="button">Problems</button>
            <button class="panel-tab" data-panel="immediate" role="tab" aria-selected="false" type="button">Immediate</button>
            <button class="panel-tab" data-panel="locals" role="tab" aria-selected="false" type="button">Locals</button>
            <button class="panel-tab" data-panel="search" role="tab" aria-selected="false" type="button">Search</button>
          </div>
        </div>
        <div id="problems-filters" role="toolbar" aria-label="Filter problems by severity">
          <button class="problems-filter" data-severity-filter="errors" type="button" aria-pressed="true"><span class="codicon codicon-error" aria-hidden="true"></span><span class="filter-count">0 Errors</span></button>
          <button class="problems-filter" data-severity-filter="warnings" type="button" aria-pressed="true"><span class="codicon codicon-warning" aria-hidden="true"></span><span class="filter-count">0 Warnings</span></button>
          <button class="problems-filter" data-severity-filter="messages" type="button" aria-pressed="true"><span class="codicon codicon-info" aria-hidden="true"></span><span class="filter-count">0 Messages</span></button>
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
          <div id="locals" hidden>
            <div id="locals-context" aria-live="polite" hidden></div>
            <div id="locals-table" role="table" aria-label="Local variables"></div>
          </div>          <div id="search" hidden>
            <div id="search-controls">
              <input id="search-query" type="text" spellcheck="false" autocomplete="off" placeholder="Find" aria-label="Search text">
              <input id="search-replace" type="text" spellcheck="false" autocomplete="off" placeholder="Replace" aria-label="Replacement text">
              <button id="search-case" class="search-toggle" type="button" title="Match case" aria-pressed="false">Aa</button>
              <button id="search-word" class="search-toggle" type="button" title="Whole word" aria-pressed="false">ab</button>
              <select id="search-scope" aria-label="Search scope">
                <option value="module">Module</option>
                <option value="project">Workbook</option>
                <option value="all">All workbooks</option>
              </select>
              <button id="search-run" type="button">Find All</button>
              <button id="search-replace-run" type="button">Replace All</button>
            </div>
            <div id="search-results" role="list" aria-label="Search results"></div>
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
    // Stamped into the bundle and reported in the ready message, so the host log always proves
    // WHICH build the page is running: a cached stale bundle is otherwise invisible.
    define: { __XLIDE_BUILD__: JSON.stringify(new Date().toISOString().slice(0, 19)) },
    target: ["chrome120"],
    platform: "browser",
    minify: true,
    sourcemap: false,
    legalComments: "none",
    // Flat names so index.html can reference ./editor.js, ./editor.css and ./codicon.ttf.
    entryNames: "[name]",
    assetNames: "[name]",
    // The spec repo's smart-editing helpers, bundled from their real sources. tsc resolves the
    // same specifier to local declarations instead (tsconfig paths): behaviour from the spec,
    // types from here, and the two cannot disagree about what runs.
    alias: {
      "xlide-spec": path.resolve(root, "../../../xlide_vscode/src"),
    },
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
