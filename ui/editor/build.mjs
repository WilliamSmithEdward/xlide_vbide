// Bundles the editor surface into dist/ with only relative asset references, so the output can
// be served from a WebView2 virtual host mapping, a static server, or a subdirectory without
// any rewriting.

import { build } from "esbuild";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { readFileSync, writeFileSync } from "node:fs";
import { access, copyFile, mkdir, readdir, rm, stat, writeFile } from "node:fs/promises";

const require = createRequire(import.meta.url);
const root = path.dirname(fileURLToPath(import.meta.url));
const dist = path.join(root, "dist");

// The product version, read from where the C# build reads it. The page shows it in the corner and
// in its help dialog, and a second copy of the number in a JSON file somewhere is a second thing to
// forget at release time.
const VERSION = (() => {
  const props = readFileSync(path.resolve(root, "../../Directory.Build.props"), "utf8");
  const match = props.match(/<Version>([^<]+)<\/Version>/);
  if (!match) throw new Error("No <Version> in Directory.Build.props, so the page cannot say which build it is.");
  return match[1].trim();
})();

// A number that goes up by one every build, so "which build are you on" has a one-word answer
// rather than a timestamp to compare digit by digit. Local to the machine and not in source
// control: it counts what this machine has built, which is exactly the question being asked
// during a dev session.
const BUILD_NUMBER = (() => {
  const file = path.join(root, ".build-number");
  let previous = 0;
  try {
    previous = Number.parseInt(readFileSync(file, "utf8").trim(), 10) || 0;
  } catch {
    // First build on this machine.
  }
  const next = previous + 1;
  writeFileSync(file, `${next}\n`);
  return next;
})();

/*
 * THE PAGE'S BLACK BOX, written by hand and kept out of the bundle on purpose.
 *
 * The console ring used to be installed by the host once the page reported itself ready. That is
 * too late to be worth much: a bundle that throws while its modules initialise never reaches
 * ready, so the ring was never created, so the one route built to say what went wrong answered
 * `{"installed": false, "lines": []}` at exactly the moment it was needed. That happened on
 * 2026-08-09 and the cause - a const read during its own temporal dead zone - was found by
 * reading source instead.
 *
 * So it lives here, ahead of everything, and catches three things the bundle cannot report about
 * itself: a synchronous throw during load, a rejected promise nobody handled, and console output
 * from before the surface existed. The JavaScript context survives a module that throws, so the
 * host can still read this ring out of a page that never finished booting.
 *
 * Deliberately plain ES5 with no imports: it must not be able to fail for any of the reasons the
 * thing it is watching can fail.
 */
const BOOT_JS = `(function () {
  if (window.__xlideConsoleInstalled) { return; }
  window.__xlideConsoleInstalled = true;

  var ring = [];
  window.__xlideConsole = ring;

  function push(line) {
    try {
      ring.push(line);
      if (ring.length > 500) { ring.shift(); }
    } catch (ignored) { /* a ring that throws would take the page with it */ }
  }

  function describe(one) {
    if (typeof one === "string") { return one; }
    if (one && one.stack) { return String(one.stack); }
    if (one && one.message) { return String(one.message); }
    try { return JSON.stringify(one); } catch (e) { return String(one); }
  }

  ["log", "info", "warn", "error", "debug"].forEach(function (level) {
    var original = console[level];
    console[level] = function () {
      try {
        var parts = [];
        for (var i = 0; i < arguments.length; i++) { parts.push(describe(arguments[i])); }
        push(level + ": " + parts.join(" "));
      } catch (ignored) { /* as above */ }
      return original.apply(console, arguments);
    };
  });

  /*
   * TOLD TO THE HOST, not left to be fetched.
   *
   * The host could read the ring above with a script call, and that is a poor way to learn a page
   * is dead: it costs a round trip inside a route that has a deadline, and the first attempt spent
   * the entire budget of the one route that most needed to answer. Pushing costs nothing and
   * arrives before anybody asks.
   *
   * chrome.webview is the WebView2 host channel. It exists from the first line of the first
   * script, independent of the bundle, which is exactly the property needed here: the bundle is
   * the thing that may never run.
   */
  function tell(text) {
    try {
      window.chrome.webview.postMessage(JSON.stringify({ type: "pageError", text: text }));
    } catch (ignored) { /* no host, or none listening yet: the ring still has it */ }
  }

  // The two that matter. A bundle that dies on load dies here, with the file and line.
  window.addEventListener("error", function (event) {
    var line = "UNCAUGHT: " + describe(event.error || event.message)
      + (event.filename ? "  at " + event.filename + ":" + event.lineno + ":" + event.colno : "");
    push(line);
    tell(line);
  });

  window.addEventListener("unhandledrejection", function (event) {
    var line = "UNHANDLED REJECTION: " + describe(event.reason);
    push(line);
    tell(line);
  });
})();
`;

// The static page is copied from index.html; no template evaluation is needed.

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
    define: {
      __XLIDE_BUILD__: JSON.stringify(new Date().toISOString().slice(0, 19)),
      __XLIDE_VERSION__: JSON.stringify(VERSION),
      __XLIDE_BUILD_NUMBER__: JSON.stringify(BUILD_NUMBER),
    },
    target: ["chrome120"],
    platform: "browser",
    minify: true,
    sourcemap: false,
    legalComments: "none",
    // Flat names so index.html can reference ./editor.js, ./editor.css and ./codicon.ttf.
    entryNames: "[name]",
    assetNames: "[name]",
    // The spec repo's smart-editing helpers, bundled from a vendored copy of their real sources so
    // that this build needs nothing outside the repository. tsc resolves the same specifier to
    // local declarations instead (tsconfig paths): behaviour from the spec, types from here, and
    // the two cannot disagree about what runs. tools/sync-spec.mjs refreshes the copy and checks it
    // against the spec repo whenever a neighbouring checkout is there to check against.
    alias: {
      "xlide-spec": path.resolve(root, "vendor/xlide-spec"),
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

  await copyFile(path.join(root, "index.html"), path.join(dist, "index.html"));
  await writeFile(path.join(dist, "boot.js"), BOOT_JS, "utf8");
  await ensureCodicon();

  const names = (await readdir(dist)).sort();
  const rows = [];
  for (const name of names) {
    const info = await stat(path.join(dist, name));
    rows.push({ name, bytes: info.size });
  }

  const width = Math.max(...rows.map((row) => row.name.length));
  console.log(`xlide ${VERSION}, build ${BUILD_NUMBER}`);
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
