// Runs the properties in properties.ts: the formatter, the procedure scanner and the tokenizer
// over generated module text.
//
// Each property runs a hundred times here. XLIDE_PROPERTY_RUNS raises that, which is what the Fuzz
// workflow does. A failure prints the smallest input fast-check found that breaks the property;
// that input becomes a case in the reader's own suite.
//
// monaco's API entry is stubbed the way format.mjs and tokenizer.mjs stub it; the Monarch
// compiler and lexer are monaco's own files, as in tokenizer.mjs.

import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const numRuns = Number(process.env.XLIDE_PROPERTY_RUNS ?? 100);
// A failure prints its seed; XLIDE_PROPERTY_SEED replays that run.
const seed = process.env.XLIDE_PROPERTY_SEED === undefined ? undefined : Number(process.env.XLIDE_PROPERTY_SEED);

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-properties-"));
const compiled = path.join(scratch, "properties.mjs");

await build({
  entryPoints: [path.join(root, "test", "properties.ts")],
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
  alias: {
    "xlide-spec": path.join(root, "vendor", "xlide-spec"),
    // monaco's package exports do not reach its Monarch modules, so they are named by path.
    "monaco-monarch": path.join(root, "node_modules", "monaco-editor", "esm", "vs", "editor", "standalone", "common", "monarch"),
  },
  plugins: [{
    name: "stub-monaco-api",
    setup(on) {
      on.onResolve({ filter: /^monaco-editor\/editor\/editor\.api\.js$/ }, (args) => ({ path: args.path, namespace: "stub" }));
      on.onLoad({ filter: /.*/, namespace: "stub" }, () => ({
        loader: "js",
        contents: [
          "const handler = { get: (t, k) => k === Symbol.toPrimitive ? () => 0 : stub(), apply: () => stub(), construct: () => stub() };",
          "function stub() { return new Proxy(function () {}, handler); }",
          "export const languages = stub(); export const editor = stub(); export const Range = stub();",
          "export const KeyCode = stub(); export const KeyMod = stub(); export const Uri = stub();",
        ].join("\n"),
      }));
    },
  }],
});

const { properties } = await import(pathToFileURL(compiled).href);

let failures = 0;
const checks = properties(numRuns, seed);
for (const { name, run } of checks) {
  try {
    run();
    console.log(`ok   ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}`);
    console.error(`     ${error.message.split("\n").join("\n     ")}`);
  }
}

await rm(scratch, { recursive: true, force: true });

console.log(`${checks.length - failures}/${checks.length} passed, ${numRuns} runs each`);
if (failures > 0) {
  process.exitCode = 1;
}
