// The tokenizer's line-to-line state, run through monaco's own Monarch compiler and lexer.
//
// A colour is otherwise only ever checked live (colouring.mjs), one rendered span at a time, and
// what this checks is a rule about LINES: a comment ending in ` _` runs on to the next line
// (MS-VBAL 3.3.1, upstream #82), so that line is comment text however much it looks like code, and
// the run ends at the first line that does not end in ` _` or at an empty line. The tokenizer
// coloured only the first line until 2026-09-23. Monarch keeps its state between lines, which is
// what carries the comment, so the real lexer is the only honest thing to ask.
//
// monaco's compiler and lexer are plain modules with no DOM in them; the lexer is given the one
// setting it reads. The page's own tokenizer definition comes from vba.ts, with monaco's API
// stubbed the way format.mjs stubs it, because the definition is a plain object.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const monarch = path.join(root, "node_modules", "monaco-editor", "esm", "vs", "editor", "standalone", "common", "monarch");

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-tokenizer-"));
const compiled = path.join(scratch, "tokenizer.mjs");

await build({
  stdin: {
    contents: [
      `export { buildVbaMonarch } from ${JSON.stringify(path.join(root, "src", "vba.ts"))};`,
      `export { compile } from ${JSON.stringify(path.join(monarch, "monarchCompile.js"))};`,
      `export { MonarchTokenizer } from ${JSON.stringify(path.join(monarch, "monarchLexer.js"))};`,
    ].join("\n"),
    resolveDir: root,
    sourcefile: "tokenizer-entry.ts",
    loader: "ts",
  },
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
  alias: { "xlide-spec": path.join(root, "vendor", "xlide-spec") },
  plugins: [{
    name: "stub-monaco-api",
    setup(on) {
      // Only the API entry vba.ts imports; the monarch modules above resolve to monaco's own files.
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

const { buildVbaMonarch, compile, MonarchTokenizer } = await import(pathToFileURL(compiled).href);

const settings = { getValue: () => 20000, onDidChangeConfiguration: () => ({ dispose() {} }) };
const tokenizer = new MonarchTokenizer(null, null, "vba", compile("vba", buildVbaMonarch([], [])), settings);

/** Each line's token types, in order, the lexer's state carried from one line to the next. */
function tokenTypes(...lines) {
  let state = tokenizer.getInitialState();
  return lines.map((line) => {
    const result = tokenizer.tokenize(line, true, state);
    state = result.endState;
    return result.tokens.map((token) => token.type);
  });
}

/** True when every token on the line is a comment. An empty line has one token of its own. */
const allComment = (types) => types.length > 0 && types.every((type) => type.startsWith("comment"));

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

check("the line a comment carries on to through ' _' is comment text (#82)", () => {
  const [first, carried, after] = tokenTypes("' note _", "End Sub", "End Sub");
  assert.ok(allComment(first), `first: ${first}`);
  assert.ok(allComment(carried), `carried: ${carried}`);
  assert.ok(!allComment(after), `after: ${after}`);
});

check("it runs on while each line ends in ' _', and stops at the first that does not", () => {
  const types = tokenTypes("' a _", "b _", "c", "x = 1");
  assert.deepEqual(types.map(allComment), [true, true, true, false], JSON.stringify(types));
});

check("an empty line ends the run", () => {
  const types = tokenTypes("' a _", "", "End Sub");
  assert.equal(allComment(types[2]), false, JSON.stringify(types));
});

check("Rem carries the same way", () => {
  const types = tokenTypes("Rem a _", "End Sub", "End Sub");
  assert.deepEqual(types.map(allComment), [true, true, false], JSON.stringify(types));
});

check("so does a comment after code, and one on an #If line", () => {
  const afterCode = tokenTypes("x = 1 ' a _", "End Sub");
  assert.ok(allComment(afterCode[1]), JSON.stringify(afterCode));
  const onDirective = tokenTypes("#If VBA7 Then ' a _", "End Sub");
  assert.ok(allComment(onDirective[1]), JSON.stringify(onDirective));
});

// Already right before the fix, pinned so the fix cannot take them away.
check("already right: an underscore with no space before it carries nothing", () => {
  const types = tokenTypes("' a_", "End Sub");
  assert.equal(allComment(types[1]), false, JSON.stringify(types));
});

check("already right: a code line continued with ' _' goes on as code", () => {
  const types = tokenTypes("x = 1 + _", "Len(\"a\")");
  assert.equal(allComment(types[1]), false, JSON.stringify(types));
});

let failures = 0;
for (const { name, fn } of checks) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}`);
    console.error(`     ${error.message.split("\n").join("\n     ")}`);
  }
}

await rm(scratch, { recursive: true, force: true });

console.log(`${checks.length - failures}/${checks.length} passed`);
if (failures > 0) {
  process.exitCode = 1;
}
