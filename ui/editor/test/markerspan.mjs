// Where a diagnostic gets underlined when the analysis gave a position but no extent.
//
// Worth testing by calling rather than by looking, because the symptom is a squiggle in the wrong
// place and "the wrong place" is only visible if you happen to be looking at the right line. The
// rule itself is arithmetic: given a line's text and a range, say where the underline ends.
//
// PERMANENT, not a stopgap. Asking the analyzer for a span with width was declined
// (xlide_vscode#49, won't fix), so zero-width diagnostics keep arriving and this keeps answering
// them - which makes these checks the specification rather than a holding pattern.
//
// The source is TypeScript, so it is compiled to a temporary module first.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scratch = await mkdtemp(path.join(tmpdir(), "xlide-markerspan-"));
const out = path.join(scratch, "markerspan.mjs");

await build({
  entryPoints: [path.join(root, "src", "markerspan.ts")],
  outfile: out,
  bundle: true,
  format: "esm",
  platform: "neutral",
});

const { widenIfEmpty } = await import(pathToFileURL(out).href);

let passed = 0;
let failed = 0;

const check = (name, got, want) => {
  try {
    assert.deepEqual(got, want);
    passed++;
    console.log(`ok   ${name}`);
  } catch {
    failed++;
    console.log(`FAIL ${name}\n       got  ${JSON.stringify(got)}\n       want ${JSON.stringify(want)}`);
  }
};

/** A model, as much of one as the rule asks for. */
const textOf = (lines) => ({
  lineCount: () => lines.length,
  lastNonWhitespaceColumn: (line) => {
    const text = lines[line - 1] ?? "";
    const trimmed = text.replace(/\s+$/, "");
    return trimmed.length === 0 ? 0 : trimmed.length + 1;
  },
});

// The module from the report: thirteen lines, and the finding sat at (1,1) with no width.
const text = textOf([
  "Sub T1_BracketedRead()",
  "    Dim v As Variant",
  "    v = [foo]",
  "End Sub",
  "",
  "Sub T2_BracketedWrite()",
  "    [foo] = 1",
  "End Sub",
  "",
  "Sub T3_PlainUndeclared()",
  "    Dim v As Variant",
  "    v = foo",
  "End Sub",
]);

// ---- a range with width is the analysis's to decide -------------------------------------------

check("a real range is returned untouched",
  widenIfEmpty(text, { startLine: 3, startColumn: 9, endLine: 3, endColumn: 14 }),
  { endLine: 3, endColumn: 14 });

check("and so is one that spans rows",
  widenIfEmpty(text, { startLine: 1, startColumn: 1, endLine: 4, endColumn: 8 }),
  { endLine: 4, endColumn: 8 });

// ---- and one without gets the line it names ---------------------------------------------------
//
// This is the reported case. Left alone it rendered on line 13, under `End Sub`, while the
// Problems pane said (1, 1).

check("an empty range at the top of the module underlines the first line's text",
  widenIfEmpty(text, { startLine: 1, startColumn: 1, endLine: 1, endColumn: 1 }),
  { endLine: 1, endColumn: "Sub T1_BracketedRead()".length + 1 });

check("an empty range mid-module underlines that line, not another",
  widenIfEmpty(text, { startLine: 7, startColumn: 5, endLine: 7, endColumn: 5 }),
  { endLine: 7, endColumn: "    [foo] = 1".length + 1 });

// A blank line has nothing to underline. One column is the honest answer, and it is the only case
// where a single character is right.
check("an empty range on a blank line takes one column",
  widenIfEmpty(text, { startLine: 5, startColumn: 1, endLine: 5, endColumn: 1 }),
  { endLine: 5, endColumn: 2 });

// A position PAST the line's text - the end of a line - cannot widen backwards.
check("an empty range past the end of the text takes one column",
  widenIfEmpty(text, { startLine: 4, startColumn: 40, endLine: 4, endColumn: 40 }),
  { endLine: 4, endColumn: 41 });

// ---- and it never names a line the model does not have ----------------------------------------
//
// A finding can outlive the text it was made against by a moment - the analysis is asynchronous -
// so a line number past the end has to clamp rather than produce a range nothing can draw.

check("a line past the end of the module clamps to the last one",
  widenIfEmpty(text, { startLine: 99, startColumn: 1, endLine: 99, endColumn: 1 }),
  { endLine: 13, endColumn: "End Sub".length + 1 });

check("and an empty model does not produce line zero",
  widenIfEmpty(textOf([]), { startLine: 1, startColumn: 1, endLine: 1, endColumn: 1 }),
  { endLine: 1, endColumn: 2 });

await rm(scratch, { recursive: true, force: true });

console.log(`\n${passed}/${passed + failed} passed`);
process.exit(failed === 0 ? 0 : 1);
