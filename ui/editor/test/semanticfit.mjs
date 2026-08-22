// The rule that keeps stale colouring off the screen, tested by calling it.
//
// This one is worth testing here rather than live, because the bug it exists for is a RACE - the
// analysed text and the model briefly disagreeing - and a race is the thing a live probe is worst
// at catching on purpose. What the rule has to do, though, is not racy at all: given a text and a
// set of ranges, say whether the ranges could be ranges into that text. That is arithmetic, and
// arithmetic can simply be asked.
//
// The source is TypeScript, so it is compiled to a temporary module first.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scratch = await mkdtemp(path.join(tmpdir(), "xlide-semanticfit-"));
const out = path.join(scratch, "semanticfit.mjs");

await build({
  entryPoints: [path.join(root, "src", "semanticfit.ts")],
  outfile: out,
  bundle: true,
  format: "esm",
  platform: "neutral",
});

const { tokensFitTheText } = await import(pathToFileURL(out).href);

let passed = 0;
let failed = 0;

const check = (name, got, want = true) => {
  try {
    assert.deepEqual(got, want);
    passed++;
    console.log(`ok   ${name}`);
  } catch {
    failed++;
    console.log(`FAIL ${name}\n       got  ${JSON.stringify(got)}\n       want ${JSON.stringify(want)}`);
  }
};

/** A text model, as much of one as the rule asks for. */
const textOf = (source) => {
  const lines = source.split("\n");
  return {
    lines,
    positionAt(offset) {
      let left = offset;
      for (let at = 0; at < lines.length; at += 1) {
        if (left <= lines[at].length) {
          return { lineNumber: at + 1, column: left + 1 };
        }
        left -= lines[at].length + 1;
      }
      return { lineNumber: lines.length, column: lines[lines.length - 1].length + 1 };
    },
    lineAt(lineNumber) {
      return lines[lineNumber - 1] ?? "";
    },
  };
};

const source = [
  "Option Explicit",
  "",
  "' Puts the dialog's button on the sheet. Idempotent on purpose: the old one",
  "Public Sub InstallSheetButton()",
  "    Dim ws As Worksheet",
  "    Set ws = ThisWorkbook.Worksheets(1)",
  "End Sub",
].join("\n");

const text = textOf(source);
const offsetOf = (word) => source.indexOf(word);
const rangeOf = (word) => ({ start: offsetOf(word), end: offsetOf(word) + word.length });

// ---- what a correct set looks like ----------------------------------------------------------

check("a set covering whole identifiers fits",
  tokensFitTheText(text, [rangeOf("Worksheet"), rangeOf("ThisWorkbook"), rangeOf("InstallSheetButton")]));

check("a token against a dot on both sides fits",
  tokensFitTheText(text, [rangeOf("Worksheets")]));

check("a token at the very start of a line fits", tokensFitTheText(text, [rangeOf("Option")]));

check("an empty set fits, having nothing to disagree with", tokensFitTheText(text, []));

// ---- and what the bug looked like -------------------------------------------------------------
//
// The measured case: a comment painted in fragments, the range starting inside "Idempotent" and
// ending inside it. Offsets like these are what a set analysed against a shorter text produces.

const inside = offsetOf("Idempotent") + 5;
check("a range starting inside a word is refused",
  tokensFitTheText(text, [{ start: inside, end: inside + 6 }]), false);

check("a range ending inside a word is refused",
  tokensFitTheText(text, [{ start: offsetOf("Idempotent"), end: offsetOf("Idempotent") + 5 }]), false);

// THE WHOLE SET GOES, not the offending token: the offsets all come from one analysis, so one
// that cannot be right means none of them can be trusted.
check("one bad range condemns the set it arrived in",
  tokensFitTheText(text, [rangeOf("Worksheet"), { start: inside, end: inside + 6 }]), false);

// ---- the shift that actually causes it ---------------------------------------------------------
//
// Every offset moved by the length of an edit above it - which is exactly what happens when the
// analysis describes the text as it was before a line was typed into it.

const shifted = [rangeOf("Worksheet"), rangeOf("ThisWorkbook"), rangeOf("Dim")]
  .map((one) => ({ start: one.start + 3, end: one.end + 3 }));
check("a set shifted by an edit above it is refused", tokensFitTheText(text, shifted), false);

// ---- and it samples rather than walking -------------------------------------------------------

const many = [];
for (let at = 0; at < 5000; at += 1) {
  many.push(rangeOf("Worksheet"));
}

// THE TRADE, STATED RATHER THAN HIDDEN. A lone bad token BETWEEN two samples is not looked at,
// so it gets painted. That is the right side of the trade and not a free lunch: what this guard
// is for is a set measured against the WRONG TEXT, and such a set is wrong throughout, so the
// first sample finds it. A set with exactly one bad token in it is not a stale set - it is an
// analyser bug, and it belongs upstream rather than behind a full walk of every colouring pass
// on a 64,802-line module.
const oneRotten = [...many];
oneRotten[1] = { start: inside, end: inside + 6 };
check("a lone bad token between two samples is painted, which is the trade",
  tokensFitTheText(text, oneRotten, 4), true);

// AND THE WORD HAS TO SIT MID-LINE FOR THIS TO SAY ANYTHING. A range that maps across a row is
// skipped - a token is single-line by construction, so one that is not cannot be judged - which
// means a set built from a word ENDING its line answers "fits" no matter how far it is shifted.
// The first draft of this check did exactly that and passed while proving nothing.
const midLine = rangeOf("ThisWorkbook");
const shiftedThroughout = [];
for (let at = 0; at < 5000; at += 1) {
  shiftedThroughout.push({ start: midLine.start + 3, end: midLine.end + 3 });
}
check("but a set that is wrong throughout - which is what a stale analysis gives - is caught at once",
  tokensFitTheText(text, shiftedThroughout, 4), false);

await rm(scratch, { recursive: true, force: true });

console.log(`\n${passed}/${passed + failed} passed`);
process.exit(failed === 0 ? 0 : 1);
