// Which procedure a line is in, tested by calling it.
//
// The rule is the native editor's own (CodeModule.ProcOfLine), and the folders suite holds
// the page to it line by line against a real module. These pin the rule itself, so a change to
// it fails here in milliseconds rather than in Excel.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-procedureat-"));
const compiled = path.join(scratch, "procedureat.mjs");

await build({
  entryPoints: [path.join(root, "src", "procedureat.ts")],
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
});

const { scanProcedures, procedureAt, describeProcedure } = await import(pathToFileURL(compiled).href);

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

const MODULE = [
  /* 1 */ "Option Explicit",
  /* 2 */ "",
  /* 3 */ "Private mTotal As Long",
  /* 4 */ "",
  /* 5 */ "' Adds one.",
  /* 6 */ "Public Sub Add()",
  /* 7 */ "    mTotal = mTotal + 1",
  /* 8 */ "End Sub",
  /* 9 */ "",
  /* 10 */ "' The total so far.",
  /* 11 */ "Public Property Get Total() As Long",
  /* 12 */ "    Total = mTotal",
  /* 13 */ "End Property",
  /* 14 */ "",
  /* 15 */ "Private Static Function Twice(ByVal n As Long) As Long",
  /* 16 */ "    Twice = n * 2",
  /* 17 */ "End Function",
  /* 18 */ "",
  /* 19 */ "' trailing",
].join("\r\n");

const at = (line) => describeProcedure(procedureAt(scanProcedures(MODULE), line));

check("the declarations section is nobody's", () => {
  assert.equal(at(1), "");
  assert.equal(at(2), "");
  assert.equal(at(3), "");
});

check("the comment and blank above the first procedure belong to it", () => {
  assert.equal(at(4), "Sub Add");
  assert.equal(at(5), "Sub Add");
  assert.equal(at(6), "Sub Add");
  assert.equal(at(8), "Sub Add");
});

check("the lines after an End belong to the next procedure", () => {
  assert.equal(at(9), "Property Get Total");
  assert.equal(at(10), "Property Get Total");
  assert.equal(at(13), "Property Get Total");
  assert.equal(at(14), "Function Twice");
});

check("the last procedure keeps everything after its End", () => {
  assert.equal(at(17), "Function Twice");
  assert.equal(at(18), "Function Twice");
  assert.equal(at(19), "Function Twice");
});

check("the ranges partition the module below the declarations", () => {
  const ranges = scanProcedures(MODULE);
  assert.deepEqual(ranges.map((one) => [one.start, one.header, one.end]), [[4, 6, 8], [9, 11, 13], [14, 15, 19]]);
});

check("every property leg is its own procedure with its kind spelled", () => {
  const text = "Property Let X(v)\nEnd Property\nPROPERTY SET X(v)\nEnd Property\nproperty get X()\nEnd Property";
  assert.deepEqual(scanProcedures(text).map((one) => one.kind), ["Property Let", "Property Set", "Property Get"]);
  assert.deepEqual(scanProcedures(text).map((one) => one.name), ["X", "X", "X"]);
});

check("a module with no procedures answers nothing everywhere", () => {
  assert.deepEqual(scanProcedures("Option Explicit\nDim x"), []);
  assert.equal(procedureAt([], 1), null);
});

check("an unclosed procedure runs to the end, and a header inside it starts another", () => {
  const text = "Sub A()\n  x = 1\nSub B()\nEnd Sub\nSub C()\n  y = 2";
  const ranges = scanProcedures(text);
  assert.deepEqual(ranges.map((one) => [one.name, one.start, one.end]), [["A", 1, 2], ["B", 3, 4], ["C", 5, 6]]);
});

check("a comment or a string is not a header, and Exit Sub is not an End", () => {
  const text = "' Sub NotOne()\nSub Real()\n    If x Then Exit Sub\n    s = \"End Sub\"\nEnd Sub";
  const ranges = scanProcedures(text);
  assert.deepEqual(ranges.map((one) => [one.name, one.start, one.end]), [["Real", 1, 5]]);
});

check("names are not ASCII", () => {
  assert.equal(scanProcedures("Sub Größe()\nEnd Sub")[0].name, "Größe");
});

check("any line ending counts a line", () => {
  assert.equal(describeProcedure(procedureAt(scanProcedures("Sub A()\rEnd Sub\rSub B()\rEnd Sub"), 3)), "Sub B");
});

let failures = 0;
for (const { name, fn } of checks) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}`);
    console.error(`     ${error.message}`);
  }
}

await rm(scratch, { recursive: true, force: true });

console.log(`${checks.length - failures}/${checks.length} passed`);
if (failures > 0) {
  process.exitCode = 1;
}
