// What stands above each procedure, tested by calling it.
//
// The lenses are planned from the module's text and the runner's list of tests, with nobody
// asked anything, so the plan is a pure function and is pinned here: which procedures get which
// lenses, on which line, at which column, in which order. The live suite (lenses.mjs) holds the
// drawn editor to this plan; these fail in milliseconds rather than in Excel.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scratch = await mkdtemp(path.join(tmpdir(), "xlide-lensplan-"));

const compile = async (name) => {
  const outfile = path.join(scratch, `${name}.mjs`);
  await build({
    entryPoints: [path.join(root, "src", `${name}.ts`)],
    outfile,
    bundle: true,
    format: "esm",
    platform: "node",
    logLevel: "silent",
  });
  return import(pathToFileURL(outfile).href);
};

const { planLenses, referencesTitle, RUN_TEST_TITLE, DEBUG_TEST_TITLE } = await compile("lensplan");
const { scanProcedures, headerNameColumn } = await compile("procedureat");

const plan = (text, tests = []) => {
  const lines = text.split(/\r\n|\r|\n/);
  return planLenses(scanProcedures(text), (line) => lines[line - 1] ?? "", tests);
};
const said = (lenses) => lenses.map((lens) => `${lens.line}:${lens.column} ${lens.kind} ${lens.procedure}${lens.test ? ` [${lens.test}]` : ""}`);

const checks = [];
const check = (name, body) => checks.push([name, body]);

const MODULE = [
  "Option Explicit",                          // 1
  "",                                         // 2
  "' @xlide-test",                            // 3
  "Public Sub AddsUp()",                      // 4
  "    Debug.Assert Total(1, 2) = 3",         // 5
  "End Sub",                                  // 6
  "",                                         // 7
  "Private Function Total(ByVal a As Long, ByVal b As Long) As Long", // 8
  "    Total = a + b",                        // 9
  "End Function",                             // 10
  "",                                         // 11
  "    Friend Static Property Get Name() As String", // 12
  "End Property",                             // 13
].join("\r\n");

check("every procedure gets a references lens, on its header line, at its name", () => {
  assert.deepEqual(said(plan(MODULE)), [
    "4:12 references AddsUp",
    "8:18 references Total",
    "12:32 references Name",
  ]);
});

check("a test the runner knows gets Run and Debug ahead of its count, carrying the runner's id", () => {
  assert.deepEqual(said(plan(MODULE, [{ id: "Tests.AddsUp", procedure: "AddsUp" }])), [
    "4:12 runTest AddsUp [Tests.AddsUp]",
    "4:12 debugTest AddsUp [Tests.AddsUp]",
    "4:12 references AddsUp",
    "8:18 references Total",
    "12:32 references Name",
  ]);
});

check("the runner's name matches whatever its case, as VBA's does", () => {
  const lenses = plan(MODULE, [{ id: "T.addsup", procedure: "ADDSUP" }]);
  assert.equal(lenses.filter((lens) => lens.kind === "runTest").length, 1);
  assert.equal(lenses[0].procedure, "AddsUp", "the lens names the procedure as the text spells it");
});

check("a Sub marked in the text but not on the runner's list gets no Run lens", () => {
  assert.deepEqual(plan(MODULE, []).filter((lens) => lens.kind !== "references"), []);
});

check("a test listed for a procedure this module does not have adds nothing", () => {
  assert.deepEqual(said(plan(MODULE, [{ id: "Other.Gone", procedure: "Gone" }])), said(plan(MODULE)));
});

check("only a Sub is a test: a Function or Property sharing the name is not", () => {
  const lenses = plan(MODULE, [{ id: "T.Total", procedure: "Total" }, { id: "T.Name", procedure: "Name" }]);
  assert.deepEqual(lenses.filter((lens) => lens.kind !== "references"), []);
});

check("a property's accessors each get their own count, on their own lines", () => {
  const text = [
    "Public Property Get Rate() As Double", "End Property",
    "Public Property Let Rate(ByVal value As Double)", "End Property",
  ].join("\n");
  assert.deepEqual(said(plan(text)), ["1:21 references Rate", "3:21 references Rate"]);
});

check("the column is where the name starts, whatever stands before it", () => {
  assert.equal(headerNameColumn("Sub A()"), 5);
  assert.equal(headerNameColumn("Public Sub A()"), 12);
  assert.equal(headerNameColumn("\tPrivate Static Function Répète() As Long"), 26);
  assert.equal(headerNameColumn("Property  Set   Target(ByVal value As Object)"), 17);
  assert.equal(headerNameColumn("Dim Sub1 As Long"), null);
  assert.equal(headerNameColumn("' Sub Commented()"), null);
});

check("a module with no procedures has no lenses, and neither has empty text", () => {
  assert.deepEqual(plan("Option Explicit\r\nPrivate count As Long\r\n"), []);
  assert.deepEqual(plan(""), []);
});

check("a header the line reader cannot give back gets no lens rather than a wrong column", () => {
  const procedures = scanProcedures("Sub A()\nEnd Sub\nSub B()\nEnd Sub");
  const lenses = planLenses(procedures, (line) => (line === 1 ? "Sub A()" : ""), []);
  assert.deepEqual(said(lenses), ["1:5 references A"]);
});

check("the count reads as a person would say it, and a count nobody could give still reads", () => {
  assert.equal(referencesTitle(0), "no references");
  assert.equal(referencesTitle(1), "1 reference");
  assert.equal(referencesTitle(2), "2 references");
  assert.equal(referencesTitle(137), "137 references");
  assert.equal(referencesTitle(null), "references");
  assert.equal(RUN_TEST_TITLE, "Run test");
  assert.equal(DEBUG_TEST_TITLE, "Debug test");
});

let failed = 0;
for (const [name, body] of checks) {
  try {
    body();
    console.log(`ok   ${name}`);
  } catch (error) {
    failed++;
    console.log(`FAIL ${name}\n     ${error.message.split("\n").join("\n     ")}`);
  }
}

await rm(scratch, { recursive: true, force: true });
console.log(`${checks.length - failed}/${checks.length} passed`);
process.exit(failed ? 1 : 0);
