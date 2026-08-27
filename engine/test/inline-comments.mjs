/*
 * The two inline comment features the shared analyzer ships, held through THIS engine:
 *
 *   suppression directives   ' @xlide-analysis-disable-<scope> [codes] [-- reason]
 *   XML doc comments         ''' <summary>...</summary>, C#'s tag vocabulary
 *
 * Both are upstream's (docs/xlide_vba_analysis_suppression_comments.md and the docs/ model in
 * xlide_vscode), and both ride paths this engine could drop: the worker's suppression filter, the
 * `documentation` field on hover/completion/signature, and the suppression quick fix. The owner's
 * ground rule is that everything stays INLINE - no sidecar metadata files - and this engine ships
 * no sidecar loader, so the inline route is the whole feature here.
 *
 * Every check compares a pair, because a filter that never fires and one that always fires read
 * identically from one side.
 *
 *   node test/inline-comments.mjs
 */

import { startEngine } from "./harness.mjs";

const { call, stop } = await startEngine("inline-comments");
const CRLF = "\r\n";

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
  console.log(`  ${ok ? "ok  " : "FAIL"} ${what}${detail !== undefined ? `  -- ${String(detail).slice(0, 140)}` : ""}`);
  if (ok) { passed += 1; } else { failures.push(what); }
};

await call("initialize", {});

/** Diagnostics for one standard module standing alone in a fresh project. */
let opened = 0;
async function findingsIn(lines) {
  const source = lines.join(CRLF);
  opened += 1;
  const id = `p${opened}`;
  await call("project/open", {
    projectId: id, generation: 1,
    modules: [{ moduleName: "Probe", source, type: "standard" }],
  });
  const answer = await call("textDocument/diagnostics", {
    documentKey: `${id}/Probe`, projectId: id, generation: 1,
    source, moduleName: "Probe", moduleType: "standard",
  });
  return { source, id, diagnostics: answer.diagnostics ?? [] };
}

const codesOf = (d) => d.diagnostics.map((one) => one.code).sort().join(",") || "(none)";

/* ---- every suppression scope, against the same finding ---------------------------------------- */

console.log("suppression directives:\n");

const bare = await findingsIn([
  "Option Explicit", "",
  "Public Sub Go()", "    missing1 = 1", "End Sub",
]);
check("the bare finding is there to suppress",
  bare.diagnostics.some((d) => d.code === "undeclared-variable"), codesOf(bare));

const nextLine = await findingsIn([
  "Option Explicit", "",
  "Public Sub Go()",
  "    ' @xlide-analysis-disable-next-line undeclared-variable -- probe",
  "    missing1 = 1",
  "End Sub",
]);
check("disable-next-line suppresses it", nextLine.diagnostics.length === 0, codesOf(nextLine));

const sameLine = await findingsIn([
  "Option Explicit", "",
  "Public Sub Go()",
  "    missing1 = 1 ' @xlide-analysis-disable-line undeclared-variable",
  "End Sub",
]);
check("a trailing disable-line suppresses it", sameLine.diagnostics.length === 0, codesOf(sameLine));

const member = await findingsIn([
  "Option Explicit", "",
  "' @xlide-analysis-disable-next-member undeclared-variable",
  "Public Sub Go()", "    missing1 = 1", "End Sub", "",
  "Public Sub Other()", "    missing2 = 2", "End Sub",
]);
check("disable-next-member covers that member and not the next",
  !member.diagnostics.some((d) => d.span.start < member.source.indexOf("Other"))
    && member.diagnostics.some((d) => d.code === "undeclared-variable"),
  codesOf(member));

const file = await findingsIn([
  "' @xlide-analysis-disable-file all -- generated",
  "Option Explicit", "",
  "Public Sub Go()", "    missing1 = 1", "End Sub",
]);
check("disable-file all silences the module", file.diagnostics.length === 0, codesOf(file));

const block = await findingsIn([
  "Option Explicit", "",
  "Public Sub Go()",
  "    ' @xlide-analysis-disable-block undeclared-variable",
  "    missing1 = 1",
  "    ' @xlide-analysis-enable-block undeclared-variable",
  "    missing2 = 2",
  "End Sub",
]);
check("a block suppresses inside and not after",
  block.diagnostics.length === 1
    && block.source.slice(block.diagnostics[0]?.span.start).startsWith("missing2"),
  codesOf(block));

const unknown = await findingsIn([
  "Option Explicit", "",
  "Public Sub Go()",
  "    ' @xlide-analysis-disable-next-line bogus-code",
  "    missing1 = 1",
  "End Sub",
]);
check("an unknown code suppresses nothing and is itself a finding",
  unknown.diagnostics.some((d) => d.code === "undeclared-variable")
    && unknown.diagnostics.some((d) => d.code === "analysis-suppression-directive"),
  codesOf(unknown));

const late = await findingsIn([
  "Option Explicit", "",
  "Public Sub Go()", "End Sub", "",
  "' @xlide-analysis-disable-file all",
]);
check("a late disable-file is a directive finding, not a guess",
  late.diagnostics.some((d) => d.code === "analysis-suppression-directive"), codesOf(late));

/* ---- ''' docs through hover, signature help, and completion ----------------------------------- */

console.log("\nXML doc comments:\n");

const docSource = [
  "Option Explicit", "",
  "''' <summary>Doubles a count for the probe.</summary>",
  "''' <param name=\"n\">how many to start from</param>",
  "''' <returns>twice n</returns>",
  "Public Function Twice(ByVal n As Long) As Long",
  "    Twice = n * 2",
  "End Function", "",
  "Public Sub Caller()",
  "    Dim answer As Long",
  "    answer = Twice(21)",
  "End Sub",
].join(CRLF);

await call("project/open", {
  projectId: "docs", generation: 1,
  modules: [{ moduleName: "Documented", source: docSource, type: "standard" }],
});

const hover = await call("textDocument/hover", {
  projectId: "docs", moduleName: "Documented", source: docSource,
  offset: docSource.indexOf("Twice(21)") + 2, moduleType: "standard",
});
check("hover on a call carries the ''' summary",
  String(hover.hover?.documentation ?? "").includes("Doubles a count"),
  JSON.stringify(hover.hover?.documentation ?? null)?.slice(0, 90));

const signature = await call("textDocument/signatureHelp", {
  projectId: "docs", moduleName: "Documented", source: docSource,
  offset: docSource.indexOf("Twice(21)") + "Twice(".length, moduleType: "standard",
});
check("signature help carries the <param> doc",
  JSON.stringify(signature.signature ?? {}).includes("how many to start from"),
  JSON.stringify(signature.signature ?? null)?.slice(0, 120));

const completionSource = docSource.replace("answer = Twice(21)", "answer = Twic");
const completion = await call("textDocument/completion", {
  projectId: "docs", moduleName: "Documented", source: completionSource,
  offset: completionSource.indexOf("answer = Twic") + "answer = Twic".length,
  moduleType: "standard",
});
const twice = (completion.items ?? []).find((one) => one.label === "Twice");
check("completion's item carries the doc",
  String(twice?.documentation ?? "").includes("Doubles a count"),
  JSON.stringify(twice?.documentation ?? null)?.slice(0, 90));

/* ---- the suppression quick fix, applied and re-measured --------------------------------------- */

console.log("\nthe suppression quick fix:\n");

const fixable = await findingsIn([
  "Option Explicit", "",
  "Public Sub Go()", "    missing1 = 1", "End Sub",
]);
const finding = fixable.diagnostics.find((d) => d.code === "undeclared-variable");
const actions = finding === undefined ? { actions: [] } : await call("textDocument/codeAction", {
  projectId: fixable.id, moduleName: "Probe", source: fixable.source,
  start: finding.span.start, end: finding.span.end, moduleType: "standard",
});
const suppress = (actions.actions ?? []).filter((a) => /suppress/i.test(a.title));
check("a suppression action is offered at the finding",
  suppress.length > 0, (actions.actions ?? []).map((a) => a.title).join(" | ").slice(0, 140));

if (suppress[0]) {
  // Applied the way the page applies it - the edits, nothing else - then asked again.
  let text = fixable.source;
  for (const edit of [...suppress[0].edits].sort((a, b) => b.start - a.start)) {
    text = text.slice(0, edit.start) + edit.text + text.slice(edit.end);
  }

  await call("project/open", {
    projectId: fixable.id, generation: 2,
    modules: [{ moduleName: "Probe", source: text, type: "standard" }],
  });
  const again = await call("textDocument/diagnostics", {
    documentKey: `${fixable.id}/Probe`, projectId: fixable.id, generation: 2,
    source: text, moduleName: "Probe", moduleType: "standard",
  });
  check("applying it suppresses the finding and adds no directive finding",
    (again.diagnostics ?? []).length === 0,
    (again.diagnostics ?? []).map((d) => d.code).join(",") || "(none)");
}

await stop();
console.log(`\n${passed} passed, ${failures.length} failed`);
if (failures.length > 0) { failures.forEach((one) => console.log(`  ${one}`)); }
process.exit(failures.length > 0 ? 1 : 0);
