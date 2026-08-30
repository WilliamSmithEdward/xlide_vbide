/*
 * The dot menu and signature help, against real receivers of every kind.
 *
 * Run against LanguageFixture.xlsm, which exists for exactly this:
 *   tools\\New-LanguageFixture.ps1
 *   node tools\\harness\\language-features.mjs
 *
 * KNOWN-UPSTREAM CASES FAIL ON PURPOSE (the KNOWN list below, each filed on xlide_vscode).
 * They are left failing rather than deleted, because a suite that drops the cases it cannot
 * pass stops being able to tell anyone when they start passing - which is how both #11 cases
 * were caught passing the day 4.0.0 landed.
 *
 * Everything asked so far was at a DECLARATION, where zero completions is the right answer, so
 * nothing had ever tested the case a developer actually uses: type a dot after something and see
 * what it offers. Each receiver here resolves by a different path - a project class, a
 * user-defined type, an enum, and the host's own type libraries - and any one of them can be the
 * only broken one.
 *
 * Asked through the provider monaco calls, so the answer is the menu the developer would see.
 */

import { open, scratchModule, wait } from "./xlide-api.mjs";

const api = await open({});

// THE WORKBOOK HOLDING `Uses`, not whichever one happens to be active.
//
// This asked `project()`, which answers about ONE workbook: the one named, or the active one.
// With a second workbook open it got that one, looked for this fixture's modules inside it, and
// failed on "timed out waiting for Uses to be the shown module" -- which reads as the surface
// being broken rather than the suite asking the wrong workbook (2026-08-07). Two workbooks open
// at once is a designed case here, so a suite that cannot name the one it means is a suite that
// only runs when nothing else is loaded.
const project = await api.projectHolding("Uses")
  .then((row) => (row === null ? null : api.project(row.project)));

if (project === null) {
  throw new Error(
    "no open workbook holds a module named Uses; open LanguageFixture.xlsm "
    + "(tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\fixtures\\LanguageFixture.xlsm)");
}

const broken = [];
let checks = 0;
const check = (what, ok, detail) => {
  checks++;
  console.log(`  ${ok ? "ok  " : "FAIL"} ${what}${detail ? "  -- " + detail : ""}`);
  if (!ok) { broken.push(what); }
};

async function until(what, predicate, budgetMs = 15000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(200);
  }
  throw new Error(`timed out waiting for ${what}`);
}

await api.pane("open", { module: "Uses", project: project.projectId });
await until("Uses to be the shown module", async () => {
  const ui = await api.ui();
  return ui.focus.model?.toLowerCase().endsWith("/uses") ? ui : null;
});

const text = (await api.readModule("Uses", project.projectId, { live: true })).text ?? "";
const lines = text.split(/\r?\n/);

/** Completions immediately after the dot on the line holding `needle`. */
async function afterDot(needle) {
  const at = lines.findIndex((line) => line.includes(needle));
  if (at < 0) { throw new Error(`no line holding ${needle}`); }

  const column = lines[at].indexOf(".", lines[at].indexOf(needle)) + 2;
  const answer = await api.act("completions", { line: at + 1, column, trigger: "." });
  return { line: at + 1, column, items: (answer.data ?? []).map((one) => one.label), detail: answer.detail };
}

console.log("the dot menu, by receiver\n");

const gadget = await afterDot("g.Spin");
console.log(`  g.        (${gadget.line}:${gadget.column}) ${gadget.items.length}: ${JSON.stringify(gadget.items)}`);
check("a class receiver offers its public members",
  ["Spin", "Describe", "Name", "Count"].every((m) => gadget.items.includes(m)),
  gadget.items.join(","));
check("a class receiver hides its private members",
  !gadget.items.includes("Reset"),
  "Reset is Private and must not be offered");

const udt = await afterDot("p.X");
console.log(`  p.        (${udt.line}:${udt.column}) ${udt.items.length}: ${JSON.stringify(udt.items)}`);
check("a user-defined type offers its fields",
  ["X", "Y", "Label"].every((m) => udt.items.includes(m)), udt.items.join(","));

const enumeration = await afterDot("Corner.TopLeft");
console.log(`  Corner.   (${enumeration.line}:${enumeration.column}) ${enumeration.items.length}: ${JSON.stringify(enumeration.items)}`);
check("an enum offers its members",
  ["TopLeft", "TopRight", "BottomLeft", "BottomRight"].every((m) => enumeration.items.includes(m)),
  enumeration.items.join(","));

const host = await afterDot("Application.Calculate");
console.log(`  Application. (${host.line}:${host.column}) ${host.items.length} items`);
check("the host object model offers members", host.items.length > 20,
  `${host.items.length} members; Calculate present: ${host.items.includes("Calculate")}`);
check("Application offers Calculate", host.items.includes("Calculate"));

const sheet = await afterDot("ActiveSheet.Calculate");
console.log(`  ActiveSheet. (${sheet.line}:${sheet.column}) ${sheet.items.length} items`);
check("ActiveSheet resolves to a Worksheet", sheet.items.includes("Range"),
  sheet.items.slice(0, 8).join(","));

/*
 * AND NOTHING IS OFFERED THAT VBA CANNOT PARSE.
 *
 * The host model carries the type library's hidden plumbing - `_Default`, `_NewEnum`,
 * `_Evaluate`, `_Run2`, `_CodeName`, 24 of them across the twenty commonest Excel types - and a
 * VBA identifier begins with a letter, so none of them can be written after a dot. The VBE says
 * so itself: `?ThisWorkbook._CodeName` is "Compile error: Syntax error" where
 * `?ThisWorkbook.CodeName` answers. Offering one is offering code that will not compile, which
 * this did until 2026-08-30 (the owner, on `_CodeName` under `ThisWorkbook.`).
 *
 * Both receivers, because they resolve through different types and the plumbing is per type.
 */
for (const [receiver, menu] of [["Application", host], ["ActiveSheet", sheet]]) {
  const untypeable = menu.items.filter((one) => String(one).startsWith("_"));
  check(`${receiver}. offers nothing VBA cannot type`, untypeable.length === 0,
    untypeable.length > 0
      ? `${untypeable.join(", ")} - accepting one inserts a line that will not compile`
      : `${menu.items.length} members, none of them hidden plumbing`);
}

console.log("\nsignature help, mid-argument:");
const callLine = lines.findIndex((l) => l.includes('g.Describe "prefix"'));
const inside = lines[callLine].indexOf('"prefix"') + 3;
const sig = await api.act("signature", { line: callLine + 1, column: inside });
console.log(`  ${JSON.stringify(sig).slice(0, 260)}`);
check("signature help answers inside a call", sig.did, sig.detail);

/*
 * THE CROSS-FORM MATRIX (#77, upstream xlide_vscode#22): EntryForm's control must be offered
 * from every module kind a project has, in both spellings - the form named directly, and a
 * variable declared As the form. Before the upstream fix all ten cells were zero: the form's
 * members answered only inside its own code-behind.
 */
console.log("\nthe cross-form matrix: EntryForm's control from every caller kind\n");

/** A module's dot-menu prober: opens its pane once, reads its live text, probes by needle. */
async function proberFor(module) {
  await api.pane("open", { module, project: project.projectId });
  await until(`${module} to be the shown module`, async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith(`/${module.toLowerCase()}`) ? ui : null;
  });

  const held = (await api.readModule(module, project.projectId, { live: true })).text ?? "";
  const own = held.split(/\r?\n/);

  return async (needle) => {
    const at = own.findIndex((line) => line.includes(needle));
    if (at < 0) { throw new Error(`no line holding ${needle} in ${module}`); }
    const column = own[at].indexOf(".", own[at].indexOf(needle)) + 2;
    const answer = await api.act("completions", { line: at + 1, column, trigger: "." });
    return (answer.data ?? []).map((one) => one.label);
  };
}

const CALLERS = [
  ["Uses", "a standard module"],
  ["Gadget", "a class module"],
  ["OtherForm", "another form"],
  ["ThisWorkbook", "the workbook module"],
  ["Sheet1", "a worksheet module"],
];

for (const [module, kind] of CALLERS) {
  const probe = await proberFor(module);

  const direct = await probe("EntryForm.NameBox");
  console.log(`  ${module.padEnd(12)} EntryForm. ${direct.length} items`);
  check(`${kind}: EntryForm. offers the control`,
    direct.includes("NameBox"), direct.slice(0, 6).join(","));
  check(`${kind}: EntryForm. offers the form surface`,
    direct.includes("Show"), `${direct.length} items`);

  const declared = await probe("f.NameBox");
  check(`${kind}: a variable As EntryForm offers the control`,
    declared.includes("NameBox"), declared.slice(0, 6).join(","));
}

// The matrix's negative half: a member the form does NOT have is a finding, because the
// VBE itself refuses to compile an unknown member on an early-bound form receiver. Filed as
// xlide_vscode#26, landed 2026-08-19: absence is provable behind an authoritative control list.
const defects = await api.problems("Defects");
const findings = defects.findings ?? [];
check("a control the form does not have is a finding",
  findings.some((one) => JSON.stringify(one).includes("NoSuchControl")),
  `${findings.length} finding(s) in Defects, none at the NoSuchControl line`);

// And the EMPTY form is vouched for too: #26 proves absence only where the host supplied the
// control list, an empty one included - and the shim folded "walked the designer, found
// nothing" into "could not read" until 2026-08-19, which silenced exactly this case while
// control-bearing forms flagged. The walk here is the real designer walk, not a seed shortcut.
const bare = `Bare${process.pid}`;
const bareUses = `BareUses${process.pid}`;
const bareScratch = scratchModule(api, project.projectId, bare);
const bareUsesScratch = scratchModule(api, project.projectId, bareUses);
try {
  await api.component("add", { kind: 3, name: bare, project: project.projectId });
  await api.component("add", { kind: "module", name: bareUses, project: project.projectId });
  await api.writeModule(bareUses,
    `Option Explicit\r\n\r\nPublic Sub Poke()\r\n    ${bare}.Nope\r\nEnd Sub\r\n`, project.projectId);
  const flagged = await until("the empty form to prove absence", async () => {
    const found = (await api.problems(bareUses)).findings ?? [];
    return found.some((one) => JSON.stringify(one).includes("Nope")) ? found : null;
  }).catch(() => null);
  check("a member on a control-less form is a finding too",
    flagged !== null,
    flagged === null
      ? "the empty form's control list never became authoritative"
      : `${flagged.length} finding(s), the Nope line among them`);
} finally {
  await bareUsesScratch.dispose();
  await bareScratch.dispose();
}

for (const one of broken) { console.log("  ! " + one); }

// Known-upstream failures are tolerated here rather than failing the run: each is filed, and
// the UPSTREAM FIXED line announces the day one starts passing (both #11 cases did with
// 4.0.0, verified 2026-08-19, and were trimmed from this list the same day).
const KNOWN = [
  // Empty since 2026-08-19: xlide_vscode#26 (a member the form does not have) landed upstream
  // and its row passes - the third trim, after both #11 cases went the same way with 4.0.0.
];
const unexpected = broken.filter((one) => !KNOWN.includes(one));
const fixed = KNOWN.filter((one) => !broken.includes(one));

if (fixed.length > 0) {
  console.log(`\n  UPSTREAM FIXED: ${fixed.join(", ")} - close the filed issue and trim the KNOWN list.`);
}
if (unexpected.length > 0) {
  console.log(`\n  ${unexpected.length} failure(s) beyond the known upstream ones.`);
}

// The gate's verdict spelling, with the tolerance folded in: the filed-upstream failures are
// not news, so they do not count against the run - they are listed above, named in the count
// here, and the UPSTREAM FIXED line announces the day the list can shrink.
const tolerated = broken.length - unexpected.length;
console.log(`\n${checks - unexpected.length} passed, ${unexpected.length} failed`
  + (tolerated > 0 ? ` (tolerating ${tolerated} known upstream, filed - see KNOWN above)` : ""));
process.exit(unexpected.length === 0 ? 0 : 1);
