/*
 * The dot menu and signature help, against real receivers of every kind.
 *
 * Run against LanguageFixture.xlsm, which exists for exactly this:
 *   tools\\New-LanguageFixture.ps1
 *   node tools\\harness\\language-features.mjs
 *
 * TWO CASES FAIL ON PURPOSE. They are analyzer defects filed upstream as xlide_vscode#11: a
 * project `Type` receiver offers its own name instead of its fields, and an `Enum` receiver
 * offers nothing. They are left failing rather than deleted, because a suite that drops the
 * cases it cannot pass stops being able to tell anyone when they start passing.
 *
 * Everything asked so far was at a DECLARATION, where zero completions is the right answer, so
 * nothing had ever tested the case a developer actually uses: type a dot after something and see
 * what it offers. Each receiver here resolves by a different path — a project class, a
 * user-defined type, an enum, and the host's own type libraries — and any one of them can be the
 * only broken one.
 *
 * Asked through the provider monaco calls, so the answer is the menu the developer would see.
 */

import { open } from "./xlide-api.mjs";

const api = await open({});
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const project = await api.project();

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

console.log("\nsignature help, mid-argument:");
const callLine = lines.findIndex((l) => l.includes('g.Describe "prefix"'));
const inside = lines[callLine].indexOf('"prefix"') + 3;
const sig = await api.act("signature", { line: callLine + 1, column: inside });
console.log(`  ${JSON.stringify(sig).slice(0, 260)}`);
check("signature help answers inside a call", sig.did, sig.detail);

console.log(`\n${checks} checks, ${broken.length} broken`);
for (const one of broken) { console.log("  ! " + one); }

// The two known-upstream failures are expected. Anything else is news, and only news fails
// this script: a known defect that has been filed should not stop a run from being useful.
const KNOWN = [
  "a user-defined type offers its fields",
  "an enum offers its members",
];
const unexpected = broken.filter((one) => !KNOWN.includes(one));
const fixed = KNOWN.filter((one) => !broken.includes(one));

if (fixed.length > 0) {
  console.log(`\n  UPSTREAM FIXED: ${fixed.join(", ")} — xlide_vscode#11 can be closed and this list trimmed.`);
}
if (unexpected.length > 0) {
  console.log(`\n  ${unexpected.length} failure(s) beyond the known upstream ones.`);
}

process.exit(unexpected.length === 0 ? 0 : 1);
