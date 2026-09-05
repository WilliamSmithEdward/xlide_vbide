/*
 * THE HIDDEN ATTRIBUTES, THROUGH THEIR ANNOTATIONS, against real Excel.
 *
 * A module carries attributes the code pane cannot show, and the editor offers no way to set them.
 * The annotations name them in the code; this proves the whole loop: the drift between annotation
 * and saved attribute is reported where the developer looks, the quick fix and the route write the
 * attribute by re-importing the module, the analyzer hears about a predeclared class the moment it
 * is applied, the saved package carries it after a save and the reader reads it back, and the
 * reverse direction - an attribute nothing annotates - is reported and removable.
 *
 * Runs against AttributesFixture.xlsm alone:
 *
 *   tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\AttributesFixture.xlsm
 *   node tools\harness\attributes.mjs
 *
 * It saves the workbook on purpose - that is the half that proves the package reader - and puts
 * the file back to the shape it was built in at the end, and at the start, so a rerun begins where
 * the first run did.
 */
import { open, wait, waitFor, reporter } from "./xlide-api.mjs";

const api = await open();
const { check, done } = reporter();

const BOOK = "AttributesFixture.xlsm";
const at = { project: BOOK };

const settles = (what, predicate, options) => waitFor(what, predicate, options).then(() => true, () => false);
const describe = (module) => api.attributes(module, at);
const codes = (drift) => drift.map((one) => one.code).sort().join(",");
const refusalOf = async (promise) => {
  try {
    await promise;
    return null;
  } catch (error) {
    return String(error.message ?? error);
  }
};

/** Every managed attribute the three writable modules may carry, for the reset to take away. */
const MANAGED = {
  Registry: [["ModuleDescription"], ["PredeclaredId"], ["Description", "Lookup"], ["VariableDescription", "Count"]],
  Bag: [["DefaultMember", "Item"], ["Enumerator", "NewEnum"]],
  Macros: [["ModuleDescription"], ["ExcelHotkey", "DoIt"], ["Description", "Hello"]],
};

/** Hello's description in the fixture, with its en dash and curly apostrophe: the two bytes Latin-1 read wrong. */
const HELLO_DESCRIPTION = "Prints a greeting – the day’s first.";

const carriesAny = (attributes) => attributes !== null && (
  attributes.predeclaredId === true || attributes.exposed === true || attributes.description !== null
  || attributes.members.length > 0 || attributes.variables.length > 0);

/** Takes every managed attribute off the three modules and saves, so the file is as built. */
async function resetToBuilt(label) {
  let removed = 0;
  for (const [module, removals] of Object.entries(MANAGED)) {
    const before = await describe(module);
    if (!carriesAny(before.attributes)) {
      continue;
    }
    for (const [kind, target] of removals) {
      const answer = await api.attributesRemove(module, kind, { ...at, target });
      removed += answer.changes.length;
    }
  }
  if (removed > 0) {
    await api.command("save");
    await wait(1500);
  }
  console.log(`  ${label}: ${removed} attribute(s) taken away`);
}

let registryText = null;

// A save writes annotations by default, and this suite saves to prove other things, so the
// setting is off for the run and put back at the end. The on-save behaviour gets its own section.
const originalSettings = await api.settings();
if (originalSettings.applyAttributesOnSave !== false) {
  await api.settings({ applyAttributesOnSave: false });
}

try {
  await resetToBuilt("reset at start");

  // ---- what the code says, what the file carries, and the drift ----
  const registry = await describe("Registry");
  check("Registry is a class the editor can re-import, with its attributes known from the saved file",
    registry.kind === "class" && registry.writable && registry.attributesKnown && !registry.asserted,
    JSON.stringify({ kind: registry.kind, writable: registry.writable, known: registry.attributesKnown, asserted: registry.asserted }));
  check("its four annotations are read, each bound to its target",
    registry.annotations.map((one) => `${one.kind}:${one.target ?? "-"}`).sort().join(",")
      === "Description:Lookup,ModuleDescription:-,PredeclaredId:-,VariableDescription:Count"
      && registry.problems.length === 0,
    JSON.stringify(registry.annotations));
  check("the saved class is not predeclared yet, so every annotation reads as not applied",
    registry.attributes?.predeclaredId === false && codes(registry.drift) === "annotation-not-applied,annotation-not-applied,annotation-not-applied,annotation-not-applied",
    JSON.stringify(registry.drift.map((one) => [one.code, one.line])));

  const bag = await describe("Bag");
  check("Bag's default member and enumerator are annotations with drift",
    bag.annotations.map((one) => one.kind).sort().join(",") === "DefaultMember,Enumerator" && bag.drift.length === 2,
    JSON.stringify(bag.drift));

  const misplaced = await describe("Misplaced");
  check("a predeclared standard module and a description above a variable are reported, not written",
    codes(misplaced.drift) === "annotation-not-applicable,annotation-problem"
      && misplaced.drift.find((one) => one.code === "annotation-not-applicable")?.line === 1,
    JSON.stringify(misplaced.drift.map((one) => [one.code, one.line, one.message.slice(0, 60)])));

  const sheet = await describe("Sheet1");
  check("a document module cannot take attributes, and its annotation says why",
    !sheet.writable && codes(sheet.drift) === "annotation-not-applicable" && /cannot be imported/.test(sheet.drift[0].message),
    JSON.stringify(sheet.drift));

  // ---- the Problems pane files the drift ----
  const filed = await settles("the drift in the problems list", async () => {
    const problems = await api.problems();
    const rows = problems.findings ?? problems.problems ?? [];
    return rows.some((row) => row.module === "Registry" && row.code === "annotation-not-applied")
      && rows.some((row) => row.module === "Sheet1" && row.code === "annotation-not-applicable")
      && rows.some((row) => row.module === "Misplaced" && row.code === "annotation-problem");
  }, { budgetMs: 20000 });
  check("the Problems pane carries the drift for every module, alongside the analyzer's findings", filed);

  // ---- what the analyzer is told before: the class is not predeclared ----
  // The seed reads the flag through SavedModules.PredeclaredIdOf, the same call the project
  // route answers with, and it logs what it told the engine. Both are read here, before and
  // after, rather than a finding the analyzer's rules may or may not raise on the shape.
  const predeclaredNow = async () =>
    (await api.project(BOOK)).components.find((one) => one.name === "Registry")?.predeclaredId ?? null;
  check("before applying, the seed's own source says the class is not predeclared", (await predeclaredNow()) === false);
  const logFrom = (await api.log({ max: 1 })).next;

  // ---- the quick fix and the hover, on the annotation's line ----
  await api.caret(2, { module: "Registry", project: BOOK });
  await wait(800);
  const fixes = await api.act("quickFixes", { line: 2, column: 1 });
  check("the annotation's line offers the apply as a quick fix",
    (fixes.data ?? []).some((one) => one.title === "Apply annotations to Registry's attributes now"),
    JSON.stringify(fixes.data));
  const hover = await api.act("hover", { line: 2, column: 4 });
  const hoverText = JSON.stringify(hover.data ?? {});
  check("hovering the annotation says what it writes and what the module has",
    hover.did && /VB_PredeclaredId = True/.test(hoverText) && /no VB_PredeclaredId|VB_PredeclaredId = False/.test(hoverText),
    hoverText.slice(0, 200));

  // ---- the write ----
  // The surface as the developer left it, to hold the write to: the same tabs in the same order,
  // the same module active and unfolded in the tree, the caret where it was. The first build
  // republished the tree with the module gone and reopened it at the end of the strip (the owner,
  // 2026-09-05: "it flickers the explorer, and doesn't put me back").
  const uiBefore = await api.ui();
  const tabsBefore = uiBefore.workspace.groups.flatMap((group) => group.tabs.map((tab) => tab.module));
  registryText = (await api.readModule("Registry", BOOK)).text;
  const applied = await api.attributesApply("Registry", at);
  check("applying writes every annotated attribute and names each change",
    applied.ok && applied.changes.length === 4 && applied.skipped.length === 0
      && applied.changes.some((one) => /VB_PredeclaredId False -> True/.test(one)),
    JSON.stringify(applied));
  check("the module's code is exactly what it was", (await api.readModule("Registry", BOOK)).text === registryText);
  await wait(800);
  const uiAfter = await api.ui();
  const tabsAfter = uiAfter.workspace.groups.flatMap((group) => group.tabs.map((tab) => tab.module));
  check("the tabs are the same tabs in the same order after the re-import", tabsAfter.join(",") === tabsBefore.join(","),
    `${tabsBefore.join(",")} -> ${tabsAfter.join(",")}`);
  check("the module is still the active one, unfolded in the tree, with the caret where it was",
    uiAfter.explorer.active === "Registry" && uiAfter.explorer.unfolded?.module === "Registry"
      && uiAfter.focus.line === uiBefore.focus.line,
    JSON.stringify({ active: uiAfter.explorer.active, unfolded: uiAfter.explorer.unfolded, line: [uiBefore.focus.line, uiAfter.focus.line] }));

  const after = await describe("Registry");
  check("the route answers from the applied set until the file is saved",
    after.asserted && after.attributes?.predeclaredId === true
      && after.attributes?.description === "Where things are looked up."
      && after.attributes?.members.some((one) => one.member === "Lookup" && one.description === "Finds a thing by its name.")
      && after.attributes?.variables.some((one) => one.variable === "Count" && one.description === "How many lookups so far.")
      && after.drift.length === 0,
    JSON.stringify({ asserted: after.asserted, attributes: after.attributes, drift: after.drift }));

  const seeded = await settles("the seed to tell the engine", async () => {
    const log = await api.log({ since: logFrom, max: 400, match: "seed: class Registry" });
    return (log.lines ?? []).some((line) => /seed: class Registry is predeclared/.test(line));
  }, { budgetMs: 25000 });
  check("the analyzer hears about the predeclared class before any save: the seed reads true and says so to the engine",
    seeded && (await predeclaredNow()) === true, `predeclaredId now ${await predeclaredNow()}`);

  const bagApplied = await api.attributesApply("Bag", at);
  const macrosApplied = await api.attributesApply("Macros", at);
  const macros = await describe("Macros");
  check("the default member, the enumerator, the hotkey and the descriptions all write",
    bagApplied.changes.length === 2 && macrosApplied.changes.length === 3
      && (await describe("Bag")).drift.length === 0 && macros.drift.length === 0
      && macros.attributes?.members.some((one) => one.member === "DoIt" && one.hotkey === "D"),
    JSON.stringify([bagApplied.changes, macrosApplied.changes]));
  // Hello's description holds an en dash and a curly apostrophe: bytes 0x96 and 0x92 in the page
  // the editor exports in, which a Latin-1 read turned into control characters, so the drift
  // never cleared and every save re-imported the module.
  check("a description past ASCII is written in the editor's code page and read back as the code spells it",
    macros.attributes?.members.some((one) => one.member === "Hello" && one.description === HELLO_DESCRIPTION),
    JSON.stringify(macros.attributes?.members));

  const again = await api.attributesApply("Registry", at);
  check("applying again changes nothing", again.ok && again.changes.length === 0, JSON.stringify(again));

  // ---- the save, and the package reader ----
  await api.command("save");
  await wait(2000);
  const saved = await describe("Registry");
  check("after a save the saved package carries the attributes, members included, and nothing is asserted",
    !saved.asserted && saved.attributesKnown && saved.attributes?.predeclaredId === true
      && saved.attributes?.members.some((one) => one.member === "Lookup" && one.description === "Finds a thing by its name.")
      && saved.drift.length === 0,
    JSON.stringify({ asserted: saved.asserted, attributes: saved.attributes }));
  const savedMacros = await describe("Macros");
  check("the saved package's stream is read in the project's own code page, so the non-ASCII description agrees with the code",
    !savedMacros.asserted && savedMacros.drift.length === 0
      && savedMacros.attributes?.members.some((one) => one.member === "Hello" && one.description === HELLO_DESCRIPTION),
    JSON.stringify({ asserted: savedMacros.asserted, members: savedMacros.attributes?.members, drift: savedMacros.drift }));

  // ---- the other direction: an attribute nothing annotates ----
  const withoutAnnotation = registryText.split(/\r?\n/).filter((line) => !/^'@PredeclaredId/.test(line)).join("\r\n");
  await api.writeModule("Registry", withoutAnnotation, BOOK);
  const reversed = await settles("the attribute without its annotation", async () =>
    (await describe("Registry")).drift.some((one) => one.code === "attribute-not-annotated" && one.annotation === "PredeclaredId"),
    { budgetMs: 20000 });
  check("deleting the annotation leaves the attribute, and the drift says so on line 1", reversed
    && (await describe("Registry")).drift.find((one) => one.code === "attribute-not-annotated")?.line === 1,
    JSON.stringify((await describe("Registry")).drift));
  await api.caret(1, { module: "Registry", project: BOOK });
  await wait(1500);
  const reverseFixes = await api.act("quickFixes", { line: 1, column: 1 });
  const titles = (reverseFixes.data ?? []).map((one) => one.title);
  check("its quick fixes are the annotation as a text edit and the removal",
    titles.includes("Add '@PredeclaredId above") && titles.includes("Remove VB_PredeclaredId from Registry"), titles.join(" | "));

  const removed = await api.attributesRemove("Registry", "PredeclaredId", at);
  check("removing takes the attribute away", removed.changes.length === 1 && (await describe("Registry")).attributes?.predeclaredId === false,
    JSON.stringify(removed));

  // ---- the fix from where the symptom is: a class used as a value, in another module ----
  // Registry says nothing about being predeclared now and its attribute is off, so Uses'
  // Registry.Lookup is "Variable not defined" again. The cure is an annotation in Registry plus
  // the write, and the finding in Uses offers it: on the lightbulb, and on the Problems pane's
  // right-click.
  const usesReported = await settles("Uses' undeclared Registry", async () =>
    ((await api.problems("Uses")).findings ?? []).some((one) => one.code === "undeclared-variable" && /Registry/.test(one.message)),
    { budgetMs: 25000 });
  check("with the class neither annotated nor predeclared, Registry.Lookup is reported in Uses", usesReported,
    JSON.stringify((await api.problems("Uses")).findings));
  await api.caret(8, { module: "Uses", project: BOOK, column: 17 });
  await wait(1500);
  const predeclareTitle = "Make Registry a predeclared class: add '@PredeclaredId and apply it now";
  const usesFixes = await api.act("quickFixes", { line: 8, column: 17 });
  check("the finding's lightbulb offers to make the class predeclared",
    (usesFixes.data ?? []).some((one) => one.title === predeclareTitle), (usesFixes.data ?? []).map((one) => one.title).join(" | "));
  const menuFixes = await api.act("problemFixes", { module: "Uses", workbook: BOOK, line: 8, column: 17 });
  check("the Problems pane's right-click carries the same fix", (menuFixes.data ?? []).includes(predeclareTitle), JSON.stringify(menuFixes));
  const ranFix = await api.act("problemFixes", { module: "Uses", workbook: BOOK, line: 8, column: 17, title: predeclareTitle });
  const predeclaredFromUses = await settles("Registry predeclared from Uses' fix", async () =>
    (await describe("Registry")).attributes?.predeclaredId === true, { budgetMs: 25000 });
  const registryNow = (await api.readModule("Registry", BOOK)).text ?? "";
  const usesClean = await settles("Uses clean", async () =>
    !((await api.problems("Uses")).findings ?? []).some((one) => one.code === "undeclared-variable"), { budgetMs: 25000 });
  check("running it adds '@PredeclaredId to Registry, applies it, and the report in Uses goes",
    ranFix.did && predeclaredFromUses && /^'@PredeclaredId/.test(registryNow) && usesClean,
    JSON.stringify({ ran: ranFix.did, predeclared: predeclaredFromUses, head: registryNow.split(/\r?\n/)[0], usesClean }));

  // Back to the built text, with the attribute off again, so the save below has drift to write.
  await api.writeModule("Registry", registryText, BOOK);
  await api.attributesRemove("Registry", "PredeclaredId", at);
  registryText = null;

  // ---- a save writes the annotations, under the setting ----
  // Registry's annotation is back and its attribute is off, so the drift is there to be written.
  await settles("Registry's drift back", async () => (await describe("Registry")).drift.some((one) => one.code === "annotation-not-applied"), { budgetMs: 20000 });
  await api.settings({ applyAttributesOnSave: true });
  await api.command("save");
  await wait(2500);
  const afterSave = await describe("Registry");
  check("saving writes the annotations first, and the saved file carries them with nothing left asserted",
    afterSave.attributes?.predeclaredId === true && !afterSave.asserted && afterSave.drift.length === 0,
    JSON.stringify({ asserted: afterSave.asserted, predeclared: afterSave.attributes?.predeclaredId, drift: afterSave.drift.map((one) => one.code) }));
  await api.settings({ applyAttributesOnSave: false });

  // ---- refusals, in words ----
  const onSheet = await refusalOf(api.attributesApply("Sheet1", at));
  check("a document module is refused with the reason", /document module/.test(onSheet ?? ""), onSheet);
  const onUses = await refusalOf(api.attributesApply("Uses", at));
  check("a module with no annotations is refused with what to do", /carries no attribute annotations/.test(onUses ?? ""), onUses);

  const stats = await api.stats();
  check("no COM wrapper was leaked by any of it", stats.comWrappersLive < 100, `${stats.comWrappersLive} live`);
} finally {
  if (registryText !== null) {
    await api.writeModule("Registry", registryText, BOOK).catch(() => {});
  }
  await resetToBuilt("reset at end").catch((error) => console.log(`  reset at end failed: ${error.message}`));
  await api.settings({ applyAttributesOnSave: originalSettings.applyAttributesOnSave !== false }).catch(() => {});
}

process.exitCode = done();
