/*
 * Import and export, end to end, and the one property that matters most about them: the dialog and
 * the api leave the project in the SAME state.
 *
 * That is not a style preference. The rule this product is built to is that an api action leaves
 * the state the equivalent UI action would, and the whole design of this surface, one service in
 * the host, two doors onto it, exists to make it true by construction. A test that only drove the
 * api would pass on a product whose button did something else entirely, which is exactly the kind
 * of green nobody should trust.
 *
 * Run against DebugFixture.xlsm with the editor open.
 */

import { mkdirSync, readFileSync, rmSync, writeFileSync, existsSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { open, wait, waitFor, waitUntilStable } from "./xlide-api.mjs";

/** Every row the sync dialog is drawing, as "name:status" - one string, easy to compare. */
const SYNC_ROWS = `(() => [...document.querySelectorAll(".sync-item")]
  .map(r => r.querySelector(".sync-item-name").textContent + ":" + r.querySelector(".sync-chip").textContent)
  .join(","))()`;

const api = await open();
let passed = 0;
let failed = 0;

// WHICH PLANNER. Everything below runs identically against either, which is the point of having
// two: the choice decides who works out what an import would do, never who does it or what the
// answer means. Pass "builtIn" to check the other one; the gate runs both.
const planner = process.argv[2] === "builtIn" ? "builtIn" : "xlide";
// Put back at the end. A suite that changes a developer's setting and walks away leaves them on
// whichever planner ran last, which is a surprise the next time they export.
const plannerWas = (await api.settings()).syncEngine;
await api.settings({ syncEngine: planner });
const chosen = (await api.settings()).syncEngine;
if (chosen !== planner) {
  console.log(`FAIL the planner would not switch: asked ${planner}, got ${chosen}`);
  process.exit(1);
}

console.log(`planner: ${planner}
`);

const check = (name, got, want = true) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (ok) {
    passed++;
    console.log(`ok   ${name}`);
  } else {
    failed++;
    console.log(`FAIL ${name}\n       got  ${JSON.stringify(got)}\n       want ${JSON.stringify(want)}`);
  }
  return ok;
};

// UNIQUE PER RUN. A fixed name inherits the previous run's module when a cleanup did not take,
// and then the plan reads "unchanged" instead of "will-create" and the suite reports a defect that
// is really yesterday's leftovers. Written down in docs/disambiguation.md after the freshness
// suite did exactly this, and done here after it happened again (2026-08-09).
const probe = `SyncProbe${process.pid}`;
const stale = `SyncStale${process.pid}`;

const folder = join(tmpdir(), `xlide-sync-${process.pid}`);
rmSync(folder, { recursive: true, force: true });
mkdirSync(folder, { recursive: true });

const project = (await api.projects()).projects[0];
const made = [];
// The one FIXTURE module this suite edits (section 5 imports a Helper.bas with a sub appended),
// held so the cleanup can put it back. `made` cannot cover it: Helper is not added, only changed.
let helperOriginal = null;

const cleanUp = async () => {
  for (const name of made) {
    try {
      await api.component("remove", { name, project: project.projectId });
    } catch (error) {
      // Said out loud. A cleanup that fails silently leaves a module behind, and the next run
      // reads it as "unchanged" and blames the product.
      console.log(`     WARNING: ${name} could not be removed (${error.message})`);
    }
  }

  // Helper goes back as found. Left dirty it holds a ByTheDialog sub, which one run gets away
  // with - a spare sub still compiles - and the second run in the same session stacks another,
  // at which point the PROJECT stops compiling ("Ambiguous name detected") and the failure lands
  // on whichever suite next asks for a compile. The gate runs this suite twice, once per
  // planner, and that is exactly how it surfaced (2026-08-12).
  if (helperOriginal !== null) {
    try {
      await api.writeModule("Helper", helperOriginal, project.projectId);
      await waitFor("Helper to be back as found", async () =>
        (await api.readModule("Helper", project.projectId)).text === helperOriginal);
    } catch (error) {
      console.log(`     WARNING: Helper was left holding this run's import (${error.message})`);
    }
  }

  rmSync(folder, { recursive: true, force: true });

  try {
    await api.settings({ syncEngine: plannerWas });
  } catch (error) {
    console.log(`     WARNING: the planner setting was left on ${planner} (${error.message})`);
  }
};

try {
  console.log(`against ${project.project}\nfolder ${folder}\n`);

  // ---------------------------------------------------------------------------------------
  console.log("1. an export writes the project out");

  const plan = await api.syncPlan("export", { folder });

  // THE PLANNER THAT ANSWERED, asserted before anything else is read.
  //
  // The fallback to the built-in planner is silent on purpose - a developer pressing Export while
  // the engine is starting should get their export - which means a suite that does not check this
  // tests whichever planner happened to answer. This one did exactly that: 31 green checks against
  // the built-in planner while asking for the shared one, because the modules could not be
  // serialised to the engine and nothing said so (2026-08-09).
  check("the planner that answered is the one that was asked for", plan.planner, planner);

  check("every module is offered, and every one is new", plan.items.every((i) => i.status === "will-create"));
  check("a standard module is a .bas", plan.items.some((i) => i.file === "Helper.bas"));
  check("a document module is a .cls", plan.items.some((i) => i.file === "ThisWorkbook.cls"));

  const wrote = await api.syncApply("export", { folder });
  check("nothing failed", wrote.failed, []);
  check("every module was written", wrote.changed.length === plan.items.length);

  const written = readdirSync(folder).sort();
  check("the files are on disk", written.length === plan.items.length);

  // A file is written beside its destination and then moved over it, so nothing ever reads a
  // module that is half written - not the companion editor watching the folder, not a build, not
  // another Excel importing from the same folder, which is the case this side has no lock for.
  // The count above would not notice a leftover, because it counts what it expected to find.
  check("and nothing half written was left beside them",
    written.every((name) => !name.endsWith(".xlide-partial")), true);

  // The header is what makes a file come back as the same KIND of module. Without it a sheet
  // would import as an ordinary class, which is a silent corruption of the workbook.
  const sheet = readFileSync(join(folder, "ThisWorkbook.cls"), "utf8");
  check("a document file carries its attribute header", sheet.includes("Attribute VB_Name"));
  check("and the marks that make it a document", sheet.includes("VB_PredeclaredId = True"));

  const again = await api.syncPlan("export", { folder });
  check("exporting twice changes nothing the second time", again.items.every((i) => i.status === "unchanged"));

  // AN UNCHANGED ROW IS DRAWN THE SAME WAY BY EITHER PLANNER, AND IS ONE LINE.
  //
  // Both stopped comparing a row whose two sides are known to agree, for the same reason and a
  // pipe apart: the built-in one skipped the comparison (414ms of planning became 5ms) and the
  // engine stopped sending it (1,417ms of a 1,710ms plan). Both write the one line the condensing
  // would have left, so the row is the same object whichever answered.
  //
  // Checked here because nothing else looks at a comparison at all, and this suite runs once per
  // planner - so the pair of runs is what holds the two together. Without it the shared planner
  // could go back to shipping 163,000 comparison entries, or stop drawing changed rows entirely,
  // and every other check would still pass (2026-08-09).
  const drawn = again.items[0]?.diff ?? [];
  check("an unchanged row draws one line, not its whole text", drawn.length === 1);
  check("and that line says how many lines agreed", drawn[0]?.kind, "gap");
  check("and counts them", /^[\d,]+ identical lines?$/.test(drawn[0]?.left ?? ""));

  // ---------------------------------------------------------------------------------------
  console.log("\n2. an import reads it back, including modules the project does not have");

  writeFileSync(join(folder, `${probe}.bas`), [
    `Attribute VB_Name = "${probe}"`,
    "Option Explicit",
    "",
    "Public Function OneTwoThree() As Long",
    "    OneTwoThree = 123",
    "End Function",
    "",
  ].join("\r\n"), "utf8");
  made.push(probe);

  const incoming = await api.syncPlan("import", { folder });
  const fresh = incoming.items.find((i) => i.file === `${probe}.bas`);
  check("the new file will create a module", fresh?.status, "will-create");
  check("and it is ticked, because it does something", fresh?.checked, true);
  check("everything already in the project is left alone",
    incoming.items.filter((i) => i.file !== `${probe}.bas`).every((i) => i.status === "unchanged"));

  const read = await api.syncApply("import", { folder });
  check("the import reported no failures", read.failed, []);

  await api.waitFor(() => true, { timeout: 1 }).catch(() => {});
  const source = await api.readModule(probe, project.projectId);
  check("the module holds what the file held", source.text.includes("OneTwoThree = 123"));

  // ---------------------------------------------------------------------------------------
  console.log("\n3. a sheet cannot be conjured from a file, and says so instead of failing");

  writeFileSync(join(folder, "SheetNine.cls"), [
    "VERSION 1.0 CLASS",
    "BEGIN",
    "  MultiUse = -1  'True",
    "END",
    'Attribute VB_Name = "SheetNine"',
    "Attribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"",
    "Attribute VB_PredeclaredId = True",
    "Option Explicit",
    "",
  ].join("\r\n"), "utf8");

  const withSheet = await api.syncPlan("import", { folder });
  const ghost = withSheet.items.find((i) => i.file === "SheetNine.cls");
  check("a document the project does not have is skipped", ghost?.status, "skipping-import");
  check("it is not ticked", ghost?.checked, false);
  check("and the reason is on the row", typeof ghost?.warning === "string" && ghost.warning.length > 0);
  check("and said once at the top as well", withSheet.warnings.length > 0);

  rmSync(join(folder, "SheetNine.cls"));

  // ---------------------------------------------------------------------------------------
  console.log("\n4. removing is opt-in, in both directions");

  writeFileSync(join(folder, `${stale}.bas`), `Attribute VB_Name = "${stale}"\r\nOption Explicit\r\n`, "utf8");

  const leaveAlone = await api.syncPlan("export", { folder, mode: "exportAll" });
  check("by default an export leaves an unmatched file alone",
    leaveAlone.items.every((i) => i.status !== "will-remove"));

  const tidy = await api.syncPlan("export", { folder, mode: "trueUp" });
  const doomed = tidy.items.find((i) => i.file === `${stale}.bas`);
  check("asked to match, it offers to delete it", doomed?.status, "will-remove");
  check("with a warning, because it is a delete", typeof doomed?.warning === "string");

  await api.syncApply("export", { folder, mode: "trueUp", ids: [doomed.id] });
  check("and the file is gone", existsSync(join(folder, `${stale}.bas`)), false);

  // Import true-up must never offer to delete a sheet or the workbook: they belong to the
  // workbook, not to the folder, and the folder will never have a file for them.
  // `force`, so a file that is not there fails the CHECK below rather than throwing out of the
  // suite. Run against the wrong fixture this crashed in cleanup with an ENOENT naming a path,
  // which says nothing about what went wrong and hides the four checks that would have.
  for (const gone of ["ThisWorkbook.cls", "Sheet1.cls", "Helper.bas"]) {
    rmSync(join(folder, gone), { force: true });
  }
  const strict = await api.syncPlan("import", { folder, mode: "trueUpStandardClass" });
  const removals = strict.items.filter((i) => i.status === "will-remove").map((i) => i.module);
  check("import true-up offers to delete the standard module", removals.includes("Helper"));
  check("but never the workbook", removals.includes("ThisWorkbook"), false);
  check("and never a sheet", removals.includes("Sheet1"), false);

  // Put them back, so the folder and the project agree again for the next part.
  await api.syncApply("export", { folder });

  // ---------------------------------------------------------------------------------------
  console.log("\n5. the dialog and the api leave the same state");

  // The same change, applied twice: once by driving the dialog's own controls, once through the
  // route. If these two ever diverge, one of them has grown its own idea of what an import is.
  const beforeText = (await api.readModule("Helper", project.projectId)).text;
  helperOriginal = beforeText;

  const edited = `${readFileSync(join(folder, "Helper.bas"), "utf8").trimEnd()}\r\n\r\nPublic Sub ByTheDialog()\r\nEnd Sub\r\n`;
  writeFileSync(join(folder, "Helper.bas"), edited, "utf8");

  // Through the dialog: open it, switch to import, tick nothing extra, press Apply.
  //
  // The open goes through the toolbar act rather than a selector now. Finding the button and
  // clicking it kept passing whatever happened to the button - renamed, disabled, or left out of
  // the build because its editor action was not bundled - because it never went through the
  // dispatch that would have noticed. The act presses the button the strip drew and says which
  // commands it is drawing when there is no such one.
  const opened = await api.act("toolbar", { command: "openSync" });
  if (!opened.did) {
    throw new Error(`the sync dialog could not be opened from the toolbar: ${opened.detail}`);
  }
  await api.until(`document.getElementById("sync-card") !== null`, { waitMs: 15000 }).catch(() => {});
  await api.eval(`(() => { document.getElementById("sync-folder").value = ${JSON.stringify(folder)};
    document.getElementById("sync-folder").dispatchEvent(new Event("change")); return "set"; })()`);
  // The folder having reached the dialog, and then the plan having been drawn AT ALL. Not that it
  // drew the right row: that is the check below and it has to stay able to fail.
  await waitFor("the dialog to take the folder", async () =>
    String((await api.eval(
      `(() => document.getElementById("sync-folder").value)()`)).result ?? "").length > 0);

  /*
   * THE PLAN HAS BEEN REDRAWN FOR THE NEW DIRECTION, which is not the same as there being rows.
   *
   * Waiting for `.sync-item` to be non-empty is satisfied instantly by the EXPORT rows still on
   * screen from a moment ago, and the check below then reads the wrong plan and fails. Measured
   * exactly that on the shared planner, which takes longer to answer than the built-in one, so
   * the suite went red on one planner and green on the other for a reason that was purely the
   * harness (2026-08-10).
   *
   * So: what the rows say has to CHANGE, and then stop changing. Neither asks whether the right
   * row is there - that is the check, and it can still fail.
   */
  const rowsBefore = String((await api.eval(SYNC_ROWS)).result ?? "");
  await api.eval(`(() => document.querySelector(".sync-direction[data-direction=import]").click())()`);
  await waitFor("the dialog to redraw for the import direction", async () =>
    String((await api.eval(SYNC_ROWS)).result ?? "") !== rowsBefore);
  await waitUntilStable(async () => String((await api.eval(SYNC_ROWS)).result ?? ""));

  const dialogRows = await api.eval(SYNC_ROWS);
  check("the dialog drew the same plan the api answers",
    String(dialogRows.result).includes("Helper.bas:update"));

  await api.eval(`(() => [...document.querySelectorAll("#sync-foot button")].find(b => b.textContent === "Apply").click())()`);
  // CHANGED, not changed to the right thing. An apply that wrote the wrong text satisfies this and
  // still fails the check below, which is the whole point of not waiting for "ByTheDialog" here.
  await waitFor("the dialog's Apply to reach the module", async () =>
    (await api.readModule("Helper", project.projectId)).text !== beforeText);

  const afterDialog = (await api.readModule("Helper", project.projectId)).text;
  check("the dialog's Apply changed the module", afterDialog.includes("ByTheDialog"));

  await api.eval(`(() => document.getElementById("sync-close").click())()`);

  // Now put it back and do the identical thing through the route.
  await api.writeModule("Helper", beforeText, project.projectId);
  // PUT BACK BEFORE THE SECOND IMPORT, waited for rather than assumed. The two applies are being
  // compared byte for byte, so an import that started from a module still holding the dialog's
  // result would agree with it for the wrong reason.
  await waitFor("the module to be back where the dialog found it", async () =>
    (await api.readModule("Helper", project.projectId)).text === beforeText);

  await api.syncApply("import", { folder });
  await waitFor("the route's apply to reach the module", async () =>
    (await api.readModule("Helper", project.projectId)).text !== beforeText);

  const afterApi = (await api.readModule("Helper", project.projectId)).text;

  check("the api's apply changed it the same way", afterApi.includes("ByTheDialog"));
  check("BYTE FOR BYTE the same state either way", afterApi === afterDialog);

  // ---------------------------------------------------------------------------------------
  console.log("\n6. an import reaches everything that shows the project, not just the store");

  const tree = await api.eval(`(() => [...document.querySelectorAll(".tree-item")]
    .map(e => (e.textContent || "").trim().split("\\n")[0]).join(","))()`);
  check("the explorer shows a module that arrived from the folder",
    String(tree.result).includes(probe));

  const panes = await api.native();
  const helperPane = panes.panes?.find((p) => p.module === "Helper");
  if (helperPane && helperPane.surfaceContent !== null) {
    check("the native editor and the page agree about the imported module",
      helperPane.hostContent, helperPane.surfaceContent);
  } else {
    console.log("     (Helper has no open pane, so there is nothing to compare)");
  }

  // ---------------------------------------------------------------------------------------
  console.log("\n7. text that is not English, and text the host will not hold");

  /*
   * WHAT THIS CAN AND CANNOT ASSERT, because the answer is a property of the MACHINE.
   *
   * VBA stores module text in the system ANSI code page, so which scripts survive depends on
   * where this runs: on code page 1252 accented Latin does and Cyrillic does not, and on 1251 it
   * is the other way round. Asserting a list would make this suite fail on a machine that is
   * working correctly, so it asks the host first and then holds it to two rules:
   *
   *   whatever the module DOES hold must survive export and import exactly, and
   *   a file carrying what the host will NOT hold must be refused, not imported and mangled.
   *
   * The second is the one with a file's life on it. Before the guard, one import and one export
   * turned a repository file carrying Cyrillic into question marks, byte for byte, reporting
   * "1 changed, 0 failed" at both ends (2026-08-09).
   *
   * The planner-independent half runs in CI: Xlide.Vbe.Core.Tests drives the whole matrix through
   * PlanExport and PlanImport, and engine/test/language.mjs drives it through the shared planner
   * on Linux as well as Windows. Only the HOST's conversion needs a real Excel, and that is here.
   */
  const langModule = `Lang${process.pid}`;
  const CANDIDATES = [
    ["Western European", "déjà vu € œuvre Straße"],
    ["Cyrillic", "Проверка русского текста"],
    ["Greek", "Δοκιμή ελληνικού κειμένου"],
    ["Japanese", "テスト用モジュール"],
    ["Latin, decomposed", "café"],
  ];

  const langSource = ["Option Explicit", "",
    ...CANDIDATES.map(([name, text], i) => `Public Const L${i} As String = "${text}" ' ${name}`)].join("\r\n");

  await api.component("add", { kind: "module", name: langModule, project: project.projectId });
  made.push(langModule);
  await api.writeModule(langModule, langSource, project.projectId).catch(() => {});
  // Tolerant: the write is allowed to be refused, which is what the `.catch` above is for, and
  // this host drops any character outside its ANSI code page - so "the text came back" is not a
  // condition that always holds. Waited for when it does, given up on when it does not.
  await waitFor("the language module to hold something", async () =>
    ((await api.readModule(langModule, project.projectId)).text ?? "").includes("Option Explicit"))
    .catch(() => null);

  const langHeld = (await api.readModule(langModule, project.projectId)).text ?? "";
  const kept = CANDIDATES.filter(([, text]) => langHeld.includes(text));
  const lost = CANDIDATES.filter(([, text]) => !langHeld.includes(text));
  console.log(`     this machine holds ${kept.length} of ${CANDIDATES.length}: `
    + `${kept.map(([n]) => n).join(", ") || "none"}`);

  const langFolder = join(folder, "lang");
  mkdirSync(langFolder, { recursive: true });
  await api.syncApply("export", { folder: langFolder, select: "all" });

  const exported = readFileSync(join(langFolder, `${langModule}.bas`), "utf8");
  check("the exported file is what the module holds, character for character",
    exported.slice(exported.indexOf("Option Explicit")).replace(/\r\n/g, "\n").trimEnd(),
    langHeld.slice(langHeld.indexOf("Option Explicit")).replace(/\r\n/g, "\n").trimEnd());

  if (kept.length > 0) {
    check("and every script the host kept is in the file",
      kept.every(([, text]) => exported.includes(text)));
  } else {
    console.log("     (this machine's code page held none of them, so there is nothing to carry)");
  }

  // A file the host cannot hold must be REFUSED, and must be left alone.
  if (lost.length > 0) {
    const [lostName, lostText] = lost[0];
    const refusedName = `Refused${process.pid}`;
    const refusedPath = join(langFolder, `${refusedName}.bas`);
    const refusedSource = [`Attribute VB_Name = "${refusedName}"`, "Option Explicit", "",
      `Public Const A As String = "${lostText}"`, ""].join("\r\n");
    writeFileSync(refusedPath, refusedSource, "utf8");
    const beforeImport = readFileSync(refusedPath);

    const attempt = await api.syncApply("import", { folder: langFolder, select: "all" });
    // A SLEEP THAT STAYS, because it is giving the defect time to happen rather than waiting for
    // a result. `attempt` is already in hand; what the checks below want is a chance for a
    // half-made module to appear before asserting that none did. Waiting for a condition here
    // would mean waiting for the bug, and there is nothing to wait for when the product is right.
    await wait(1500);

    // `check` compares against an EXPECTED VALUE; it has no detail argument. Passing the detail
    // as one makes every call fail with the detail printed as what it wanted, which is what
    // happened the first time this was written.
    if (!check(`a file carrying ${lostName}, which this host cannot hold, is refused`,
      (attempt.failed ?? []).length > 0)) {
      console.log(`       import said: ${JSON.stringify(attempt)}`);
    }

    check("and the refusal names the character rather than just failing",
      (attempt.failed ?? []).some((line) => /cannot store/.test(line)));
    check("and the file on disk is untouched",
      beforeImport.equals(readFileSync(refusedPath)));

    const conjured = (await api.project()).components.some((c) => c.name === refusedName);
    check("and no half-made module was left in the project", conjured, false);
    if (conjured) { made.push(refusedName); }
  } else {
    console.log("     (this machine's code page held every sample, so there is nothing to refuse)");
  }

  // Section 8, an import must not write over unwritten edits, lives in import-guard.mjs
  // rather than here. It needs the module ON THE SURFACE with edits pending, and by this point
  // this suite has opened, closed and re-opened enough that the shim's document key for a
  // freshly opened module no longer matches what a live read asks with. A check whose
  // precondition is unreliable reports on the precondition and not on its subject, so it runs
  // where the precondition holds (2026-08-09).

  // Reading every module of a project and exporting each header walks a great deal of COM. The
  // number that matters is not how many are live, which rests above zero by design, because the
  // session holds the editor and its projects for as long as it is up, but whether a wrapper was
  // counted home WITHOUT being released. That one is given back by the finaliser thread instead,
  // where releasing an apartment-threaded object ends the host process.
  const wrappers = await api.stats();
  const notReleased = wrappers.comWrappersGivenBack - wrappers.comWrappersDisposed;
  console.log(`\n     com wrappers: taken ${wrappers.comWrappersTaken},`
    + ` given back ${wrappers.comWrappersGivenBack},`
    + ` released ${wrappers.comWrappersDisposed}, live ${wrappers.comWrappersLive}`);
  check("every wrapper given back was actually released", notReleased, 0);
} finally {
  await cleanUp();
}

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
