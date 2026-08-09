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
import { open } from "./xlide-api.mjs";

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

  // The header is what makes a file come back as the same KIND of module. Without it a sheet
  // would import as an ordinary class, which is a silent corruption of the workbook.
  const sheet = readFileSync(join(folder, "ThisWorkbook.cls"), "utf8");
  check("a document file carries its attribute header", sheet.includes("Attribute VB_Name"));
  check("and the marks that make it a document", sheet.includes("VB_PredeclaredId = True"));

  const again = await api.syncPlan("export", { folder });
  check("exporting twice changes nothing the second time", again.items.every((i) => i.status === "unchanged"));

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
  rmSync(join(folder, "ThisWorkbook.cls"));
  rmSync(join(folder, "Sheet1.cls"));
  rmSync(join(folder, "Helper.bas"));
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

  const edited = `${readFileSync(join(folder, "Helper.bas"), "utf8").trimEnd()}\r\n\r\nPublic Sub ByTheDialog()\r\nEnd Sub\r\n`;
  writeFileSync(join(folder, "Helper.bas"), edited, "utf8");

  // Through the dialog: open it, switch to import, tick nothing extra, press Apply.
  await api.eval(`(() => document.querySelector("button[data-command=openSync]").click())()`);
  await api.until(`document.getElementById("sync-card") !== null`, { waitMs: 15000 }).catch(() => {});
  await api.eval(`(() => { document.getElementById("sync-folder").value = ${JSON.stringify(folder)};
    document.getElementById("sync-folder").dispatchEvent(new Event("change")); return "set"; })()`);
  await new Promise((r) => setTimeout(r, 2500));
  await api.eval(`(() => document.querySelector(".sync-direction[data-direction=import]").click())()`);
  await new Promise((r) => setTimeout(r, 3000));

  const dialogRows = await api.eval(`(() => [...document.querySelectorAll(".sync-item")]
    .map(r => r.querySelector(".sync-item-name").textContent + ":" + r.querySelector(".sync-chip").textContent)
    .join(","))()`);
  check("the dialog drew the same plan the api answers",
    String(dialogRows.result).includes("Helper.bas:update"));

  await api.eval(`(() => [...document.querySelectorAll("#sync-foot button")].find(b => b.textContent === "Apply").click())()`);
  await new Promise((r) => setTimeout(r, 4000));

  const afterDialog = (await api.readModule("Helper", project.projectId)).text;
  check("the dialog's Apply changed the module", afterDialog.includes("ByTheDialog"));

  await api.eval(`(() => document.getElementById("sync-close").click())()`);

  // Now put it back and do the identical thing through the route.
  await api.writeModule("Helper", beforeText, project.projectId);
  await new Promise((r) => setTimeout(r, 1500));
  await api.syncApply("import", { folder });
  await new Promise((r) => setTimeout(r, 1500));
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
