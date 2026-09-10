/*
 * WORD, WHICH IS NOT EXCEL WITH A DIFFERENT ICON EITHER, against a real one.
 *
 * Access got its suite on 2026-09-06, when the test runner turned out never to have run there.
 * Word had a launcher from 2026-08-19 and nothing to open with it, so every Word branch in the
 * product was written by symmetry and proven by nobody: the `_WwG` document pane the application
 * is reached through, `Application.Run` spelled `Module.Proc`, and the Documents collection the
 * save and the dirty check behind source control walk. This suite is where they are proven
 * (2026-09-10), and it carries the two things a Word session has that no Excel session does:
 *
 *   - NORMAL IS ALWAYS OPEN. Word's global template is a VBA project of its own, so the editor
 *     lists two projects for one document, and anything that writes into "the active project"
 *     has a way to write into the developer's own Normal.dotm. Every write here names the
 *     document, and Normal is held to what it was when the suite started.
 *   - A DOCUMENT CARRIES ITS MODULES, unlike an Access database, so the editor's Save is the
 *     whole save and Documents(name).Saved is the dirty flag.
 *
 * Runs against WordFixture.docm alone:
 *
 *   tools\harness\Start-Word.ps1 -Document artifacts\fixtures\WordFixture.docm -Fresh
 *   node tools\harness\word.mjs
 *
 * It puts back everything it changes, so a rerun starts where the first run did.
 */
import { mkdirSync, realpathSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { open, wait, waitFor, reporter } from "./xlide-api.mjs";

const api = await open();
const { check, done } = reporter();

const BOOK = "WordFixture.docm";
const at = { project: BOOK };
const isBook = (name) => String(name ?? "").toLowerCase() === BOOK.toLowerCase();

const refusalOf = async (promise) => {
  try {
    await promise;
    return null;
  } catch (error) {
    return String(error.message ?? error);
  }
};

const scaffolding = (components) =>
  components.filter((one) => /^XlideRun|^XlideTestDispatch$/i.test(one.name)).map((one) => one.name);

const namesOf = (project) => project.components.map((one) => one.name).sort().join(",");

// NAMED PER RUN. The editor refuses the name of a form removed earlier in the same session - the
// add takes, the rename to it fails with a bare HRESULT and the add is taken back out - so a
// second run in one session on a fixed name would report a UserForm Word cannot add. It can;
// what it cannot do is reuse the name.
const FORM = `XlideWordForm${Date.now() % 100000}`;
let helpersText = null;
let formAdded = false;

try {
  // ---- two projects by construction ----
  const listed = (await api.projects()).projects;
  const mine = listed.find((one) => isBook(one.project));
  const normal = listed.find((one) => /normal/i.test(one.project));
  check("the editor lists the document and Word's Normal template: two projects for one document",
    listed.length === 2 && mine !== undefined && normal !== undefined,
    listed.map((one) => one.project).join(", "));
  const normalBefore = normal ? namesOf(await api.project(normal.project)) : "(no Normal listed)";

  // ThisDocument rides along: a Word document has a document module the way a workbook has
  // ThisWorkbook, which an Access database does not.
  const FIXTURE = "Helpers,Pricing,PricingTests,ThisDocument,XlideAssert";
  const project = await api.project(BOOK);
  check("the document's modules are all there: the fixture's standard modules and ThisDocument",
    namesOf(project) === FIXTURE
      && project.components.every((one) => one.kind === (one.name === "ThisDocument" ? "document" : "module")),
    project.components.map((one) => `${one.name}:${one.kind}`).join(","));

  // ---- the document's project is the active one ----
  // The Immediate window evaluates and the runner runs in the ACTIVE project, and with Normal
  // beside the document that is a real question here. Opening one of the document's modules is
  // how a developer answers it, so it is how this suite does.
  await api.caret(1, { module: "Pricing", column: 1, ...at });

  // ---- the host is reachable, which is what the Immediate window proves ----
  // "The host application could not be reached" is what these answer when the window walk does
  // not find Word's document pane, so they are the check that the host is reachable AT ALL.
  const arithmetic = await api.immediate("?1+1");
  check("the Immediate window evaluates in Word", arithmetic.ran && arithmetic.text === "2" && !arithmetic.failed,
    JSON.stringify(arithmetic));

  const intoTheProject = await api.immediate("?Pricing.DiscountRate(10)");
  check("and evaluates against the document's own code", intoTheProject.text === "0.1", JSON.stringify(intoTheProject));

  const greeting = await api.immediate("?Greeting()");
  check("a Module.Proc name resolves, which is the spelling Word's Run takes",
    greeting.text === "hello from word", JSON.stringify(greeting));

  // ---- the support module, per file ----
  // The whole-session standing is the WORST among the files the pane lists, and in Word the list
  // always holds Normal, which has no tests and no XlideAssert. So the document's own standing
  // is what is checked, the way a developer who has chosen their document in the pane sees it.
  const listedTests = await api.tests({ action: "refresh" });
  const files = listedTests.files ?? [];
  const standing = files.find((one) => isBook(one.file));
  const normalStanding = files.find((one) => /normal/i.test(one.file));
  check("the pane lists both files: XlideAssert INSTALLED in the document with its four tests, "
    + "Normal beside it with none, and nothing installed there",
    files.length === 2 && standing?.support === "installed" && standing?.tests === 4
      && normalStanding !== undefined && normalStanding.tests === 0 && normalStanding.support !== "installed",
    JSON.stringify(files));
  check("all four tests are discovered, every one in the document",
    listedTests.rows.length === 4 && listedTests.rows.every((one) => isBook(one.file)),
    listedTests.rows.map((one) => `${one.file}:${one.procedure}`).join(","));

  // ---- the run itself ----
  const ran = await api.tests({ action: "run" });
  const outcome = (procedure) => ran.rows.find((one) => one.procedure === procedure);
  check("the run reaches the host and reports per test", /^ran 4 in/.test(ran.detail), ran.detail);
  check("the three that should pass, pass",
    ["NoDiscountUnderFive", "FivePercentFromFive", "ExpressAddsItsFee"].every((one) => outcome(one)?.status === "passed"),
    ran.rows.map((one) => one.procedure + "=" + one.status).join(", "));
  check("and the one that fails on purpose fails, in the assertion's own words",
    outcome("FailsOnPurpose")?.status === "failed"
      && /Expected True but was False/i.test(outcome("FailsOnPurpose")?.message ?? ""),
    outcome("FailsOnPurpose")?.message);

  // The scaffolding goes again: a generated run module is written into the document and taken
  // out, and a document saved with one left standing would carry it for ever.
  const after = await api.project(BOOK);
  check("the generated run modules are gone afterwards", scaffolding(after.components).length === 0,
    scaffolding(after.components).join(",") || "none");

  const rerun = await api.tests({ action: "run" });
  check("and a second run is as good as the first, so nothing was left half-standing",
    rerun.rows.filter((one) => one.status === "passed").length === 3,
    rerun.rows.map((one) => one.procedure + "=" + one.status).join(", "));

  // ---- what Word has that Access does not ----
  // Word's VBA carries MSForms as Excel's does, so a UserForm is an ordinary add here - the same
  // route that refuses in Access, answering the other way for the other host.
  await api.component("add", { kind: "form", name: FORM, ...at });
  formAdded = true;
  const withForm = await api.project(BOOK);
  check("a UserForm can be added in Word, whose VBA carries MSForms",
    withForm.components.some((one) => one.name === FORM && one.kind === "form"),
    withForm.components.map((one) => `${one.name}:${one.kind}`).join(","));

  await api.pane("close", { module: FORM, ...at, answer: "discard" }).catch(() => {});
  await api.component("remove", { name: FORM, ...at });
  formAdded = false;
  check("and removed again, the document's own modules untouched",
    namesOf(await api.project(BOOK)) === FIXTURE);

  // ---- a module round trip through the door ----
  helpersText = (await api.readModule("Helpers", BOOK)).text ?? "";
  const edited = helpersText.replace("hello from word", "hello again");
  await api.writeModule("Helpers", edited, BOOK, { by: "word.mjs" });
  const readBack = (await api.readModule("Helpers", BOOK)).text ?? "";
  check("a module written through the door reads back as it was written",
    readBack.includes("hello again"), readBack.split(/\r?\n/).length + " lines");

  const afterWrite = await api.immediate("?Greeting()");
  check("and the host runs the new text, not the old",
    afterWrite.text === "hello again", JSON.stringify(afterWrite));

  await api.writeModule("Helpers", helpersText, BOOK, { by: "word.mjs" });
  helpersText = null;
  check("and it goes back", (await api.immediate("?Greeting()")).text === "hello from word");

  // ---- source control, whose save and dirty checks are the host's ----
  //
  // Excel answers them from Workbooks; Word answers from Documents, the collection with the same
  // Name, Saved and Save members - the branch written by symmetry on 2026-09-09 and proven here.
  // A commit saves the document, an unsaved module edit reads as dirty, a checkout over it is
  // refused, the editor's Save cleans it, and a commit over an unsaved edit saves first.
  const scratch = join(tmpdir(), `xlide-word-scm-${process.pid}`);
  rmSync(scratch, { recursive: true, force: true });
  mkdirSync(scratch, { recursive: true });
  const folder = realpathSync.native(scratch);
  const scm = (args = {}) => api.scm({ ...args, ...at });
  try {
    await scm({ action: "settings", folder });
    await scm({ action: "init" });
    await scm({ action: "identity", name: "Word Suite", email: "word@example.com" });
    const first = await scm({ action: "commit", message: "the fixture", by: "word.mjs" });
    check("the document's modules commit, saving the document on the way",
      (first.committed ?? []).includes("Helpers") && first.status?.dirty === false,
      `${(first.committed ?? []).length} committed, dirty ${first.status?.dirty}`);

    helpersText = (await api.readModule("Helpers", BOOK)).text ?? "";
    await api.writeModule("Helpers", helpersText.replace("hello from word", "hello unsaved"), BOOK, { by: "word.mjs" });
    const dirty = await waitFor("the document to read as dirty", async () => (await scm()).dirty === true, { budgetMs: 10000 })
      .then(() => true).catch(() => false);
    check("an unsaved module edit makes the document dirty, which is Documents(name).Saved", dirty);
    const refused = await refusalOf(scm({ action: "checkout", ref: "HEAD" }));
    check("and a checkout over it is refused, as it is in Excel", refused !== null && /unsaved|save/i.test(refused), refused ?? "went through");
    check("with the edit still in the module", ((await api.readModule("Helpers", BOOK)).text ?? "").includes("hello unsaved"));

    await api.command("save");
    const clean = await waitFor("the save to clean the document", async () => (await scm()).dirty === false, { budgetMs: 10000 })
      .then(() => true).catch(() => false);
    check("the editor's own Save cleans it", clean);

    const second = await scm({ action: "commit", message: "the edit", modules: ["Helpers"], by: "word.mjs" });
    check("and the edit commits, the document clean",
      (second.committed ?? []).includes("Helpers") && second.status?.dirty === false, `dirty ${second.status?.dirty}`);

    // Put back through a commit too, which is the save-then-commit path over a dirty document.
    await api.writeModule("Helpers", helpersText, BOOK, { by: "word.mjs" });
    const third = await scm({ action: "commit", message: "put back", modules: ["Helpers"], by: "word.mjs" });
    check("a commit over an unsaved edit saves the document itself first",
      (third.committed ?? []).includes("Helpers") && third.status?.dirty === false, `dirty ${third.status?.dirty}`);
    helpersText = null;
  } finally {
    await scm({ action: "forget" }).catch(() => {});
    await wait(500);
    rmSync(scratch, { recursive: true, force: true });
  }

  // ---- Normal, untouched ----
  // The one check that would catch a write aimed at "the active project" landing in the
  // developer's template: everything above added, ran, removed, committed and saved, and Normal
  // holds exactly what it held at the start.
  const normalAfter = normal ? namesOf(await api.project(normal.project)) : "(no Normal listed)";
  check("Word's Normal template holds what it held before any of it",
    normal !== undefined && normalAfter === normalBefore, `${normalBefore} -> ${normalAfter}`);

  const stats = await api.stats();
  check("no COM wrapper was leaked by any of it", stats.comWrappersLive < 100, `${stats.comWrappersLive} live`);
} finally {
  if (helpersText !== null) {
    await api.writeModule("Helpers", helpersText, BOOK, { by: "word.mjs" }).catch(() => {});
  }
  if (formAdded) {
    await api.pane("close", { module: FORM, ...at, answer: "discard" }).catch(() => {});
    await api.component("remove", { name: FORM, ...at }).catch(() => {});
  }
}

done();
