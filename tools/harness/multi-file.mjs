/*
 * TWO FILES OPEN AT ONCE, which is the state everything below is about.
 *
 * A session holds as many VBA projects as the host has files open. The analyzer has always read
 * all of them; the test runner used to see only the front one, so the other file's tests were
 * invisible rather than absent. Widening it raises questions a single-file session cannot ask:
 *
 *   a module name is not an identity - both fixtures hold an InvoiceTests, and a result filed
 *     under the name alone lands on the wrong file's test;
 *   XlideAssert is a module INSIDE a file, so support is per file, and a run must refuse for
 *     the file that lacks it without refusing for the file that has it;
 *   the panes' scope selector grows a file tier, and a whole-file scope has to mean the file
 *     rather than every module that happens to share its name.
 *
 * Runs against TestFixture.xlsm and TestTwinFixture.xlsm - and DebugFixture.xlsm beside them,
 * which holds no tests and no XlideAssert. A THIRD FILE WITH NOTHING TO SAY is not padding: the
 * pane's storage and its painting had two different rules for such a file, and every symptom of
 * that lived where no session with only test-bearing files could see it (see the block below).
 *
 *   tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\TestFixture.xlsm,artifacts\fixtures\TestTwinFixture.xlsm,artifacts\fixtures\DebugFixture.xlsm
 *   node tools\harness\multi-file.mjs
 *
 * It puts back everything it changes: the twin's XlideAssert is removed to prove the per-file
 * gate and reinstalled through the same route the chip presses.
 */
import { open, reporter, waitFor, waitUntilStable } from "./xlide-api.mjs";

const api = await open();
const { check, done } = reporter();

const MAIN = "TestFixture.xlsm";
const TWIN = "TestTwinFixture.xlsm";
const PLAIN = "DebugFixture.xlsm";

const ask = (script) => api.ask(script);
const showPane = (which) =>
  ask(`document.querySelector('.panel-tab[data-panel="${which}"]')?.click(), "shown"`);
const pickScope = (select, startsWith) => ask(
  `(() => { const s = document.querySelector("#${select}");`
  + ` const o = [...s.querySelectorAll("option")].find(x => x.textContent.startsWith(${JSON.stringify(startsWith)}));`
  + " if (!o) return null; s.value = o.value; s.dispatchEvent(new Event('change')); return s.value; })()");

const BROKEN = `ScopeFile${process.pid % 10000}`;
const BROKEN_SOURCE = [
  "Option Explicit",
  "",
  "Public Sub NamesNothing()",
  "    missingHere = 1",
  "End Sub",
  "",
].join("\r\n");

let plantedIn = null;

try {
  // ---- both files are open, and both are the runner's ----
  const listed = await api.tests();
  const named = listed.files.map((file) => file.file).sort();
  // NOT a count: a file with no tests is listed only while the developer is IN it, so a session
  // with the plain fixture beside the pair has two files or three depending on which one the
  // editor adopted at launch. What must hold is that neither test-bearing file is missing.
  check("the runner sees both open files, not just the front one",
    named.includes(MAIN) && named.includes(TWIN), named.join(", "));

  if (!named.includes(MAIN) || !named.includes(TWIN)) {
    throw new Error(`this suite needs ${MAIN} and ${TWIN} open together; it found ${named.join(", ") || "neither"}`);
  }

  const inMain = listed.rows.filter((row) => row.file === MAIN);
  const inTwin = listed.rows.filter((row) => row.file === TWIN);
  check("every row says which file it came from", listed.rows.every((row) => row.file), `${listed.rows.length} rows`);
  check("each file's own tests are discovered", inMain.length === 10 && inTwin.length === 5,
    `${inMain.length} in ${MAIN}, ${inTwin.length} in ${TWIN}`);

  // The pair exists FOR this: both fixtures hold a module called InvoiceTests.
  check("the fixtures collide on a module name, so the questions below are not vacuous",
    inMain.some((row) => row.module === "InvoiceTests") && inTwin.some((row) => row.module === "InvoiceTests"),
    "New-TestTwinFixture.ps1 builds the collision on purpose");

  // ---- a run files each result against the file it came from ----
  const ran = await api.tests({ action: "run", timeoutMs: 240000 });
  check("a run reaches every test in every open file",
    ran.detail === "ran 15 across 2 files", ran.detail);

  const mainRed = ran.rows.find((row) => row.file === MAIN && row.procedure === "Rounding_KnownWrong");
  const twinRed = ran.rows.find((row) => row.file === TWIN && row.procedure === "Handling_KnownWrong");
  check("each file's deliberate red fails in its own words - a result filed by name alone would swap them",
    mainRed?.message?.includes("rounding drifted") === true
    && twinRed?.message?.includes("handling fee moved") === true,
    `${mainRed?.message} / ${twinRed?.message}`);
  check("the twin's expected failure is an xfail, not a red",
    ran.rows.find((row) => row.file === TWIN && row.procedure === "Balance_RefusesOverdraft")?.status === "xfail");

  // ---- one file at a time ----
  const onlyTwin = await api.tests({ action: "run", file: TWIN, timeoutMs: 120000 });
  check("a file-scoped run runs that file alone", onlyTwin.detail === `ran 5 in ${TWIN}`, onlyTwin.detail);
  check("...and the other file's results are still standing, not blanked by the rerun",
    onlyTwin.rows.filter((row) => row.file === MAIN && row.status !== "none").length === 10);

  const named404 = await api.tests({ action: "run", file: "NoSuchBook.xlsm" });
  check("a file nobody has open is an answer, not a run of whatever was in front",
    named404.detail === "no open file called NoSuchBook.xlsm", named404.detail);

  // A module name that lives in both files, run without saying which: it runs in both, because
  // both are what the name means.
  const bothInvoices = await api.tests({ action: "run", module: "InvoiceTests", timeoutMs: 180000 });
  check("a module name held by two files runs in both when neither is named",
    bothInvoices.detail === "ran 13 across 2 files", bothInvoices.detail);

  const oneInvoice = await api.tests({ action: "run", module: "InvoiceTests", file: TWIN, timeoutMs: 120000 });
  check("naming the file picks which copy of a shared module name to run",
    oneInvoice.detail === `ran 3 in ${TWIN}`, oneInvoice.detail);

  // ---- support is per file ----
  await api.component("remove", { name: "XlideAssert", project: TWIN });
  const halved = await waitFor("the twin to read as missing its support module", async () => {
    const now = await api.tests();
    return now.files.find((file) => file.file === TWIN)?.support === "missing" ? now : null;
  }, { budgetMs: 20000 });
  check("support is read per file, not once for the session",
    halved.files.find((file) => file.file === MAIN)?.support === "installed",
    JSON.stringify(halved.files));
  check("the session's one-word standing is the worst of the files that hold tests",
    halved.support === "missing", halved.support);

  const refused = await api.tests({ action: "run", timeoutMs: 60000 });
  check("a run that would touch the file without support refuses, and names it",
    refused.detail.includes(TWIN) && refused.detail.includes("not installed"), refused.detail);

  const allowed = await api.tests({ action: "run", file: MAIN, timeoutMs: 120000 });
  check("...while the file that HAS support still runs - one file's gap is not the other's",
    allowed.detail === `ran 10 in ${MAIN}`, allowed.detail);

  const installed = await api.tests({ action: "install", file: TWIN });
  check("installing into one named file fixes that file", installed.detail === `${TWIN}: installed`, installed.detail);

  // ---- the Tests pane's two selects ----
  await showPane("tests");
  await ask("(() => { for (const one of document.querySelectorAll('.scope-select')) {"
    + " one.value = one.classList.contains('scope-select-file') ? '@allfiles' : '@all';"
    + " one.dispatchEvent(new Event('change')); } return 1; })()");
  const offered = await waitFor("the file select to offer both open files", async () => {
    const shape = await ask(
      '(() => [...document.querySelectorAll("#tests-scope-file option")].map(o => o.textContent))()');
    // Not a count: the plain fixture beside the pair is offered too while the editor is in it.
    return Array.isArray(shape) && shape.some((one) => one.startsWith(`${MAIN} (10)`))
      && shape.some((one) => one.startsWith(`${TWIN} (5)`)) ? shape : null;
  }, { budgetMs: 20000 });
  check("the file select offers All Files and each open file, with its own count",
    offered[0].startsWith("All Files (15)")
    && offered.some((one) => one.startsWith(`${MAIN} (10)`))
    && offered.some((one) => one.startsWith(`${TWIN} (5)`)), offered.join(" | "));

  const across = await ask(
    '(() => [...document.querySelectorAll("#tests-scope-module option")].map(o => o.textContent))()');
  check("across files a shared module name says which file it is, or the two would be one choice",
    across.filter((one) => one.startsWith("InvoiceTests - ")).length === 2, across.join(" | "));

  // Objects come back as objects: the door parses a script's JSON answer on the way out, and a
  // second parse here reads "[object Object]" (the trap docs\driving-excel.md names).
  const tree = await ask(
    '(() => ({files:[...document.querySelectorAll("#tests-list .tests-file")].map(n=>n.textContent),'
    + ' modules:[...document.querySelectorAll("#tests-list .tests-module")].map(n=>n.textContent)}))()');
  check("the tree heads each file's block with the file, then its modules",
    tree.files.length === 2 && tree.modules.length === 3, JSON.stringify(tree));

  await pickScope("tests-scope-file", TWIN);
  const only = await ask(
    '(() => ({rows:document.querySelectorAll("#tests-list .tests-row").length,'
    + ' files:[...document.querySelectorAll("#tests-list .tests-file")].map(n=>n.textContent),'
    + ' modules:[...document.querySelectorAll("#tests-scope-module option")].map(o=>o.textContent),'
    + ' label:document.querySelector("#tests-run .tests-label").textContent,'
    + ' title:document.querySelector("#tests-run").title}))()');
  check("choosing a file shows that file's tests alone", only.rows === 5, JSON.stringify(only));
  check("...and drops the file heading, which would then be saying one thing over and over",
    only.files.length === 0);
  check("the module select narrows to that file's modules, unqualified because they cannot collide",
    only.modules.some((one) => one.startsWith("InvoiceTests (3)"))
    && only.modules.some((one) => one.startsWith("LedgerTests (2)"))
    && !only.modules.some((one) => one.includes(" - ")), only.modules.join(" | "));
  check("Run All becomes Run File, and says which", only.label === "Run File" && only.title.includes(TWIN),
    only.title);

  // ---- the Problems pane's two selects ----
  for (const where of [MAIN, TWIN]) {
    await api.component("add", { kind: "module", name: BROKEN, project: where });
    await api.writeModule(BROKEN, BROKEN_SOURCE, where);
  }

  plantedIn = [MAIN, TWIN];
  await showPane("problems");
  // Waited on the COUNTS, not the names: the file select lists every open file whether or not
  // it has findings, so a file appearing in it says nothing about whether its planted module
  // has been analysed yet.
  const problemFiles = await waitFor("both files' findings to be counted in the Problems file select", async () => {
    const shape = await ask(
      '(() => [...document.querySelectorAll("#problems-scope-file option")].map(o => o.textContent))()');
    const counted = (name) => shape?.some((one) => one.startsWith(name) && !one.endsWith("(0)"));
    return Array.isArray(shape) && counted(MAIN) && counted(TWIN) ? shape : null;
  }, { budgetMs: 40000 });
  // NOT a length check: the developer may have anything else open beside the pair, and a suite
  // that failed because a scratch workbook was up would be testing the session, not the product.
  check("the Problems pane offers a file select of its own, counting each file's findings",
    problemFiles[0].startsWith("All Files")
    && problemFiles.some((one) => one.startsWith(MAIN) && !one.endsWith("(0)"))
    && problemFiles.some((one) => one.startsWith(TWIN) && !one.endsWith("(0)")),
    problemFiles.join(" | "));

  const picked = await pickScope("problems-scope-file", TWIN);
  check("a whole-file problem scope is offered", typeof picked === "string" && picked.startsWith("file:"), String(picked));
  const rows = await ask(
    '(() => [...new Set([...document.querySelectorAll("#panel-list .row")].map(r => r.dataset.project))])()');
  check("...and it shows that file's findings alone",
    Array.isArray(rows) && rows.length === 1 && rows[0] === TWIN, JSON.stringify(rows));

  /*
   * ---- a file with NOTHING TO SAY is still the file the developer is in ----
   *
   * The pane's cache had two writers - the COM walk and the analysis pass - and they disagreed
   * about a file holding neither tests nor XlideAssert: the walk kept it while it was active,
   * the pass dropped it outright. One disagreement, three symptoms, none of them reachable in a
   * session where every open file holds tests (all measured 2026-08-21):
   *
   *   the file being worked in was absent from the pane, so there was no install to press and
   *     the chip spoke for a different file entirely;
   *   a file a walk had listed VANISHED the moment its text moved, mid-typing;
   *   the reconcile behind the tree asks whether every open file is known, and for this file
   *     the answer could never become yes - so every tree publish walked every module of every
   *     project again (95ms per module add with every file known, 228-412ms without).
   */
  await showPane("tests");
  const plain = await api.project(PLAIN);
  const plainModule = plain.components.find((one) => one.kind === "module")?.name;
  check(`${PLAIN} is open beside the pair, with a standard module and no tests of its own`,
    typeof plainModule === "string", JSON.stringify(plain.components?.map((one) => one.name)));

  await api.caret(1, { module: plainModule, project: PLAIN });
  const spell = (answer) => answer.files.map((one) => `${one.file}:${one.support}:${one.tests}`).join(" | ");
  const inPlain = await waitFor(`${PLAIN} to be the pane's own file`, async () => {
    const now = await api.tests();
    return now.files.some((one) => one.file === PLAIN) ? now : null;
  }, { budgetMs: 20000 });
  const standing = inPlain.files.find((one) => one.file === PLAIN);
  check("the file being worked in is listed even with nothing to say, so an install has somewhere to land",
    standing.support === "missing" && standing.tests === 0, spell(inPlain));

  const walked = await api.tests({ action: "refresh" });
  const cached = await api.tests();
  check("a refresh and a repaint answer the same files - one store, one rule about what is shown",
    spell(walked) === spell(cached), `${spell(walked)}  vs  ${spell(cached)}`);

  // Its TEXT MOVES: a module added is a change the analysis pass reads, which is the pass that
  // used to drop the file on its way past.
  await api.component("add", { kind: "module", name: BROKEN, project: PLAIN });
  plantedIn = [...(plantedIn ?? []), PLAIN];
  // Read until it STOPS MOVING, rather than once: a check that asked the moment the add
  // returned would pass on the answer standing before the pass had been anywhere near it.
  const settled = await waitUntilStable(async () => spell(await api.tests()),
    { quiet: 4, pollMs: 400, budgetMs: 25000 });
  check("...and it is still listed once the pass that read the change has been past",
    settled.includes(PLAIN), settled);

  const askedPlain = await api.tests({ action: "run", file: PLAIN });
  check("an action naming it answers about that file, rather than refusing to find it",
    askedPlain.detail === `no tests to run in ${PLAIN}`, askedPlain.detail);

  // The install the developer came for: scoped to their own file, the chip is that file's.
  await pickScope("tests-scope-file", PLAIN);
  const chip = await ask(
    '(() => ({install:document.querySelector("#tests-install")?.textContent,'
    + ' empty:document.querySelector("#tests-list .tests-empty")?.textContent}))()');
  check("scoped to it, the chip offers ITS install and the list says where the tests are instead",
    chip.install === "Install XlideAssert" && chip.empty?.includes(`No tests in ${PLAIN}`),
    JSON.stringify(chip));
} finally {
  await ask("(() => { for (const one of document.querySelectorAll('.scope-select')) {"
    + " one.value = one.classList.contains('scope-select-file') ? '@allfiles' : '@all';"
    + " one.dispatchEvent(new Event('change')); } return 'reset'; })()").catch(() => {});
  for (const where of plantedIn ?? []) {
    await api.component("remove", { name: BROKEN, project: where }).catch(() => {});
  }

  // The twin's support module goes back however the run ended: the fixture is saved with it.
  await api.tests({ action: "install", file: TWIN }).catch(() => {});
  process.exitCode = done();
}
