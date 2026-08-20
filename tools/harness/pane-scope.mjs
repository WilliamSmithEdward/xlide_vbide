/*
 * The scope selector, on both list panes.
 *
 * Both panes answer for a whole workspace - the analyzer reads every module of every open
 * project, the runner discovers every test in the project - and the selector narrows the VIEW
 * without narrowing what the product knows. That distinction is the whole feature, and it is
 * where a defect would hide: a narrowed pane that also narrowed the tree badges, or a Run All
 * that ran tests the pane was not showing, would each be a surface disagreeing with itself.
 *
 * Driven through the page, because the scope is the page's own state - the host is not told
 * about it, and must not be: an editor that phoned home about which module a developer is
 * looking at would be an editor that reanalyses when they scroll.
 *
 * Brings its own modules and takes them away, so it never edits a fixture's code and runs
 * against whatever workbook is open.
 *
 *   node tools\harness\pane-scope.mjs
 */
import { open, reporter, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();

// Unique per run: a suite that reuses names inherits the previous run's answers under them.
const RUN = process.pid % 10000;
const BROKEN = `ScopeBroken${RUN}`;
const ALSO = `ScopeAlso${RUN}`;
const CLEAN = `ScopeClean${RUN}`;
const TESTS = `ScopeTests${RUN}`;

const BROKEN_SOURCE = [
  "Option Explicit",
  "",
  "Public Sub OneComplaint()",
  "    missingHere = 1",
  "End Sub",
  "",
].join("\r\n");

const ALSO_SOURCE = [
  "Option Explicit",
  "",
  "Public Sub TwoComplaints()",
  "    alsoMissing = 1",
  "    andAgain = 2",
  "End Sub",
  "",
].join("\r\n");

const CLEAN_SOURCE = [
  "Option Explicit",
  "",
  "Public Sub NothingWrong()",
  "    Dim held As Long",
  "    held = 1",
  "End Sub",
  "",
].join("\r\n");

const TESTS_SOURCE = [
  "' @xlide-test",
  "Public Sub ScopedPasses()",
  "    XlideAssert.IsTrue True",
  "End Sub",
  "",
  "' @xlide-test",
  "Public Sub ScopedAlsoPasses()",
  "    XlideAssert.AreEqual 2, 1 + 1",
  "End Sub",
].join("\r\n");

/** Runs a script in the page and answers what it returned, already unwrapped. */
const ask = (script) => api.ask(script);

/** Shows one tool pane by pressing its own tab, the way the developer reaches it. */
const showPane = (which) =>
  ask(`document.querySelector('.panel-tab[data-panel="${which}"]')?.click(), "shown"`);

/**
 * The option that offers one module, by what it READS. Its value is the pane's own idea of a
 * module's identity - the Problems pane files by workbook and module, the Tests pane by module
 * alone, because tests are the active project's - and a suite that hard-coded either shape
 * would be testing its own guess.
 */
const optionFor = (select, module) => ask(
  `(() => [...document.querySelectorAll("#${select} option")]`
  + `.filter(o => o.textContent.startsWith(${JSON.stringify(module + " (")}))`
  + ".map(o => o.value))()");

/** Picks a scope, through the change event the control itself fires. */
const pickScope = (select, value) => ask(
  `(() => { const s = document.querySelector("#${select}");`
  + ` s.value = ${JSON.stringify(value)};`
  + ` s.dispatchEvent(new Event("change")); return s.value; })()`);

const made = [];

try {
  // ---- the Problems pane ----
  for (const [name, source] of [[BROKEN, BROKEN_SOURCE], [ALSO, ALSO_SOURCE], [CLEAN, CLEAN_SOURCE]]) {
    const answer = await api.component("add", { kind: "module", name, project: project.projectId });
    if (answer?.ok) {
      made.push(answer.name ?? name);
    }

    await api.writeModule(name, source, project.projectId);
  }

  check("the suite's three modules were added", made.length === 3, made.join(", ") || "none");

  await showPane("problems");
  const brokenOption = await waitFor("both broken modules to reach the scope selector", async () => {
    const mine = await optionFor("problems-scope", BROKEN);
    const other = await optionFor("problems-scope", ALSO);
    return mine?.length === 1 && other?.length === 1 ? mine[0] : null;
  }, { budgetMs: 30000 });

  check("a module with findings earns an option; a clean one does not",
    (await optionFor("problems-scope", CLEAN)).length === 0);

  // The counts on the options are the pane's own arithmetic; the door's problems route is the
  // independent answer they have to match.
  const trueCount = (await api.problems(BROKEN)).findings.length;
  const labelled = await ask(
    `(() => [...document.querySelectorAll("#problems-scope option")]`
    + `.find(o => o.value === ${JSON.stringify(brokenOption)})?.textContent)()`);
  check("the option carries the module's own count",
    labelled === `${BROKEN} (${trueCount})`, `${labelled} against ${trueCount} from the door`);

  await pickScope("problems-scope", brokenOption);
  const scoped = await ask(
    '(() => ({ modules: [...new Set([...document.querySelectorAll("#panel-list .row")]'
    + '.map(r => r.dataset.moduleName))], counts: [...document.querySelectorAll('
    + '"#problems-filters .filter-count")].map(n => n.textContent),'
    + ' narrowed: document.querySelector("#problems-scope").classList.contains("scope-narrowed") }))()');
  check("a scoped pane lists that module alone",
    scoped.modules.length === 1 && scoped.modules[0] === BROKEN, JSON.stringify(scoped.modules));
  check("the severity counts count within the scope, or they contradict the list",
    scoped.counts.join(" ").includes(`${trueCount} `), scoped.counts.join(", "));
  check("a narrowed selector says so, so a hidden row is never read as a clean project",
    scoped.narrowed === true);

  // The tree badges keep the WHOLE session while the panel shows one module: a scope is a
  // view, not a change to what the product knows. The workbook is expanded first, because a
  // collapsed one renders no rows at all and "no badge" would then mean "not looked at".
  await api.act("expandWorkbook", { workbook: project.projectId }).catch(() => {});
  const badge = await waitFor("the tree to show the module the panel is scoped away from", async () => {
    const found = await ask(
      `(() => { const item = [...document.querySelectorAll("#sidebar-tree .tree-item")]`
      + `.find(n => n.textContent.includes(${JSON.stringify(ALSO)}));`
      + ` return item ? item.textContent : null; })()`);
    return typeof found === "string" ? found : null;
  }, { budgetMs: 15000 }).catch(() => null);
  check("the tree still badges a module the panel is scoped away from",
    typeof badge === "string" && /\d/.test(badge), badge ?? "not in the tree");

  // ---- the empty state, and its way back ----
  await api.pane("open", { module: CLEAN, project: project.projectId });
  await pickScope("problems-scope", "@current");
  const empty = await waitFor("the empty state for a clean module", async () => {
    const said = await ask('(() => document.querySelector("#panel-list .panel-empty")?.textContent ?? null)()');
    return typeof said === "string" && said.includes(CLEAN) ? said : null;
  }, { budgetMs: 15000 });
  check("a scope with nothing in it says which module, not nothing at all",
    empty.startsWith(`No problems in ${CLEAN}.`) && empty.includes("elsewhere"), empty);

  const back = await ask(
    '(() => { document.querySelector("#panel-list .panel-empty-act").click();'
    + ' return { scope: document.querySelector("#problems-scope").value,'
    + ' rows: document.querySelectorAll("#panel-list .row").length }; })()');
  check("Show All is the way back out of a scope", back.scope === "@all" && back.rows > 0,
    JSON.stringify(back));

  // ---- auto-adjust: a module that leaves takes its option with it ----
  await pickScope("problems-scope", brokenOption);
  await api.component("remove", { name: BROKEN, project: project.projectId });
  made.splice(made.indexOf(BROKEN), 1);
  const settled = await waitFor("the scope to notice the module has gone", async () => {
    const now = await ask(
      '(() => ({ value: document.querySelector("#problems-scope").value,'
      + ' options: [...document.querySelectorAll("#problems-scope option")].map(o => o.value) }))()');
    return now && !now.options.includes(brokenOption) ? now : null;
  }, { budgetMs: 30000 });
  check("a scope whose module has gone falls back to All rather than filtering against nothing",
    settled.value === "@all", JSON.stringify(settled));

  // ---- the Tests pane ----
  const support = await api.tests({ action: "install" });
  check("XlideAssert is installed for the scoped run", support.support === "installed", support.support);

  await api.component("add", { kind: "module", name: TESTS, project: project.projectId });
  made.push(TESTS);
  await api.writeModule(TESTS, TESTS_SOURCE, project.projectId);
  await waitFor("the runner to discover the scoped tests", async () =>
    (await api.tests()).rows.filter((row) => row.module === TESTS).length === 2 ? true : null,
  { budgetMs: 20000 });

  await showPane("tests");
  const testsOption = await waitFor("the tests scope to offer the module", async () => {
    const mine = await optionFor("tests-scope", TESTS);
    return mine?.length === 1 ? mine[0] : null;
  }, { budgetMs: 15000 });
  await pickScope("tests-scope", testsOption);
  const testScoped = await waitFor("the tests pane to narrow", async () => {
    const now = await ask(
      '(() => ({ modules: [...document.querySelectorAll("#tests-list .tests-module")].map(n => n.textContent),'
      + ' rows: document.querySelectorAll("#tests-list .tests-row").length,'
      + ' runLabel: document.querySelector("#tests-run .tests-label").textContent,'
      + ' runTitle: document.querySelector("#tests-run").title,'
      + ' summary: document.querySelector("#tests-summary").textContent }))()');
    return now && now.modules.length === 1 ? now : null;
  }, { budgetMs: 15000 });
  check("a scoped tests pane shows that module's heading alone",
    testScoped.modules[0] === TESTS && testScoped.rows === 2, JSON.stringify(testScoped.modules));
  check("Run All becomes Run Module when the pane is scoped, and names it",
    testScoped.runLabel === "Run Module" && testScoped.runTitle.includes(TESTS), testScoped.runTitle);
  check("the tally counts what is showing, not what exists",
    testScoped.summary.endsWith("of 2"), testScoped.summary);

  // The door mirrors the button: runFailed narrowed to a module answers about that module,
  // never about failures the pane has scoped out of sight.
  const narrowedRerun = await api.tests({ action: "runFailed", module: TESTS, timeoutMs: 60000 });
  check("runFailed with a module answers about that module alone",
    narrowedRerun.detail === `nothing has failed in ${TESTS}`, narrowedRerun.detail);

  // ---- the run clock ----
  const before = Date.now();
  const ran = await api.tests({ action: "run", module: TESTS, timeoutMs: 120000 });
  const stamped = ran.ranAt === null ? NaN : new Date(ran.ranAt).getTime();
  check("a finished run stamps when it landed",
    Number.isFinite(stamped) && stamped >= before - 60_000 && stamped <= Date.now() + 60_000,
    String(ran.ranAt));

  const said = await waitFor("the pane to say when the run landed", async () => {
    const now = await ask('(() => document.querySelector("#tests-ran")?.textContent ?? null)()');
    return typeof now === "string" && now.startsWith("Ran ") ? now : null;
  }, { budgetMs: 15000 });
  check("the pane says it in words, beside the tally", /^Ran \d/.test(said), said);
} finally {
  await ask('(() => { const s = document.querySelector("#tests-scope"); if (s) { s.value = "@all"; '
    + 's.dispatchEvent(new Event("change")); } const p = document.querySelector("#problems-scope"); '
    + 'if (p) { p.value = "@all"; p.dispatchEvent(new Event("change")); } return "reset"; })()')
    .catch(() => {});
  for (const name of made) {
    await api.component("remove", { name, project: project.projectId }).catch(() => {});
  }

  process.exitCode = done();
}
