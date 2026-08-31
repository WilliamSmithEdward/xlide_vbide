/*
 * The VBA test runner, driven end to end through the door: install the support module, write
 * a module of tests covering every outcome the runner can answer, run them, and hold each
 * status to what the directives promised. The pane is fed by the same snapshot this route
 * answers, so what passes here is what the panel shows.
 *
 *   node tools\harness\test-runner.mjs
 */
import { open, reporter, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const name = `TestProbe${process.pid % 10000}`;
const CRLF = "\r\n";

const MODULE = [
  "' @xlide-test tags=smoke",
  "Public Sub Adds()",
  "    XlideAssert.AreEqual 4, 2 + 2",
  "End Sub",
  "",
  "' A directive block can carry metadata, and an aside after -- is never metadata.",
  "' @xlide-test tags=\"smoke,math\" owner=william timeout=5s -- flaky on Tuesdays",
  "Public Sub FailsOnPurpose()",
  "    XlideAssert.AreEqual 5, 2 + 2, \"arithmetic drifted\"",
  "    XlideAssert.WriteLine \"still ran to the end\"",
  "End Sub",
  "",
  "' @xlide-test",
  "Public Sub Raises()",
  "    Err.Raise 13",
  "End Sub",
  "",
  "' @xlide-test-skip reason=\"not on this machine\"",
  "Public Sub SkipsAlways()",
  "    XlideAssert.Fail \"never runs\"",
  "End Sub",
  "",
  "' @xlide-test-xfail reason=\"known defect 12\"",
  "Public Sub KnownBroken()",
  "    XlideAssert.IsTrue False",
  "End Sub",
  "",
  "' @xlide-test-xfail reason=\"was fixed and nobody noticed\"",
  "Public Sub SecretlyFixed()",
  "    XlideAssert.IsTrue True",
  "End Sub",
  "",
  "' @xlide-test expected-error=13",
  "Public Sub WantsTypeMismatch()",
  "    Err.Raise 13",
  "End Sub",
  "",
  "' @xlide-test expected-error=13",
  "Public Sub WantsButGetsAnother()",
  "    Err.Raise 9",
  "End Sub",
  "",
  "' @xlide-test expected-error",
  "Public Sub AnyErrorWillDo()",
  "    Err.Raise 76",
  "End Sub",
  "",
  "' @xlide-test",
  "Public Sub ThrowsHelper()",
  "    XlideAssert.Throws 13, \"MustMismatch\"",
  "End Sub",
  "",
  "Public Sub MustMismatch()",
  "    Err.Raise 13",
  "End Sub",
  "",
  "' @xlide-test",
  "Private Sub NotDiscovered()",
  "End Sub",
  "",
  "' @xlide-tost",
  "Public Sub TypoedDirective()",
  "End Sub",
].join("\r\n");

try {
  // ---- the support module's install gate ----
  const installed = await api.tests({ action: "install" });
  check("installing XlideAssert answers installed", installed.support === "installed", installed.support);

  const refused = await api.tests({ action: "run" }).catch((error) => ({ error: String(error) }));
  check("a run with no tests still answers rather than failing",
    refused.detail !== undefined || refused.error !== undefined);

  // ---- discovery ----
  await api.component("add", { kind: "module", name, project: project.projectId });
  await api.writeModule(name, MODULE, project.projectId);
  const listed = await waitFor("discovery to see the module", async () => {
    const now = await api.tests();
    return now.rows.filter((row) => row.module === name).length >= 10 ? now : null;
  }, { budgetMs: 15000 });
  const mine = listed.rows.filter((row) => row.module === name);

  check("discovery finds every directive-marked public Sub", mine.length === 10,
    mine.map((row) => row.procedure).join(", "));
  check("a Private Sub under a directive is not offered - the runner could never call it",
    !mine.some((row) => row.procedure === "NotDiscovered"));
  check("a typoed directive discovers nothing", !mine.some((row) => row.procedure === "TypoedDirective"));
  check("an unmarked helper Sub is not a test", !mine.some((row) => row.procedure === "MustMismatch"));

  // AUTO-REDISCOVERY: the pane heard about this module from the analysis pass alone - the
  // write above provoked a pass, the pass handed its snapshot over, and the rows repainted
  // with no tests() call and no refresh press. Read off the pane's own DOM, which paints
  // whether or not the pane is showing.
  const autoHeard = await waitFor("the pane to hear the new tests on its own", async () =>
    (await api.ask(`document.querySelector("#tests-list")?.textContent.includes(${JSON.stringify(name)}) === true`)) === true,
  { budgetMs: 20000 }).then(() => true).catch(() => false);
  check("the pane rediscovers on its own when code changes - no refresh pressed", autoHeard);

  const tagged = mine.find((row) => row.procedure === "FailsOnPurpose");
  check("metadata rides discovery: tags, owner, and a timeout in seconds",
    tagged?.tags.join(",") === "smoke,math" && tagged?.owner === "william" && tagged?.timeoutMs === 5000,
    JSON.stringify({ tags: tagged?.tags, owner: tagged?.owner, timeoutMs: tagged?.timeoutMs }));
  check("a declaration line is where the row navigates to",
    (mine.find((row) => row.procedure === "Adds")?.line ?? 0) === 2,
    `Adds is at line ${mine.find((row) => row.procedure === "Adds")?.line}`);

  /* ---- the facets, on a board where nothing has run --------------------------------------
   *
   * tags= and outcomes= narrow any run verb the way the pane's filters narrow its list, and
   * the pane's Run Displayed button is exactly run + the facets - so what these pin is the
   * selection both doors share. Run before the full suite on purpose: only a board where
   * nothing has run can tell outcomes=notRun from "everything".
   */
  const smokeRun = await api.tests({ action: "run", module: name, tags: "smoke", timeoutMs: 120000 });
  const smokeOf = (procedure) => smokeRun.rows.find((row) => row.module === name && row.procedure === procedure);
  check("run with tags= runs only the tagged tests - any of the listed",
    /^ran 2 in /.test(smokeRun.detail), smokeRun.detail);
  check("...the smoke pair landed", smokeOf("Adds")?.status === "passed" && smokeOf("FailsOnPurpose")?.status === "failed",
    `${smokeOf("Adds")?.status}/${smokeOf("FailsOnPurpose")?.status}`);
  check("...and an untagged test was not touched", smokeOf("Raises")?.status === "none", smokeOf("Raises")?.status);

  const combo = await api.tests({ action: "run", module: name, tags: "math", outcomes: "failed", timeoutMs: 60000 });
  check("tags= and outcomes= compose: math AND currently-failed is one test",
    /^ran 1 in /.test(combo.detail), combo.detail);

  const untouched = await api.tests({ action: "run", module: name, outcomes: "notRun", timeoutMs: 180000 });
  check("outcomes=notRun runs what has never run - the skip-marked row is skipped, not notRun",
    /^ran 7 in /.test(untouched.detail), untouched.detail);

  const skipOnly = await api.tests({ action: "run", module: name, tags: "untagged", outcomes: "skipped", timeoutMs: 60000 });
  check("untagged is a tag word, and skip-marked answers to the skipped group",
    /^ran 1 in /.test(skipOnly.detail)
    && skipOnly.rows.find((row) => row.id === `${name}.SkipsAlways`)?.status === "skipped",
    skipOnly.detail);

  const noSuchTag = await api.tests({ action: "run", module: name, tags: "nosuch" });
  check("an unknown tag is an empty selection that names itself",
    noSuchTag.detail === `no tests in ${name} with tag nosuch`, noSuchTag.detail);

  const strayGroup = await api.tests({ action: "run", module: name, outcomes: "bogus" });
  check("a stray outcome word is refused as a typo, never run as nothing",
    /'bogus' is not an outcome group/.test(strayGroup.detail), strayGroup.detail);

  const fileCombo = await api.tests({
    action: "run", module: name, file: project.projectId, tags: "smoke", outcomes: "passed", timeoutMs: 60000,
  });
  check("the facets compose with file= - workbook-aware to the end",
    /^ran 1 in /.test(fileCombo.detail), fileCombo.detail);

  // ---- the run, every outcome at once ----
  const ran = await api.tests({ action: "run", module: name, timeoutMs: 180000 });
  const status = (procedure) => ran.rows.find((row) => row.module === name && row.procedure === procedure);

  check("a passing test passes", status("Adds")?.status === "passed", JSON.stringify(status("Adds")));
  check("a failed assertion fails with the latched message, user text first",
    status("FailsOnPurpose")?.status === "failed"
    && /arithmetic drifted Expected <5> but was <4>\./.test(status("FailsOnPurpose")?.message ?? ""),
    status("FailsOnPurpose")?.message);
  check("...and the WriteLine after the failed assertion still arrived - assertions are not fatal",
    (status("FailsOnPurpose")?.output ?? []).includes("still ran to the end"),
    JSON.stringify(status("FailsOnPurpose")?.output));
  check("a raised error fails with its number and source",
    status("Raises")?.status === "failed" && /VBA error 13/.test(status("Raises")?.message ?? ""),
    status("Raises")?.message);
  check("a skip directive skips without running - its Fail never latched",
    status("SkipsAlways")?.status === "skipped" && status("SkipsAlways")?.message === "not on this machine",
    JSON.stringify(status("SkipsAlways")));
  check("an expected failure reads xfail, not red",
    status("KnownBroken")?.status === "xfail", status("KnownBroken")?.status);
  check("an expected failure that PASSES reads xpass, which is a red worth seeing",
    status("SecretlyFixed")?.status === "xpass"
    && /Expected failure did not occur/.test(status("SecretlyFixed")?.message ?? ""),
    JSON.stringify(status("SecretlyFixed")));
  check("expected-error=13 passes when 13 is raised",
    status("WantsTypeMismatch")?.status === "passed", status("WantsTypeMismatch")?.status);
  check("expected-error=13 fails when 9 arrives, naming both",
    status("WantsButGetsAnother")?.status === "failed"
    && /Expected VBA error 13, but got VBA error 9/.test(status("WantsButGetsAnother")?.message ?? ""),
    status("WantsButGetsAnother")?.message);
  check("a bare expected-error takes any error at all",
    status("AnyErrorWillDo")?.status === "passed", status("AnyErrorWillDo")?.status);
  check("Assert.Throws reaches its target through the staged dispatcher",
    status("ThrowsHelper")?.status === "passed", JSON.stringify(status("ThrowsHelper")));

  // ---- one test alone, and the failed set ----
  const one = await api.tests({ action: "run", test: `${name}.Adds`, timeoutMs: 60000 });
  // The answer says how many ran and in which file, because a session holds more than one.
  check("run with test= runs exactly that test",
    one.rows.find((row) => row.id === `${name}.Adds`)?.status === "passed"
    && /^ran 1 in /.test(one.detail), one.detail);

  const failedAgain = await api.tests({ action: "runFailed", timeoutMs: 120000 });
  const rerunIds = [];
  for (const row of failedAgain.rows) {
    if (row.module === name && row.status !== "none" && row.durationMs > 0
      && ["failed", "error", "xpass"].includes(row.status)) {
      rerunIds.push(row.procedure);
    }
  }

  check("runFailed reruns the failing set and it fails the same way",
    failedAgain.rows.find((row) => row.id === `${name}.FailsOnPurpose`)?.status === "failed"
    && failedAgain.rows.find((row) => row.id === `${name}.SecretlyFixed`)?.status === "xpass");

  // ---- the injected modules leave no trace ----
  const components = (await api.tests()).rows;
  check("the generated runner and dispatcher are gone when the run ends",
    !components.some((row) => row.module.startsWith("XlideRun") || row.module === "XlideTestDispatch"));
  const support = await api.tests();
  check("XlideAssert still reads installed after runs - injection re-cases nothing",
    support.support === "installed", support.support);

  /* ---- the pane's tag filter and Run Displayed ----------------------------------------------
   *
   * The pane's Run Displayed is the door's run + tags= + outcomes= by construction; what these
   * pin is the pane's own half - the facet controls, the counts following them, the button's
   * number, and that a press reaches the host wearing the facets AND the file. Everything is
   * read off the screen with the Tests tab fronted, because the summary-button ghost taught
   * that state and screen are two different instruments.
   */
  const paneFile = mine[0].file;
  await api.ask(`document.querySelector('.panel-tab[data-panel="tests"]')?.click(), "shown"`);
  const scopeKey = `${paneFile.toLowerCase()}\u0000${name.toLowerCase()}`;
  await api.ask(`(() => { const s = document.querySelector("#tests-scope-module");`
    + ` s.value = ${JSON.stringify(scopeKey)}; s.dispatchEvent(new Event("change")); return s.value; })()`);

  const pane = async () => {
    const said = await api.ask(`JSON.stringify({
      tagsHidden: document.querySelector(".tag-select")?.hidden ?? null,
      tagLabel: document.querySelector("#tests-tags .tests-label")?.textContent ?? null,
      narrowed: document.querySelector("#tests-tags")?.classList.contains("tag-narrowed") ?? null,
      popupOnScreen: (document.querySelector(".tag-select-popup")?.getBoundingClientRect().width ?? 0) > 0,
      rows: document.querySelectorAll("#tests-list .tests-row").length,
      displayed: document.querySelector("#tests-run-displayed .tests-label")?.textContent ?? null,
      displayedOff: document.querySelector("#tests-run-displayed")?.disabled ?? null,
      chips: [...document.querySelectorAll(".tests-filter .filter-count")].map((one) => one.textContent),
      empty: document.querySelector("#tests-list .tests-empty span")?.textContent ?? null,
    })`);
    return JSON.parse(typeof said === "string" ? said : JSON.stringify(said));
  };

  const press = (selector) => api.ask(
    `(() => { const b = document.querySelector(${JSON.stringify(selector)});`
    + ` if (!b) return "missing"; b.click(); return "pressed"; })()`);

  const scoped = await pane();
  check("scoped to the probe module, the tag filter offers itself",
    scoped.tagsHidden === false && scoped.tagLabel === "Tags", JSON.stringify(scoped));
  check("and Run Displayed counts everything while nothing narrows",
    scoped.displayed === "Run Displayed (10)" && scoped.displayedOff === false,
    `${scoped.displayed} disabled=${scoped.displayedOff}`);

  check("opening the popup puts it on screen", await press("#tests-tags") === "pressed"
    && (await pane()).popupOnScreen === true);
  check("choosing smoke narrows the list to its two tests",
    await press('.tag-select-popup input[data-tag="smoke"]') === "pressed"
    && (await pane()).rows === 2, JSON.stringify(await pane()));

  const smokeFaceted = await pane();
  check("the chip counts follow the tag facet - each chip governs what it can show",
    smokeFaceted.chips.includes("1 Passed") && smokeFaceted.chips.includes("1 Failed"),
    smokeFaceted.chips.join(" | "));
  check("the button wears the facet: a narrowed pane says so",
    smokeFaceted.narrowed === true && smokeFaceted.tagLabel === "Tags: smoke"
    && smokeFaceted.displayed === "Run Displayed (2)",
    JSON.stringify(smokeFaceted));

  // Solo the Failed chip the way a pointer does, then press Run Displayed: the host must
  // receive runModule + the module + THE FILE + tags=smoke + outcomes=failed, and run one.
  await api.ask(`(() => { const b = document.querySelector(".tests-filter-failed");
    b.dispatchEvent(new PointerEvent("pointerdown", { ctrlKey: true, bubbles: true })); return "solo"; })()`);
  const soloed = await pane();
  check("soloing Failed over the smoke facet leaves one displayed",
    soloed.rows === 1 && soloed.displayed === "Run Displayed (1)", JSON.stringify(soloed));

  const beforeRan = (await api.tests()).ranAt;
  await press("#tests-run-displayed");
  await waitFor("the displayed run to land", async () => {
    const now = await api.tests();
    return now.running === false && now.ranAt !== beforeRan ? now : null;
  }, { budgetMs: 120000, pollMs: 500 });
  const logged = await api.log({ match: "outcomes=failed -> ran 1 in", max: 20000 });
  check("Run Displayed pressed the facets and the file through to the host, and ran exactly one",
    logged.lines.some((line) => line.includes(`runModule ${name} in ${paneFile} tags=smoke outcomes=failed -> ran 1 in`)),
    logged.lines.at(-1) ?? "no such log line");

  // A row's tag chip is the same facet: pressing math on the failed row widens it to two
  // tags. The offer is three - smoke, math, and "(untagged)", which counts as a choice
  // wherever untagged tests exist.
  check("a row's tag chip toggles the facet",
    await api.ask(`(() => { const chip = [...document.querySelectorAll("#tests-list .tests-tag")]
      .find((one) => one.textContent === "math"); if (!chip) return "missing"; chip.click(); return "pressed"; })()`) === "pressed"
    && (await pane()).tagLabel === "Tags: 2 of 3", (await pane()).tagLabel);

  // Facets that empty the list say WHICH filters did it, and one press clears them all.
  await api.ask(`(() => { const b = document.querySelector(".tests-filter-skipped");
    b.dispatchEvent(new PointerEvent("pointerdown", { ctrlKey: true, bubbles: true })); return "solo"; })()`);
  const emptied = await pane();
  check("an emptied list names the tag and outcome filters together",
    emptied.rows === 0 && /hidden by the tag and outcome filters/.test(emptied.empty ?? ""),
    JSON.stringify(emptied));
  check("and Clear Filters brings every row back",
    await press("#tests-list .panel-empty-act") === "pressed"
    && (await pane()).rows === 10 && (await pane()).tagLabel === "Tags",
    JSON.stringify(await pane()));

  // The pane goes back to everything before the wrecker takes the stage.
  await api.ask(`(() => { const s = document.querySelector("#tests-scope-module");`
    + ` s.value = "@all"; s.dispatchEvent(new Event("change")); return s.value; })()`);

  /* ---- a project that cannot execute a line -------------------------------------------------
   *
   * VBA compiles the WHOLE project before it runs anything, so one module that does not parse
   * stops every test in every other module. What the pane did then was the worst version of
   * this: it kept the PREVIOUS run's green rows, left one row reading "running" for ever, and
   * refused every later run as already in flight - so a developer who broke the build saw a
   * green Tests pane and a runner that never worked again in that session (#10).
   */
  const wrecker = `${name}Wrecker`;
  await api.component("remove", { name: wrecker, project: project.projectId }).catch(() => {});
  await api.component("add", { kind: "module", name: wrecker, project: project.projectId });
  await api.writeModule(wrecker,
    ["Option Explicit", "", "Public Sub Broken(", "    If Then Else", "    Dim", "End"].join(CRLF),
    project.projectId);

  // THE PRECONDITION, PROVED. "The macro would not run" reads identically for a project that
  // will not compile and one whose macro is simply absent, so the state is established first.
  const cannotRun = await api.immediate("?1+1").catch(() => ({ failed: true }));
  check("nothing in the project can run once one module will not parse",
    cannotRun.failed === true, JSON.stringify(cannotRun.text));

  const wrecked = await api.tests({ action: "run", timeoutMs: 180000 })
    .catch((err) => ({ threw: String(err.message).slice(0, 90) }));
  const after = await waitFor("the run to settle",
    async () => {
      const now = await api.tests();
      return now.running === false ? now : null;
    }, { budgetMs: 120000, pollMs: 1000 });

  const wreckedRows = after.rows.filter((row) => row.module === name);
  check("no test reads passed in a project that cannot execute a line",
    wreckedRows.every((row) => row.status !== "passed"),
    JSON.stringify(wreckedRows.reduce(
      (all, row) => ({ ...all, [row.status]: (all[row.status] ?? 0) + 1 }), {})));
  check("the run releases its in-flight flag rather than wedging the pane",
    after.running === false, `running=${after.running}`);
  check("and the reason names the compile error rather than an HRESULT",
    wreckedRows.some((row) => /compile error/i.test(String(row.message ?? ""))),
    String(wreckedRows.find((row) => row.message)?.message ?? "no message at all"));
  void wrecked;

  // And it comes BACK. A refusal that outlived the cause would be the same wedge by another name.
  await api.component("remove", { name: wrecker, project: project.projectId });
  await waitFor("the project to compile again",
    async () => (await api.immediate("?1+1").catch(() => ({ failed: true }))).failed === false,
    { budgetMs: 30000, pollMs: 1000 });

  const recovered = await api.tests({ action: "run", timeoutMs: 180000 });
  check("and the runner recovers once the module is gone",
    recovered.rows.some((row) => row.module === name && row.status === "passed"),
    recovered.detail);
} finally {
  await api.component("remove", { name: `${name}Wrecker`, project: project.projectId }).catch(() => {});
  await api.component("remove", { name, project: project.projectId }).catch(() => {});
  process.exitCode = done();
}
