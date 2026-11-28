/*
 * The lenses above procedures: Run test and Debug test over a test, and over every procedure a
 * count of the places that use it.
 *
 * What is planned is pinned without Excel, in ui\editor\test\lensplan.mjs. What only a live
 * editor can say is here:
 *
 *   - are the planned lenses actually DRAWN, above the right lines, saying the right words?
 *   - is a count the same number the references list shows for that procedure?
 *   - does pressing a count open that list, and pressing Run test run that one test?
 *   - does a count follow a use added in ANOTHER module, which changes nothing in this one?
 *   - does a lens follow its procedure when the text above it moves?
 *   - does the switch take them all away, and bring them back?
 *   - what does it cost: how many counts are asked of the host, and are they asked twice?
 *   - can the words be read, in both themes?
 *
 * A lens is pressed through `act("lens")`, which presses the drawn anchor through the editor's
 * own pointer handling - a lens that is planned and not drawn cannot be pressed, which is the
 * point of driving it that way.
 *
 * It writes two modules of its own and takes them away again, installs the tests' support
 * module the way the Tests pane's button does (test-runner.mjs does the same and leaves it),
 * and puts the developer's switch back.
 *
 * Run against the debug fixture:
 *   node tools\harness\lenses.mjs
 */

import { open, reporter, scratchModule, wait, waitFor } from "./xlide-api.mjs";
import { contrastByTheme, contrastSaid } from "./contrast.mjs";

const api = await open({});
const { check, done } = reporter();
const project = await api.project();

const stamp = process.pid % 100000;
const probe = `LensProbe${stamp}`;
const user = `LensUser${stamp}`;
const TEST = `Passes${stamp}`;
const TWICE = `Twice${stamp}`;
const UNUSED = `Unused${stamp}`;

// Line numbers are part of what is asserted, so the text is laid out by line here.
const MODULE = [
  "Option Explicit",                                        // 1
  "",                                                       // 2
  "' @xlide-test",                                          // 3
  `Public Sub ${TEST}()`,                                   // 4
  "    Dim total As Long",                                  // 5
  `    total = ${TWICE}(2) + ${TWICE}(3)`,                  // 6
  "End Sub",                                                // 7
  "",                                                       // 8
  `Public Function ${TWICE}(ByVal value As Long) As Long`,  // 9
  `    ${TWICE} = value * 2`,                               // 10
  "End Function",                                           // 11
  "",                                                       // 12
  `Private Sub ${UNUSED}()`,                                // 13
  "End Sub",                                                // 14
].join("\r\n");

// The second module, before and after it gains a use of the first one's function.
const USER_BEFORE = [
  "Option Explicit", "", `Public Function Also${stamp}() As Long`, `    Also${stamp} = 7`, "End Function",
].join("\r\n");
const USER_AFTER = USER_BEFORE.replace("= 7", `= ${TWICE}(7)`);

const lensesNow = async () => (await api.ui()).lenses;
const said = (lenses) => lenses.planned.map((lens) => `${lens.line} ${lens.kind} ${lens.procedure}`);
const titleOf = (lenses, line, kind) => lenses.planned.find((lens) => lens.line === line && lens.kind === kind)?.title ?? null;
const drawnAt = (lenses, line) => lenses.drawn.find((row) => row.line === line)?.titles ?? [];

/** Waits until every planned lens has been asked for and says something, and answers the reading. */
const settled = (what, also = () => true) => waitFor(what, async () => {
  const lenses = await lensesNow();
  return lenses.planned.length > 0 && lenses.planned.every((lens) => lens.title !== null) && also(lenses) ? lenses : null;
}, { budgetMs: 20000, pollMs: 250 });

const switchWas = (await api.settings()).codeLens;
const probeModule = scratchModule(api, project.projectId, probe);
const userModule = scratchModule(api, project.projectId, user);

try {
  console.log("the lenses above procedures\n");

  await api.settings({ codeLens: true });
  const support = await api.tests({ action: "install" });
  check("the tests' support module is installed, so a test can be run", support.support === "installed", support.support);

  // The second module first, with nothing in it yet that uses the first. Adding a module brings
  // it forward, as it does from the Insert menu, and the part below that writes to it needs the
  // module being read to stay on screen the whole time.
  await api.component("add", { kind: "module", name: user, project: project.projectId });
  await api.writeModule(user, USER_BEFORE, project.projectId);

  await api.component("add", { kind: "module", name: probe, project: project.projectId });
  await api.writeModule(probe, MODULE, project.projectId);
  await waitFor("the runner to list the test", async () =>
    (await api.tests()).rows.some((row) => row.module === probe && row.procedure === TEST), { budgetMs: 20000 });
  await api.pane("open", { module: probe, project: project.projectId });

  // ---- planned, and drawn ----
  let lenses = await settled("the lenses to be planned and their counts asked for",
    (now) => now.planned.some((lens) => lens.kind === "runTest"));
  check("the plan is the test's two lenses, then one count for every procedure",
    JSON.stringify(said(lenses)) === JSON.stringify([
      `4 runTest ${TEST}`, `4 debugTest ${TEST}`, `4 references ${TEST}`,
      `9 references ${TWICE}`, `13 references ${UNUSED}`,
    ]), said(lenses).join("; "));
  check("the test's lenses say Run test and Debug test",
    titleOf(lenses, 4, "runTest") === "Run test" && titleOf(lenses, 4, "debugTest") === "Debug test",
    `${titleOf(lenses, 4, "runTest")} / ${titleOf(lenses, 4, "debugTest")}`);

  lenses = await waitFor("the editor to draw them", async () => {
    const now = await lensesNow();
    return now.drawn.length === 3 && now.drawn.every((row) => row.line !== null) ? now : null;
  }, { budgetMs: 15000, pollMs: 250 });
  for (const line of [4, 9, 13]) {
    const planned = lenses.planned.filter((lens) => lens.line === line).map((lens) => lens.title);
    check(`above line ${line} the editor drew exactly what was planned`,
      JSON.stringify(drawnAt(lenses, line)) === JSON.stringify(planned),
      `drawn ${JSON.stringify(drawnAt(lenses, line))}, planned ${JSON.stringify(planned)}`);
  }

  // ---- a count is the references list's number ----
  // The list counts the declaration; the lens does not. Everything else must be the same set.
  const listFor = async (word, headerLine) => {
    const answer = await api.act("references", { word });
    return (answer.data ?? []).filter((one) => !(one.module === probe && one.line === headerLine));
  };
  const twiceUses = await listFor(TWICE, 9);
  check("the function's count is the number of uses the references list shows for it",
    titleOf(lenses, 9, "references") === `${twiceUses.length} references` && twiceUses.length >= 2,
    `lens ${JSON.stringify(titleOf(lenses, 9, "references"))}, list ${twiceUses.length}: ${twiceUses.map((one) => `${one.module}:${one.line}`).join(", ")}`);
  check("a procedure nothing uses says so in words",
    titleOf(lenses, 13, "references") === "no references", String(titleOf(lenses, 13, "references")));
  check("a test nothing calls says so too", titleOf(lenses, 4, "references") === "no references", String(titleOf(lenses, 4, "references")));

  // ---- what it costs ----
  const askedOnce = lenses.countsAsked;
  await wait(1500);
  check("looking again asks the host nothing new", (await lensesNow()).countsAsked === askedOnce,
    `${askedOnce} then ${(await lensesNow()).countsAsked}`);

  // ---- a use added in another module moves the count here ----
  // Written into a module that is NOT the one on screen, and BEFORE any test is run: a run
  // gives the module a new text version, which asks for its counts again all by itself and so
  // hides whether the page ever heard the analysis move (it did hide it: this check stayed green
  // with that hook taken out, until it was moved up here and given the two guards below).
  // Nothing about the module being read changes - not its text, not its version, not which
  // editor shows it - so the only thing that can move its count is that hook.
  const versionNow = () => api.ask("window.xlideBridge.workspace.activeEditor().getModel().getVersionId()");
  const showingBefore = (await api.ui()).statusModule;
  const versionBefore = await versionNow();
  await api.writeModule(user, USER_AFTER, project.projectId);
  const wanted = `${twiceUses.length + 1} references`;
  const moved = await settled("the count to follow a use added elsewhere", (now) => titleOf(now, 9, "references") === wanted)
    .catch(async () => lensesNow());
  check("a use added in another module is counted here, though this module's text never changed",
    titleOf(moved, 9, "references") === wanted, `${titleOf(moved, 9, "references")}, wanted ${wanted}`);
  check("and the editor drew the new count", drawnAt(await lensesNow(), 9).includes(wanted), JSON.stringify(drawnAt(await lensesNow(), 9)));
  const showingAfter = (await api.ui()).statusModule;
  check("and the module being read stayed on screen throughout",
    showingAfter === showingBefore && String(showingAfter).includes(probe), `${showingBefore} then ${showingAfter}`);
  const versionAfter = await versionNow();
  check("and its text version never moved, so nothing but the analysis moving asked again",
    versionAfter === versionBefore, `${versionBefore} then ${versionAfter}`);

  // ---- pressing a count opens the list ----
  const pressedCount = await api.act("lens", { line: 9, title: "references" });
  check("the count is pressed where it is drawn", pressedCount.did === true, pressedCount.detail);
  const listUp = await waitFor("the references list to open", async () =>
    (await api.ui()).dialogs.some((row) => row.id === "references-card"), { budgetMs: 8000 }).then(() => true, () => false);
  check("and that opens xlide's references list", listUp, (await api.ui()).dialogs.map((row) => row.id).join(", "));
  await api.act("closeDialogs");

  // ---- pressing Run test runs that test ----
  const ranBefore = (await api.tests()).ranAt;
  const pressedRun = await api.act("lens", { line: 4, title: "Run test" });
  check("Run test is pressed where it is drawn", pressedRun.did === true, pressedRun.detail);
  const afterRun = await waitFor("the test to have run", async () => {
    const now = await api.tests();
    return !now.running && now.ranAt !== ranBefore ? now : null;
  }, { budgetMs: 60000, pollMs: 300 }).catch(() => null);
  const row = afterRun?.rows.find((one) => one.module === probe && one.procedure === TEST);
  check("and that test ran and passed", row?.status === "passed", `${row?.status} ${row?.message ?? ""}`);
  check("and only that test ran",
    (afterRun?.rows ?? []).filter((one) => one.module !== probe && one.status === "running").length === 0);

  // ---- a lens follows its procedure ----
  await api.writeModule(probe, `' a line above${"\r\n"}${MODULE}`, project.projectId);
  const shifted = await settled("the lenses to follow the text", (now) => now.planned.some((lens) => lens.line === 10))
    .catch(async () => lensesNow());
  check("a line added above moves every lens down with its procedure",
    JSON.stringify([...new Set(shifted.planned.map((lens) => lens.line))]) === JSON.stringify([5, 10, 14]),
    said(shifted).join("; "));
  check("and the count is still the function's", titleOf(shifted, 10, "references") === wanted, String(titleOf(shifted, 10, "references")));

  // ---- can it be read ----
  const readable = await contrastByTheme(api, ".monaco-editor .contentWidgets");
  for (const theme of ["dark", "light"]) {
    check(`the lenses were read in ${theme}`, readable[theme].read >= 5, contrastSaid(readable[theme]));
    check(`and hold AA contrast in ${theme}`, readable[theme].fails.length === 0, contrastSaid(readable[theme]));
  }

  // ---- the switch ----
  await api.settings({ codeLens: false });
  const gone = await waitFor("the lenses to go", async () => {
    const now = await lensesNow();
    return now.planned.length === 0 && now.drawn.length === 0 ? now : null;
  }, { budgetMs: 10000, pollMs: 250 }).catch(async () => lensesNow());
  check("the switch off takes every lens away", gone.planned.length === 0 && gone.drawn.length === 0,
    `${gone.planned.length} planned, ${gone.drawn.length} drawn`);
  const askedWhenOff = gone.countsAsked;
  await wait(1200);
  check("and off asks the host nothing", (await lensesNow()).countsAsked === askedWhenOff);

  const viaPage = await api.act("settings", { key: "codeLens", value: "true" });
  const back = await settled("the lenses to come back", (now) => now.drawn.length === 3).catch(async () => lensesNow());
  check("the dialog's own switch brings them back", viaPage.did === true && back.drawn.length === 3,
    `${viaPage.detail}; ${back.drawn.length} drawn`);
  check("the setting reads back on", (await api.settings()).codeLens === true);
} finally {
  await api.act("closeDialogs").catch(() => {});
  await api.settings({ codeLens: switchWas }).catch((error) => console.log(`  (the switch was left on: ${error.message})`));
  await userModule.dispose();
  await probeModule.dispose();
}

process.exit(done());
