/*
 * Where a squiggle lands after Format Module, through the whole stack.
 *
 * THE DEFECT. A finding crosses from the engine as a character offset, and an offset means
 * nothing without the text it was counted in. A diagnostics request may leave the source out,
 * and then the engine picks between its live copy of the module and the copy the project was
 * seeded with. That choice was invisible to the add-in, which converted the offsets against the
 * text it had last read out of the editor. While the two agreed nothing looked wrong.
 *
 * Formatting is what made them disagree. The page holds the formatted text the moment the format
 * runs; the editor underneath still holds the original until the write-back. A statement sitting
 * at indent 0 gains a tab, every character on the line moves one along, and a finding that
 * belonged at 6:12 was drawn at 6:6, six columns to the left, underlining the wrong word, and it
 * stayed there, because nothing about waiting makes two different texts the same.
 *
 * Indent 0 is the case that matters and the case a lazier fixture misses: a line already indented
 * four spaces trades them for one tab of the same visual width, and although the columns do move
 * the finding still lands inside the same word, so the bug hides.
 *
 * engine/test/positions.mjs pins the engine's half of this without a host. This is the other
 * half: the page, the add-in, and the native editor underneath, which must all end up describing
 * the same text.
 *
 *   node tools\harness\format-positions.mjs
 */

import { open } from "file:///F:/GitHub/xlide/xlide_vbide/tools/harness/xlide-api.mjs";

const api = await open({});
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
/*
 * THE WORKBOOK HOLDING THE TARGET, not whichever one happens to be active.
 *
 * This named its module outright and read `project()`, which answers about the ACTIVE workbook.
 * Run against a session holding a different fixture it died on "no module named HelpersExtra",
 * which reads as a broken product rather than a suite asking the wrong workbook. The third suite
 * to make this exact mistake in one afternoon (2026-08-07), which is why the client grew
 * `projectHolding`.
 */
const target = "HelpersExtra";
const home = await api.projectHolding(target);

/*
 * NEVER `process.exit` FROM A SUITE, and this is the one place it was tempting.
 *
 * Forcing the process down while the client still has connections open aborts node itself:
 * "Assertion failed: !(handle->flags & UV_HANDLE_CLOSING), src\win\async.c" on Windows, which
 * replaces the suite's own exit code with 127 and prints a line that reads like a product crash.
 * Measured 2026-08-07: one request then exit is fine, several requests then exit aborts, and
 * setting `exitCode` and letting node drain is clean either way.
 *
 * So the run is SKIPPED rather than killed.
 */
const runnable = home !== null;

if (!runnable) {
  console.log(`no open workbook holds a module named ${target}.`);
  console.log("open the rename fixture and run this again:");
  console.log("  tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\fixtures\\RenameFixture.xlsm");
  process.exitCode = 2;
}

const project = runnable ? await api.project(home.project) : null;
const original = runnable ? (await api.readModule(target, project.projectId)).text ?? "" : "";

/*
 * THE FIXTURE'S OWN DELIBERATE ERROR, PUT ASIDE FOR THE ONE CHECK THAT NEEDS A CLEAN PROJECT.
 *
 * `Consumer` calls `Recalculate` unqualified while `Helpers` and `Rival` both export it. That is
 * on purpose: it is what the rename cases are built around, and VBA itself refuses to compile it.
 * Since the analyzer learned to report it (xlide_vscode #12, 2026-08-09) this project can never
 * publish an empty finding set, and the retirement check below has no reachable clean state to
 * observe - it failed on a correct product for a correct reason.
 *
 * So the call is qualified for the duration and put back in the finally. Qualified is the fix the
 * finding itself asks for, so this is the fixture briefly in its corrected state rather than a
 * check working around an inconvenient truth.
 */
const AMBIGUOUS = "Consumer";
const ambiguousOriginal = runnable
  ? (await api.readModule(AMBIGUOUS, project.projectId)).text ?? ""
  : "";

let passed = 0;
const failures = [];

function check(what, ok, detail) {
  if (ok) {
    passed++;
    console.log(`ok   ${what}`);
  } else {
    failures.push(`${what}${detail ? `: ${detail}` : ""}`);
    console.log(`FAIL ${what}${detail ? `\n     ${detail}` : ""}`);
  }
}

async function until(what, predicate, budgetMs = 25000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(250);
  }
  throw new Error(`timed out waiting for ${what}`);
}

// The offending statement at COLUMN 1, which is the developer's own recipe.
const SEED = [
  "Option Explicit",
  "",
  "Public Sub Probe()",
  "    Dim n As Long",
  "    n = 1",
  "Workbooks.Close n",
  "    Debug.Print n",
  "End Sub",
  "",
];

/** Where the statement really is, where the finding says it is, and whether a squiggle is on it. */
async function look(label) {
  const lines = ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
    .split(/\r?\n/);
  const line = lines.findIndex((one) => one.includes("Close")) + 1;
  const column = line > 0 ? lines[line - 1].indexOf("Close") + 1 : 0;

  const finding = ((await api.problems(target)).findings ?? [])
    .find((one) => (one.message ?? "").includes("Close"));

  const spot = line > 0 ? await api.ui({ line, column }) : null;
  const squiggles = spot?.at?.squiggles?.length ?? 0;

  console.log(`\n  ${label}`);
  console.log(`      the word is at    ${line}:${column}   ${JSON.stringify(lines[line - 1] ?? "")}`);
  console.log(`      the finding says  ${finding ? `${finding.line}:${finding.column}` : "(none)"}`);
  console.log(`      squiggles on it   ${squiggles}\n`);

  return { line, column, finding, squiggles };
}

if (runnable) try {
  await api.pane("close", { module: target, project: project.projectId, answer: "discard" });
  await wait(1500);
  await api.writeModule(target, SEED.join("\r\n"), project.projectId);
  await wait(2000);
  await api.pane("open", { module: target, project: project.projectId });
  await until(`${target} shown`, async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith(`/${target.toLowerCase()}`) ? ui : null;
  });

  await until("the finding to appear", async () =>
    ((await api.problems(target)).findings ?? []).some((one) => (one.message ?? "").includes("Close")));

  // LET THE ANALYSER CATCH UP. A finding appearing is not a pass having finished, and formatting
  // into the middle of one is a different experiment from formatting a module it has settled on
  // (the developer, 2026-08-08).
  await wait(4000);

  const before = await look("before the format, the statement at column 1:");

  check("the finding is on the word to start with",
    before.finding?.line === before.line && before.finding?.column === before.column,
    `word at ${before.line}:${before.column}, finding at ${before.finding?.line}:${before.finding?.column}`);

  console.log("  running Format Module");
  await api.act("format");
  await wait(1500);

  const straight = await look("immediately after:");

  check("the format moved the statement along",
    straight.column > before.column,
    `column ${before.column} to ${straight.column}, so the format did nothing to expose`);

  check("the finding followed the word it is about",
    straight.finding?.line === straight.line && straight.finding?.column === straight.column,
    `word at ${straight.line}:${straight.column}, finding at ${straight.finding?.line}:${straight.finding?.column}`);

  check("the squiggle is drawn on the word",
    straight.squiggles > 0,
    "the underline is somewhere else");

  // The engine's copy and the surface's, named outright. When they differ the finding is still
  // allowed to be right, because the position now travels with it, but a difference here is what
  // the old arithmetic turned into a misplaced squiggle.
  const held = await api.engineSource(target, { text: true });
  console.log(`      engine holds ${held.engineLines} line(s), surface ${held.surfaceLines}, ` +
    `agreeing: ${held.engineContent === held.surfaceContent}`);

  // AND IT STAYS RIGHT. The old symptom did not heal, so settling is worth asking about.
  await wait(4500);
  const settled = await look("four seconds later:");

  check("it is still on the word once everything settles",
    settled.finding?.line === settled.line && settled.finding?.column === settled.column,
    `word at ${settled.line}:${settled.column}, finding at ${settled.finding?.line}:${settled.finding?.column}`);

  // PARITY WITH THE EDITOR UNDERNEATH. A format that leaves the native pane holding different
  // text is not a format that finished, whatever the page shows (the developer, 2026-08-07:
  // full parity with content, for every action that touches the editor surface).
  const sync = await until("the write-back to reach the native pane", async () => {
    const answer = await api.inSync();
    return answer.agreed ? answer : null;
  }, 20000).catch(async () => api.inSync());

  const below = await api.native({ text: true });

  // NOT A TAB ANYWHERE. VBA's code store will not hold one: the editor expands every tab it is
  // handed to the next four-column stop, on both write paths this product uses and mid-line as
  // well as leading (measured 2026-08-07). While the page could be told to indent with tabs, it
  // and the workbook disagreed for as long as a module stayed open, so the option was removed
  // and indentation is spaces. A tab reaching the page is that decision coming undone.
  check("the page indents with spaces, which is the only thing this host will store",
    !(below.surfaceText ?? "").includes("\t"),
    JSON.stringify((below.surfaceText ?? "").slice(0, 120)));

  check("the native editor holds what the page holds",
    sync.agreed,
    `native ${sync.nativeModule} ${sync.nativeLines} line(s), surface ${sync.surfaceModule} ` +
    `${sync.surfaceLines} line(s), content agreeing: ${sync.contentAgrees}`);

  const everyPane = await api.parityAll();
  check("every other open pane still agrees with the editor",
    everyPane.agreed,
    everyPane.stale.map((one) => one.module).join(", "));


  // BACKSPACE TAKES BACK A LEVEL, which is what makes indenting with spaces bearable now that
  // the tabs option is gone. In LEADING WHITESPACE only: with anything else on the line before
  // the caret it deletes one character, the way it always has.
  const size = (await api.settings()).formatIndentSize;
  const indented = straight.line;

  // The end of the indent, which is where a whole level is at stake. `straight.column` is the
  // word, and putting the caret there tests ordinary deletion instead.
  const indentEnds = size + 1;

  await api.caret(indented, { module: target, project: project.projectId, column: indentEnds });
  await wait(600);

  const level = await api.act("backspace");
  check(`Backspace at the end of the indent takes back all ${size} spaces at once`,
    level.data?.column === 1,
    `caret went to column ${level.data?.column}, wanted 1, line now ${JSON.stringify(level.data?.text)}`);

  check("and takes the whole level, leaving no stray spaces",
    !(level.data?.text ?? "").startsWith(" "),
    JSON.stringify(level.data?.text));

  // With code before the caret it is an ordinary Backspace again, which is the developer's own
  // condition: "assuming no other chars precede it".
  const line = ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
    .split(/\r?\n/)[indented - 1] ?? "";
  const afterWord = line.indexOf("Workbooks") + "Workbooks".length + 1;

  await api.caret(indented, { module: target, project: project.projectId, column: afterWord });
  await wait(600);

  const single = await api.act("backspace");
  check("Backspace with code before it still deletes one character",
    single.data?.column === afterWord - 1 && (single.data?.text ?? "").includes("Workbook."),
    `caret at column ${single.data?.column}, line now ${JSON.stringify(single.data?.text)}`);

  /*
   * ENTER, AND WHAT HANGS OFF IT.
   *
   * Untestable until 2026-08-09. `type` inserts a string, and Monaco applies its enter rules only
   * to a newline typed as ONE character, so nothing here could be driven: auto-indent, smart
   * Enter's block layout, and the closer it writes were all live-untested. `act("press", {key})`
   * exists for this, and these are the first things it is pointed at.
   *
   * The accented loop variable is deliberate. The indentation rules were made Unicode-aware the
   * same day, and until this could be pressed the change shipped reasoned rather than measured.
   * What it proves is the OUTCOME - the block is built and its closer names the loop variable -
   * rather than which of Monaco's rules and this product's smart Enter did which part.
   */
  for (const [what, loopVariable] of [["an ASCII", "item"], ["an accented", "\u00C9l\u00E9ment"]]) {
    const opener = `    For Each ${loopVariable} In coll`;
    await api.writeModule(
      target,
      ["Option Explicit", "", "Public Sub Go()", opener, "End Sub", ""].join("\r\n"),
      project.projectId);
    await wait(1800);

    // The caret is placed through the SURFACE, and the line written rather than typed: `type`
    // goes through the host's keyboard pipeline while `press` reaches the page's editor, so a
    // measurement that mixes them presses Enter somewhere other than where it typed.
    await api.caret(4, { module: target, project: project.projectId, column: opener.length + 1 });
    await wait(600);
    await api.act("press", { key: "Enter" });
    await wait(900);

    const built = ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
      .split("\n").map((one) => one.replace("\r", ""));

    const body = built.findIndex((one) => one.trim().length === 0 && one.length > 0);
    check(`Enter after ${what} For Each indents the body past the opener`,
      body > 0 && body < built.length && (built[body].length > opener.search(/\S/)),
      `the lines after it were ${JSON.stringify(built.slice(3, 8))}`);

    // CASE-INSENSITIVELY, because the editor unifies identifier case across a project and will
    // respell the loop variable if it knows the name: `item` came back as `Item`, in the opener
    // as well as the closer. That is VBA's doing and it is correct; asserting the spelling the
    // probe typed would fail on it.
    check(`and the closer it writes names ${loopVariable}`,
      built.some((one) => one.trim().toLowerCase() === `next ${loopVariable}`.toLowerCase()),
      `no "Next ${loopVariable}" in ${JSON.stringify(built.slice(3, 9))}`);
  }

  /*
   * AND THE FINDING GOES AWAY WHEN THE CODE DOES.
   *
   * The pass used to publish per project and only when that project had something to say, so a
   * workbook going clean said nothing and the last non-empty set stood forever: the error stayed
   * on screen, pointing at a line that no longer held it. Asking `problems()` about a module
   * whose text had been restored answered that it still called Close with an argument it does not
   * contain (2026-08-07).
   *
   * The seed is put back to something with nothing wrong in it, and the finding has to retire.
   */
  /*
   * WITH THE PANE CLOSED, which is the case that was broken and the only one that tests it.
   *
   * A module OPEN on the surface is re-analysed live on every pause, and that path publishes per
   * module and clears its own findings, so an open module's error retires whatever the project
   * pass does. Written while CLOSED there is no live path, and the project pass is the only thing
   * that can retire it. A first version of this check left the pane open, passed against a build
   * with the defect deliberately restored, and proved nothing (2026-08-07).
   */
  await api.pane("close", { module: target, project: project.projectId, answer: "discard" });
  await wait(1500);

  /*
   * RESTORED TO THE ORIGINAL, so the WHOLE PROJECT goes clean, which is the only state the defect
   * shows in. Removing just the offending line leaves the module without the procedure its caller
   * expects, so another module complains, the pass is non-empty and publishes anyway, and the
   * stale finding is cleared by accident. That is exactly why a first version of this check passed
   * against a build with the guard deliberately restored, and proved nothing.
   */
  await api.writeModule(target, original, project.projectId);

  // And the fixture's permanent error, so "clean" is a state this project can actually be in.
  if (ambiguousOriginal.includes('Recalculate "ambiguous"')) {
    await api.writeModule(
      AMBIGUOUS,
      ambiguousOriginal.replace('Recalculate "ambiguous"', 'Helpers.Recalculate "ambiguous"'),
      project.projectId);
    await wait(1800);
  }

  const retired = await until("every finding to retire", async () => {
    const all = (await api.problems()).findings ?? [];
    return all.length === 0 ? true : null;
  }, 25000).catch(() => false);

  check("the findings retire once the whole project is clean",
    retired === true,
    "problems() still reports findings for code that is no longer there. A project that goes "
    + "clean must publish an EMPTY set, or the last non-empty one stands forever: "
    + JSON.stringify(((await api.problems()).findings ?? [])
      .map((one) => `${one.module} ${one.line}:${one.column}`)));
} finally {
  await api.pane("close", { module: target, project: project.projectId, answer: "discard" });
  await wait(1800);
  await api.writeModule(target, original, project.projectId);
  await wait(1800);

  if (ambiguousOriginal.length > 0) {
    await api.writeModule(AMBIGUOUS, ambiguousOriginal, project.projectId);
    await wait(1200);
  }

  const restored = ((await api.readModule(target, project.projectId)).text ?? "").trim() === original.trim();
  const ambiguousBack = ambiguousOriginal.length === 0
    || ((await api.readModule(AMBIGUOUS, project.projectId)).text ?? "").trim() === ambiguousOriginal.trim();
  console.log(`\n${target} restored: ${restored}, ${AMBIGUOUS} restored: ${ambiguousBack}`);

  console.log(`\n${passed} passed, ${failures.length} failed`);
  for (const failure of failures) {
    console.log(`  ${failure}`);
  }

  process.exitCode = failures.length === 0 ? 0 : 1;
}
