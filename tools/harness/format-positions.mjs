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

import { open, wait, waitFor, waitUntilStable, reporter } from "./xlide-api.mjs";

const api = await open({});
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

const { check, done } = reporter();

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

/**
 * Where the finding SITS, as a value that can be compared with the one before it.
 *
 * Deliberately not whether it is in the right place: that is what the checks decide, and a wait
 * that asks the same question as the check it precedes cannot fail, only time out.
 */
const findingSpot = async () => {
  const finding = ((await api.problems(target)).findings ?? [])
    .find((one) => (one.message ?? "").includes("Close"));
  return finding ? `${finding.line}:${finding.column}` : "(none)";
};

/** True once no tab anywhere is holding the target. */
const tabGone = async () =>
  !(await api.ui()).workspace.groups.some((group) =>
    group.tabs.some((tab) => tab.module?.toLowerCase() === target.toLowerCase()));

if (runnable) try {
  await api.pane("close", { module: target, project: project.projectId, answer: "discard" });
  await waitFor("the tab to go before the module is reseeded", tabGone);

  await api.writeModule(target, SEED.join("\r\n"), project.projectId);
  await waitFor("the seed to be in the module", async () =>
    ((await api.readModule(target, project.projectId)).text ?? "").includes("Workbooks.Close n"));

  await api.pane("open", { module: target, project: project.projectId });
  await waitFor(`${target} shown`, async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith(`/${target.toLowerCase()}`) ? ui : null;
  });

  await waitFor("the finding to appear", async () =>
    ((await api.problems(target)).findings ?? []).some((one) => (one.message ?? "").includes("Close")));

  // LET THE ANALYSER CATCH UP. A finding appearing is not a pass having finished, and formatting
  // into the middle of one is a different experiment from formatting a module it has settled on
  // (the developer, 2026-08-08).
  //
  // Was `wait(4000)`: a guess at how long settling takes, too long on a quiet machine and too
  // short on a busy one. Settling means the position stops CHANGING, so that is what is waited
  // for. Whether the position is RIGHT is the check below, and it can still fail.
  await waitUntilStable(findingSpot);

  const before = await look("before the format, the statement at column 1:");

  check("the finding is on the word to start with",
    before.finding?.line === before.line && before.finding?.column === before.column,
    `word at ${before.line}:${before.column}, finding at ${before.finding?.line}:${before.finding?.column}`);

  console.log("  running Format Module");
  /*
   * THE FINDING HAS BEEN REPUBLISHED, and only then whether it has stopped moving.
   *
   * Stability alone is not enough here and measuring that cost a run: for the first few hundred
   * milliseconds after the format the finding sits perfectly still at its OLD position, because
   * the analyser has not run again yet. Three quiet polls is satisfied by that, and the suite then
   * reported word at 6:15 against finding at 6:11 - which is precisely the symptom this file
   * exists to catch, arrived at by measuring too early rather than by the product being wrong.
   *
   * So: wait for it to move AT ALL, then for it to settle. Neither asks whether it moved to the
   * right place; that is the check below, and it can still fail. Tolerant, because a finding that
   * never moves is a real result and belongs in the check rather than in a timeout.
   */
  const spotBeforeFormat = await findingSpot();
  await api.act("format");
  await waitFor("the finding to be republished after the format",
    async () => (await findingSpot()) !== spotBeforeFormat).catch(() => null);
  await waitUntilStable(findingSpot);

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
  //
  // THE ONE SLEEP IN THIS FILE THAT STAYS. Every other one was a guess at how long something
  // takes and is now a wait for the thing itself; this one is not waiting for anything, it is
  // letting time pass on purpose, because "the finding is still on the word after four seconds"
  // is the assertion. Replacing it with a condition would delete the check.
  await wait(4500);
  const settled = await look("four seconds later:");

  check("it is still on the word once everything settles",
    settled.finding?.line === settled.line && settled.finding?.column === settled.column,
    `word at ${settled.line}:${settled.column}, finding at ${settled.finding?.line}:${settled.finding?.column}`);

  // PARITY WITH THE EDITOR UNDERNEATH. A format that leaves the native pane holding different
  // text is not a format that finished, whatever the page shows (the developer, 2026-08-07:
  // full parity with content, for every action that touches the editor surface).
  const sync = await waitFor("the write-back to reach the native pane", async () => {
    const answer = await api.inSync();
    return answer.agreed ? answer : null;
  }, { budgetMs: 20000 }).catch(async () => api.inSync());

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
  await waitFor("the caret to reach the end of the indent", async () => {
    const focus = (await api.ui()).focus;
    return focus?.line === indented && focus?.column === indentEnds;
  });

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
  await waitFor("the caret to reach the far side of the word", async () => {
    const focus = (await api.ui()).focus;
    return focus?.line === indented && focus?.column === afterWord;
  });

  const single = await api.act("backspace");
  check("Backspace with code before it still deletes one character",
    single.data?.column === afterWord - 1 && (single.data?.text ?? "").includes("Workbook."),
    `caret at column ${single.data?.column}, line now ${JSON.stringify(single.data?.text)}`);

  /*
   * AND ON A LINE THAT IS NOTHING BUT INDENT, IT TAKES ALL OF IT.
   *
   * Backspace already goes back a level rather than a space, but a blank line two levels in
   * still cost a press per level to clear, and clearing it is the whole of what a developer is
   * doing there - there is nothing else on the line to keep (the owner, 2026-08-21). The line
   * stays, empty; joining it upwards is what the next press does, as it always did.
   *
   * Driven through the KEY. Smart Backspace is a keybinding with no command id, so the action
   * used to reach the editor's stock deleteLeft and none of the product's own rules; the two
   * checks above passed either way because neither of them is one.
   */
  // Built through the PAGE - Enter for a blank line at this indent, Tab for a second level -
  // because by now the surface holds edits the module has not been given, and a write here
  // would be replaced by the surface's own write-back rather than landing (it timed out doing
  // exactly that while this was being written). TWO levels is the point: one level is a single
  // level either way, so it cannot tell this rule from the editor's own tab stops.
  const nowAt = ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
    .split(/\r?\n/)[indented - 1] ?? "";
  const endOfLine = nowAt.length + 1;
  await api.caret(indented, { module: target, project: project.projectId, column: endOfLine });
  await waitFor("the caret to reach the end of the indented line", async () => {
    const focus = (await api.ui()).focus;
    return focus?.line === indented && focus?.column === endOfLine;
  });

  // Tabbed up to two levels rather than once: the checks above have already taken the indent
  // off this line, so how deep Enter starts depends on what they left behind.
  await api.act("press", { key: "Enter" });
  const want = (size * 2) + 1;
  for (let press = 0; press < 3; press += 1) {
    if (((await api.ui()).focus?.column ?? 0) >= want) {
      break;
    }

    await api.act("press", { key: "Tab" });
  }

  const twoLevels = await waitFor("a blank line indented two levels", async () => {
    const focus = (await api.ui()).focus;
    return (focus?.column ?? 0) >= want ? focus : null;
  });
  check(`Enter and Tab leave the caret ${size * 2} spaces in, on a line holding nothing else`,
    twoLevels.column === want, `the caret was at column ${twoLevels.column}`);

  const cleared = await api.act("backspace");
  check(`one Backspace clears a blank line indented ${size * 2} spaces, not one level of it`,
    cleared.data?.column === 1 && (cleared.data?.text ?? "").length === 0,
    `caret at column ${cleared.data?.column}, line now ${JSON.stringify(cleared.data?.text)}`);

  const joined = await api.act("backspace");
  check("...and the next Backspace joins it upwards, the way it always did",
    (joined.data?.line ?? 0) === indented,
    `caret went to line ${joined.data?.line}, wanted ${indented}, on ${JSON.stringify(joined.data?.text)}`);

  /*
   * AND THE EDITOR UNDERNEATH HOLDS WHAT THIS ONE DOES.
   *
   * The two rules above - a blank line keeping its indent, and Backspace taking all of it - are
   * both about whitespace, which is exactly the kind of difference a page can show while the
   * module holds something else. Neither is finished until the native pane agrees, so the
   * write-back is waited for and the two are compared, the way the format check above does it.
   */
  const afterTyping = await waitFor("the write-back to reach the native pane", async () => {
    const answer = await api.inSync();
    return answer.agreed ? answer : null;
  }, { budgetMs: 20000 }).catch(async () => api.inSync());

  check("the native editor holds what these rules left behind, whitespace included",
    afterTyping.agreed && afterTyping.contentAgrees,
    `native ${afterTyping.nativeLines} line(s), surface ${afterTyping.surfaceLines} line(s), `
    + `content agreeing: ${afterTyping.contentAgrees}`);

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
    await waitFor("the For Each opener to be in the module", async () =>
      ((await api.readModule(target, project.projectId)).text ?? "").includes("For Each"));

    // The caret is placed through the SURFACE, and the line written rather than typed: `type`
    // goes through the host's keyboard pipeline while `press` reaches the page's editor, so a
    // measurement that mixes them presses Enter somewhere other than where it typed.
    await api.caret(4, { module: target, project: project.projectId, column: opener.length + 1 });
    await waitFor("the caret to reach the end of the opener", async () => {
      const focus = (await api.ui()).focus;
      return focus?.line === 4 && focus?.column === opener.length + 1;
    });

    await api.act("press", { key: "Enter" });
    // Enter builds the block AND writes the closer, so the module gains lines. Waited for by LINE
    // COUNT rather than by looking for the closer: naming the closer is what the checks below do.
    await waitFor("Enter to build the block", async () =>
      ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
        .split(/\r?\n/).length > 6);

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
   * TWO ENTERS LEAVE TWO INDENTED LINES, not one indented and one at column 1.
   *
   * The editor trims indentation it inserted itself as soon as the caret leaves the line without
   * typing anything - on by default, and sensible for a language whose files are read in diffs.
   * VBA's editor does not: press Enter twice in the VBE and both lines keep the indent, and the
   * module it writes holds the spaces. So the trim made this surface disagree with the one it
   * covers, and arrowing back up landed at column 1 instead of where the code sits (reported
   * 2026-08-21).
   *
   * Pressed rather than typed, for the reason the block above gives: only a real Enter runs the
   * editor's enter rules at all.
   */
  // Driven entirely through the PAGE - no write to the module. By this point the surface holds
  // edits the module has not been given, and a host rewrite deliberately leaves an unwritten
  // document alone, so writing a module here and reading it back tests the write path rather
  // than the Enter (it timed out doing exactly that while this check was being written).
  const beforeTwo = ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
    .split(/\r?\n/);

  // An indented line that is not a block opener: a plain Enter, not smart Enter's block layout.
  const plain = beforeTwo.findIndex((one) =>
    /^\s+\S/.test(one)
    && !/\b(if|for|do|while|with|select|sub|function|property|type|enum)\b/i.test(one));
  check("the module holds an indented statement to press Enter after",
    plain >= 0, `lines were ${JSON.stringify(beforeTwo.slice(0, 8))}`);

  if (plain >= 0) {
    const at = plain + 1;
    const endOfIt = beforeTwo[plain].length + 1;
    await api.caret(at, { module: target, project: project.projectId, column: endOfIt });
    await waitFor("the caret to reach the end of the indented statement", async () => {
      const focus = (await api.ui()).focus;
      return focus?.line === at && focus?.column === endOfIt;
    });

    await api.act("press", { key: "Enter" });
    await api.act("press", { key: "Enter" });
    await waitFor("both new lines to exist", async () =>
      ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
        .split(/\r?\n/).length >= beforeTwo.length + 2);

    const afterTwo = ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
      .split(/\r?\n/);
    const indent = beforeTwo[plain].search(/\S/);
    const first = afterTwo[plain + 1] ?? "";
    check("pressing Enter twice leaves the FIRST line indented, not emptied",
      first.trim() === "" && first.length >= indent,
      `it came back as ${JSON.stringify(first)}, around it ${JSON.stringify(afterTwo.slice(plain, plain + 4))}`);
    // AND ARROWING BACK UP LANDS ON THE INDENT, which is the gesture that was reported: the
    // caret after the second Enter sits on the SECOND line, which keeps its whitespace either
    // way, so asking about that one alone passes with the trim on and proves nothing.
    //
    // Through the editor's own cursorUp - what the ArrowUp key is bound to. `press` carries the
    // keys that edit (Enter, Tab, Backspace, Delete, Escape) and not the ones that only move,
    // and a caret set through the route would not carry the column the arrow remembers.
    const back = await api.ask(
      '(() => { const ed = globalThis.xlideBridge.workspace.activeEditor();'
      + ' ed.trigger("keyboard", "cursorUp", null); const at = ed.getPosition();'
      + ' return { line: at.lineNumber, column: at.column }; })()');
    check("...and arrowing back up onto it lands on that indent, not at column 1",
      back?.line === plain + 2 && (back?.column ?? 0) > indent,
      `the caret came back to line ${back?.line}, column ${back?.column}`);
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
  await waitFor("the pane to close, which is the case this check is about", tabGone);

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
    await waitFor("the fixture's permanent error to be qualified away", async () =>
      ((await api.readModule(AMBIGUOUS, project.projectId)).text ?? "")
        .includes("Helpers.Recalculate"));
  }

  const retired = await waitFor("every finding to retire", async () => {
    const all = (await api.problems()).findings ?? [];
    return all.length === 0 ? true : null;
  }, { budgetMs: 25000 }).catch(() => false);

  check("the findings retire once the whole project is clean",
    retired === true,
    "problems() still reports findings for code that is no longer there. A project that goes "
    + "clean must publish an EMPTY set, or the last non-empty one stands forever: "
    + JSON.stringify(((await api.problems()).findings ?? [])
      .map((one) => `${one.module} ${one.line}:${one.column}`)));
} finally {
  await api.pane("close", { module: target, project: project.projectId, answer: "discard" });
  await api.writeModule(target, original, project.projectId);
  // Read back rather than slept on, and never allowed to throw: this is the `finally`, so a
  // failure here would replace whatever really went wrong with a timeout about the tidying up.
  await waitFor("the fixture to be back as it was", async () =>
    ((await api.readModule(target, project.projectId)).text ?? "").trim() === original.trim())
    .catch(() => null);

  if (ambiguousOriginal.length > 0) {
    await api.writeModule(AMBIGUOUS, ambiguousOriginal, project.projectId);
    await waitFor(`${AMBIGUOUS} to be back as it was`, async () =>
      ((await api.readModule(AMBIGUOUS, project.projectId)).text ?? "").trim()
        === ambiguousOriginal.trim()).catch(() => null);
  }

  const restored = ((await api.readModule(target, project.projectId)).text ?? "").trim() === original.trim();
  const ambiguousBack = ambiguousOriginal.length === 0
    || ((await api.readModule(AMBIGUOUS, project.projectId)).text ?? "").trim() === ambiguousOriginal.trim();
  console.log(`\n${target} restored: ${restored}, ${AMBIGUOUS} restored: ${ambiguousBack}`);

  process.exitCode = done();
}
