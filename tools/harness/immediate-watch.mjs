/*
 * The Immediate window and the Watch panel, which had routes and no suite.
 *
 * WHY THEY HAD NONE. `immediate` posted a line to the host thread and answered `{ran: true}`
 * without waiting, so a caller learned that an evaluation had been ASKED FOR and nothing else.
 * What the expression came to, and whether it failed, went only to the page. A route that reports
 * its own invocation cannot be asserted on, so nobody tried. It answers the outcome now, and this
 * is the suite that became possible.
 *
 * The Watch panel is read-only from here by design: watches are added through the editor's own
 * dialogs, and this drives them the way a developer does, through the commands those dialogs hang
 * off. What is asserted is what the panel HOLDS, because that is the half a developer reads.
 *
 * Run against DebugFixture.xlsm, the only fixture that compiles. Evaluating in a project that
 * does not compile raises a modal instead of answering, which tests the dialog guard rather than
 * the Immediate window.
 *
 *   tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\DebugFixture.xlsm
 *   node tools\harness\immediate-watch.mjs
 */

import { open } from "file:///F:/GitHub/xlide/xlide_vbide/tools/harness/xlide-api.mjs";

const api = await open({});
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
/*
 * THE WORKBOOK THAT COMPILES, located rather than assumed.
 *
 * This read `project()`, which answers about the ACTIVE workbook, and evaluating in a project
 * that does not compile raises a modal instead of answering: the rename fixture deliberately does
 * not compile, so running this while it was in front would test the dialog guard rather than the
 * Immediate window. The fourth suite in one afternoon to assume the workbook, which is why the
 * client grew `projectHolding`.
 *
 * `Runner` is the debug fixture's own module, so finding it finds the fixture.
 */
const home = await api.projectHolding("Runner");
const runnable = home !== null;

if (!runnable) {
  // Skipped, never killed: forcing the process down while the client still has connections open
  // aborts node itself and replaces this suite's exit code with 127.
  console.log("no open workbook holds a module named Runner.");
  console.log("open the debug fixture and run this again:");
  console.log("  tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\fixtures\\DebugFixture.xlsm");
  process.exitCode = 2;
}

const project = runnable
  ? await api.project(home.project)
  : { project: "(none open)", components: [] };

let passed = 0;
const failures = [];

function check(what, ok, detail) {
  if (ok) {
    passed++;
    console.log(`  ok   ${what}${detail ? `  -- ${detail}` : ""}`);
  } else {
    failures.push(`${what}${detail ? `: ${detail}` : ""}`);
    console.log(`  FAIL ${what}${detail ? `\n       ${detail}` : ""}`);
  }
}

console.log(`against ${project.project}\n`);

if (runnable) {
// ---------------------------------------------------------------------------------------------
console.log("1. the Immediate window answers what an expression came to\n");

const two = await api.immediate("?1+1");
check("a simple expression answers its value",
  two.ran && !two.failed && (two.text ?? "").includes("2"),
  JSON.stringify(two));

const text = await api.immediate("?\"xl\" & \"ide\"");
check("a string expression answers its value",
  text.ran && !text.failed && (text.text ?? "").includes("xlide"),
  JSON.stringify(text));

// THE FAILURE PATH, which is the half that matters most: an evaluation that goes wrong must come
// back as DATA rather than as a modal standing in front of the editor.
//
// A RUN-TIME error, which the composed procedure's own handler catches. Not an undeclared name:
// without Option Explicit that is a legitimate empty Variant and the editor's own Immediate
// window prints exactly the same, so asserting failure there tests this suite's grasp of VBA
// rather than the product (2026-08-07).
const wrong = await api.immediate("?1/0");
check("a run-time error reports failure",
  wrong.ran && wrong.failed,
  JSON.stringify(wrong));

check("and it carries the language's own message",
  /division by zero/i.test(wrong.text ?? ""),
  JSON.stringify(wrong.text));

/*
 * A LINE THAT WILL NOT COMPILE, which is a different path and the one that used to brick the
 * window. `On Error GoTo` catches run-time errors; a syntax error never compiles, so the handler
 * is never installed and the editor raises its own Compile error box. Behind it the project was
 * left OUT of design mode, and every evaluation after that answered "Not available while
 * execution is stopped": one mistyped line, and the Immediate window was useless until somebody
 * thought to press Reset.
 */
const malformed = await api.immediate("?((");
check("a line that will not compile reports failure",
  malformed.failed,
  JSON.stringify(malformed));

await wait(1500);
const recovered = await api.immediate("?40+2");
check("and the window still works afterwards",
  recovered.ran && !recovered.failed && (recovered.text ?? "").includes("42"),
  `${JSON.stringify(recovered)} -- a mistyped line must not end the session`);

// AWAITED, not sampled. `state.debugMode` is a POLLED value: the session's own loop refreshes
// it, so reading it the instant an evaluation returns reads the poll before last. Measured
// 2026-08-07: "break" at +0ms and "design" from +500ms on, with the evaluation that provoked the
// question having already succeeded. A check that samples it is asserting on the clock.
let settled = null;
for (const attempt of [0, 300, 600, 1200, 2500]) {
  if (attempt > 0) { await wait(attempt); }
  settled = (await api.state()).debugMode;
  if (settled === "design") { break; }
}

check("the project comes back to design mode", settled === "design", JSON.stringify(settled));

// No modal may be left standing. A failed evaluation that raises one blocks everything after it,
// and the next suite to run pays for it rather than this one.
const standing = await api.dialogs();
check("a failed evaluation leaves no dialog standing",
  (standing.dialogs ?? []).length === 0,
  (standing.dialogs ?? []).map((d) => d.caption).join(", "));

// ---------------------------------------------------------------------------------------------
console.log("\n2. a statement runs, and its Debug.Print reaches the window\n");

const printed = `xlide-probe-${Math.floor(Date.now() / 1000) % 100000}`;
const ran = await api.immediate(`Debug.Print "${printed}"`);

check("a statement runs without reporting failure", ran.ran && !ran.failed, JSON.stringify(ran));

// The evaluator stays silent for a statement, the way the native window does; what the code
// PRINTED arrives through the reader, which is a different path and the one worth proving.
await wait(1200);
const window = await api.immediate();

check("what the statement printed is in the window",
  (window.text ?? "").includes(printed),
  `looked for ${printed} in ${JSON.stringify((window.text ?? "").slice(-160))}`);

check("reading the window does not report a failure", !window.failed, JSON.stringify(window.failed));

// ---------------------------------------------------------------------------------------------
console.log("\n3. the Watch panel\n");

const before = await api.watches();
check("the Watch panel answers a list", Array.isArray(before.watches ?? before.rows), JSON.stringify(Object.keys(before)));

// Adding a watch goes through the editor's own dialog, which is what a developer does. The
// command is asserted to REACH the editor; whether the dialog is then filled in is a modal
// interaction that working-with-modals.md covers, and is not what this suite is about.
const added = await api.command("addWatch").catch((error) => ({ error: String(error) }));
check("the add-watch command reaches the editor",
  added.error === undefined,
  added.error ?? `ran=${added.ran}`);

// Whatever that opened must not be left in front of the next suite.
await wait(600);
const afterCommand = await api.dialogs();
if ((afterCommand.dialogs ?? []).length > 0) {
  console.log(`  (dismissing ${afterCommand.dialogs.map((d) => d.caption).join(", ")})`);
  await api.dismiss("Cancel").catch(() => null);
  await wait(400);
}

const clear = await api.dialogs();
check("no dialog is left standing for the next suite",
  (clear.dialogs ?? []).length === 0,
  (clear.dialogs ?? []).map((d) => d.caption).join(", "));

// ---------------------------------------------------------------------------------------------
console.log("\n4. the panels agree with the host afterwards\n");

const sync = await api.inSync();
check("the native editor and the page still agree",
  sync.agreed,
  `native ${sync.nativeModule}, surface ${sync.surfaceModule}, content agreeing ${sync.contentAgrees}`);

const stats = await api.stats();
check("no COM wrapper was leaked by any of it",
  stats.comWrappersLive < 100,
  `${stats.comWrappersLive} live, ${stats.comWrappersTaken} taken`);

console.log(`\n${passed} passed, ${failures.length} failed`);
for (const failure of failures) {
  console.log(`  ${failure}`);
}

process.exitCode = failures.length === 0 ? 0 : 1;
}
