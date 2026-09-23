/*
 * The Immediate panel, driven where the developer drives it: a line typed into the page, Enter
 * pressed, and the editor's own box answered the way a person answers it.
 *
 * WHY THE DOOR'S SUITE COULD NOT SEE #30. immediate-watch.mjs evaluates through the `immediate`
 * route, which runs the line on the host thread and watches it from the door's own thread. The
 * panel ran its line inside the browser's WebMessageReceived handler instead, and a line that
 * would not compile left VBA stopped in the scratch procedure with `Application.Run` suspended
 * under that handler. WebView2 never re-enters a callback, so from then on nothing the page sent
 * reached the host: Debug.Print printed nothing, Break and Reset did nothing, and a script asked
 * of the page never answered. The route kept working the whole time, and the gate was green.
 *
 * THE BOX IS RAISED WITH NO REQUEST IN FLIGHT. The door takes a dialog that appears within 2.75s
 * of a request (its watch delays, 250 + 750 + 1750ms) for that request's own, and clears it at the
 * next one. A developer's box never meets that, so the page presses Enter itself four seconds after
 * the request that typed the line, and the box is watched through `dialogs`, which reads windows
 * and needs nothing from the host thread.
 *
 * Run against DebugFixture.xlsm, the only fixture that compiles: in one that does not, every line
 * raises a box, and the line under test is not the one that did.
 *
 *   tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\DebugFixture.xlsm
 *   node tools\harness\immediate-panel.mjs
 */

import { open, wait, reporter } from "./xlide-api.mjs";

const api = await open({});

// `Runner` is the debug fixture's own module, so finding it finds the fixture.
const home = await api.projectHolding("Runner");
if (home === null) {
  // Skipped, never killed: forcing the process down while the client still has connections open
  // aborts node itself and replaces this suite's exit code with 127.
  console.log("no open workbook holds a module named Runner.");
  console.log("open the debug fixture and run this again:");
  console.log("  tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\fixtures\\DebugFixture.xlsm");
  process.exit(2);
}

const { check, done } = reporter();
console.log(`against ${home.project}\n`);

/** Longer than the door's dialog attribution window, 250 + 750 + 1750ms. */
const QUIET_MS = 4000;

/** Types each line into the panel and presses Enter, in the page, after `delayMs`. */
const enter = (lines, delayMs = 0) => api.ask(`(() => {
  const input = document.querySelector("#immediate-input");
  if (!input) { return "no input"; }
  const press = () => {
    for (const line of ${JSON.stringify(lines)}) {
      input.value = line;
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    }
  };
  if (${delayMs} > 0) { setTimeout(press, ${delayMs}); } else { press(); }
  return "typed";
})()`);

/** The panel's lines as { kind, text }, oldest first. */
const panelLines = async () => {
  const lines = await api.ask(`[...document.querySelectorAll("#immediate-log > *")]
    .map((one) => ({ kind: one.className || "", text: one.textContent.trim() }))`);
  return Array.isArray(lines) ? lines : [];
};

/** Waits in the page for a line after the first `from` whose kind and text match. */
const panelShows = (from, kind, pattern, waitMs = 8000) => api.until(`[...document.querySelectorAll("#immediate-log > *")]
  .slice(${from})
  .some((one) => (one.className || "").includes(${JSON.stringify(kind)})
    && new RegExp(${JSON.stringify(pattern)}, "i").test(one.textContent))`, { waitMs });

const standingBoxes = async () => (await api.dialogs()).dialogs ?? [];

// A line prints, before anything else is asked of the panel.
let from = (await panelLines()).length;
await enter(['Debug.Print "panel-before"']);
check("a line entered in the panel prints",
  (await panelShows(from, "result", "^panel-before$"))?.met === true,
  JSON.stringify((await panelLines()).slice(from)));

/*
 * The line that will not compile, pressed with no request in flight.
 */
from = (await panelLines()).length;
await enter(['debug.prnt "A"'], QUIET_MS);

let box = null;
for (const until = Date.now() + QUIET_MS + 8000; Date.now() < until && box === null;) {
  box = (await standingBoxes()).find((one) => /compile error/i.test(one.text ?? "")) ?? null;
  if (box === null) { await wait(200); }
}

check("the line raises the editor's own compile error box",
  box !== null,
  box !== null
    ? JSON.stringify(box.text)
    : "no box: the line compiled, or something answered it before it could be seen");

if (box !== null) {
  // Left for the developer, as the editor's own Immediate window leaves it: the box may as easily
  // be their own MsgBox asking them something, and a panel that pressed OK on it would be answering
  // for them.
  await wait(1500);
  check("and it is left standing for the developer to answer",
    (await standingBoxes()).some((one) => one.window === box.window),
    JSON.stringify(await standingBoxes()));

  // The page answers while the box stands. Under the handler it could not: a modal inside a
  // WebView2 callback holds every later callback, the answer to this script among them.
  const asked = Date.now();
  const alive = await api.ask("6 * 7");
  check("the page answers while the box stands",
    alive === 42,
    `answered ${JSON.stringify(alive)} after ${Date.now() - asked}ms`);

  const pressed = await api.dismiss("OK", box.caption);
  check("the developer's OK lands on it", pressed?.ran === true, JSON.stringify(pressed));

  check("the panel says what the box said",
    (await panelShows(from, "failed", "compile error"))?.met === true,
    JSON.stringify((await panelLines()).slice(from)));

  const said = (await panelLines()).slice(from);
  check("and not the error of the stopped call being reset",
    !said.some((line) => /HRESULT|COM component/i.test(line.text)),
    JSON.stringify(said));

  check("in one line, as the box's words and nothing else",
    said.filter((line) => line.kind.includes("failed")).every((line) => !/\s{2}/.test(line.text)),
    JSON.stringify(said));
}

// AWAITED, not sampled: `debugMode` is polled, and the stop takes a reset to clear.
let mode = null;
for (const pause of [0, 300, 600, 1200, 2500]) {
  if (pause > 0) { await wait(pause); }
  mode = (await api.state()).debugMode;
  if (mode === "design") { break; }
}

check("the editor is back in design mode", mode === "design", JSON.stringify(mode));
check("and no box is left standing", (await standingBoxes()).length === 0, JSON.stringify(await standingBoxes()));

// The next line prints - the issue's own second step.
from = (await panelLines()).length;
await enter(['Debug.Print "panel-after"']);
check("the next line prints",
  (await panelShows(from, "result", "^panel-after$"))?.met === true,
  JSON.stringify((await panelLines()).slice(from)));

// IN THE ORDER TYPED. The handler ran each line to its end before the browser delivered the next;
// the lines now run off the browser's callback, and nothing about a pool thread keeps them in order
// unless something is made to.
from = (await panelLines()).length;
await enter(['Debug.Print "order-1"', 'Debug.Print "order-2"', 'Debug.Print "order-3"']);
await panelShows(from, "result", "^order-3$");
const results = (await panelLines()).slice(from)
  .filter((line) => line.kind.includes("result"))
  .map((line) => line.text);
check("lines typed back to back answer in the order they were typed",
  JSON.stringify(results) === JSON.stringify(["order-1", "order-2", "order-3"]),
  JSON.stringify(results));

// The recovery takes the scratch module away as well; a reset alone leaves it in the project.
// Asked by name, because every component list this door answers leaves the scratch module out on
// purpose, and one said it was gone while VBA sat stopped inside it (scratchModuleStands).
check("the scratch module is gone", !(await api.scratchModuleStands()));

process.exitCode = done();
