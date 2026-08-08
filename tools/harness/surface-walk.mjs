/*
 * A randomised walk over the surface, checking every invariant after every step.
 *
 * Scripted sequences only reach the states somebody thought of. This picks each action from a
 * deterministic pseudo-random stream, so a failure is replayable from the seed it prints, and it
 * checks the whole surface after every step rather than at the end. State-machine defects live in
 * orders nobody writes down.
 *
 * TWO THINGS HERE ARE LOAD-BEARING, and both were added after the walk lied.
 *
 * The actions are WEIGHTED. Unweighted, there are two ways to close a tab and one to open one, so
 * the workspace drains and stays empty: the first run reported 1,223 checks and nothing broken
 * while spending 54 of 70 steps with nothing open at all.
 *
 * And it reports the STATES IT REACHED. A run that never holds two modules of the same name
 * passes every label check vacuously and looks identical to a run that proves something. That
 * line is how the `pane` route's dropped project argument was found: the walk opened both
 * workbooks' Helpers, said `collision=0`, and was right.
 *
 * Run it against a session with TWO workbooks open, or the interesting half does not exist:
 *   tools\harness\Start-Excel.ps1 -Workbook artifactsixtures\RenameFixture.xlsm,artifactsixtures\TwinFixture.xlsm
 *   node tools\harness\surface-walk.mjs --steps 80 --seed 424242
 */

import { open } from "./xlide-api.mjs";

const seedArg = process.argv.indexOf("--seed");
const SEED = seedArg >= 0 ? Number(process.argv[seedArg + 1]) : 20260807;
const stepsArg = process.argv.indexOf("--steps");
const STEPS = stepsArg >= 0 ? Number(process.argv[stepsArg + 1]) : 60;

// mulberry32: small, deterministic, and reproducible from the seed printed on failure.
let state = SEED >>> 0;
function random() {
  state = (state + 0x6d2b79f5) >>> 0;
  let t = state;
  t = Math.imul(t ^ (t >>> 15), t | 1);
  t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
}
const pick = (list) => list[Math.floor(random() * list.length)];

const api = await open({});
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const project = await api.project();
const first = await api.ui();
const workbooks = first.explorer.workbooks.map((b) => b.name);

// Every workbook's modules, carrying which workbook each belongs to: a bare name is ambiguous
// across workbooks, which is the whole point of running this with two of them open.
const modules = first.explorer.workbooks.flatMap((b) =>
  b.modules.filter((m) => m.kind !== "document").map((m) => ({ name: m.name, workbook: b.name })));

const broken = [];
let checks = 0;

// Which interesting STATES the walk actually reached. A sweep whose collision checks never saw
// a collision passed them vacuously, and would report the same clean run with the check deleted.
const reached = { collision: 0, split: 0, empty: 0, dialog: 0, pending: 0, twoWorkbooksOpen: 0 };

function check(where, claim, ok, detail) {
  checks++;
  if (!ok) { broken.push(`${where}: ${claim}${detail ? " -- " + detail : ""}`); }
}

async function closeOne() {
  const said = await api.act("closeActive");
  if (!said.did && /Unsaved changes/.test(said.detail)) {
    await api.act("answerCloseConfirm", { answer: "discard" });
  }
}

/*
 * WEIGHTED, and the weights are the point.
 *
 * An unweighted walk over these actions drains the workspace: there are two ways to close a tab
 * and one to open one, so it reaches empty and stays there. The first run of this reported 1,223
 * checks and nothing broken while spending 54 of 70 steps with nothing open, never splitting,
 * and never once holding two modules of the same name — every check about those passed because
 * the state never happened (2026-08-07). Opening is weighted to outpace closing, and the
 * coverage line at the end says which states were actually reached.
 */
const ACTIONS = [
  [6, "open a module", async () => {
    const one = pick(modules);
    await api.pane("open", { module: one.name, project: one.workbook.toLowerCase() });
  }],
  [2, "close the active tab", closeOne],
  [2, "cycle forward", async () => { await api.act("cycleTab", { delta: 1 }); }],
  [2, "cycle back", async () => { await api.act("cycleTab", { delta: -1 }); }],
  [2, "split right", async () => { await api.act("split", { direction: "right" }); }],
  [1, "split down", async () => { await api.act("split", { direction: "down" }); }],
  [2, "unfold a module", async () => {
    const one = pick(modules);
    await api.act("unfoldModule", { module: one.name, workbook: one.workbook });
  }],
  [2, "toggle a workbook", async () => {
    if (workbooks.length > 0) {
      await api.act("expandWorkbook", { workbook: pick(workbooks), open: random() < 0.5 });
    }
  }],
  [1, "open the settings dialog", async () => { await api.act("settings"); }],
  [1, "open the sponsor dialog", async () => { await api.act("sponsors"); }],
  [3, "escape any dialog", async () => { await api.act("closeDialogs"); }],
  [2, "focus the editor", async () => { await api.act("focusEditor"); }],
  [1, "Ctrl+W", async () => { await api.act("key", { code: "KeyW", ctrl: true, target: "document" }); }],
  [2, "put the caret somewhere", async () => {
    await api.caret(1 + Math.floor(random() * 5));
  }],
];

async function sweep(where) {
  const ui = await api.ui();
  const w = ui.workspace;

  /*
   * PARITY WITH THE NATIVE EDITOR, which is the definition of tested for anything that touches
   * the editor (the developer, 2026-08-08).
   *
   * The surface covers the host's own code panes; it does not replace them. Run, Step, Compile
   * and ToggleBreakpoint act on the native ACTIVE CODE PANE and the caret inside it, not on the
   * page — so a page showing one module while the native pane holds another is a Run that
   * executes where nobody is looking and a breakpoint on the wrong line, with nothing on screen
   * to say so. Every check in this repo read the page and the workbook and never the panes
   * below, until this one.
   */
  if (!w.empty) {
    const below = await api.native();
    const page = ui.focus.model ? ui.focus.model.split("/").pop() : null;
    const same = (a, b) => (a ?? "").toLowerCase() === (b ?? "").toLowerCase();

    check(where, "the native pane, the surface and the page name the same module",
      same(below.activeModule, below.surfaceModule) && same(below.surfaceModule, page),
      `native=${below.activeModule} surface=${below.surfaceModule} page=${page}`);

    // The strip and the host's open panes are two lists of the same thing.
    const tabs = w.groups.flatMap((g) => g.tabs).map((t) => t.module.toLowerCase()).sort();
    const panes = below.panes.map((one) => one.module.toLowerCase()).sort();
    check(where, "the tab strip and the host's open panes hold the same modules",
      tabs.length === panes.length && tabs.every((m, at) => m === panes[at]),
      `page=[${tabs}] native=[${panes}]`);

    /*
     * THE CONTENT OF EVERY OPEN MODULE, not only the one on screen.
     *
     * Names agreeing is not parity: a surface holding an empty document for a module the host
     * has 42 lines of passes every name check there is and shows a blank editor. And the active
     * module is not the only one that can drift — a background tab holds a copy nobody is
     * looking at, so it can be wrong until it is clicked, and then the developer finds out.
     *
     * A pane the surface holds no text for is not a disagreement: it has simply never been
     * shown. Holding the WRONG text is the defect.
     */
    const stale = below.panes.filter((pane) =>
      pane.surfaceContent !== null && pane.hostContent !== pane.surfaceContent);

    check(where, "every open module's text matches the workbook's",
      stale.length === 0,
      stale.map((one) => `${one.module}: host=${one.hostContent} surface=${one.surfaceContent}`).join("; "));
  }

  check(where, "the empty view agrees with having no tabs",
    ui.emptyViewShown === w.empty, `emptyView=${ui.emptyViewShown} empty=${w.empty}`);

  for (const g of w.groups) {
    const actives = g.tabs.filter((t) => t.active).length;
    check(where, `group #${g.number} shows exactly one of its tabs`,
      g.tabs.length === 0 ? actives === 0 : (actives === 1 || g.pending !== null),
      `${g.tabs.length} tabs, ${actives} active, pending=${g.pending?.module ?? "none"}`);

    check(where, `group #${g.number} awaits only what it holds`,
      !g.pending || g.tabs.some((t) => t.module === g.pending.module),
      g.pending ? `awaits ${g.pending.module}` : "");

    const keys = new Set(g.tabs.map((t) => `${(t.project ?? "").toLowerCase()}\u0000${t.module.toLowerCase()}`));
    const ghosts = g.recent.filter((k) => !keys.has(k));
    check(where, `group #${g.number} remembers only open documents`, ghosts.length === 0, ghosts.join(","));
  }

  // Exactly one group is the active one, always.
  check(where, "exactly one group is active",
    w.groups.filter((g) => g.active).length === 1,
    `${w.groups.filter((g) => g.active).length} active of ${w.groups.length}`);

  // A group with no tabs should not survive beside a group that has some.
  const empties = w.groups.filter((g) => g.tabs.length === 0).length;
  check(where, "an emptied group dissolves unless it is the only one",
    empties === 0 || w.groups.length === 1,
    `${empties} empty of ${w.groups.length}`);

  const counts = new Map();
  for (const t of w.groups.flatMap((g) => g.tabs)) {
    counts.set(t.module.toLowerCase(), (counts.get(t.module.toLowerCase()) ?? 0) + 1);
  }
  for (const t of w.groups.flatMap((g) => g.tabs)) {
    const collides = (counts.get(t.module.toLowerCase()) ?? 0) > 1;
    check(where, `the label of ${t.module} is qualified exactly when it is ambiguous`,
      collides === (t.label !== t.module) || (collides && !t.project));
  }

  if (w.groups.length > 1) { reached.split++; }
  if (w.empty) { reached.empty++; }
  if (ui.dialogs.length > 0) { reached.dialog++; }
  if (w.groups.some((g) => g.pending)) { reached.pending++; }
  if (new Set(w.groups.flatMap((g) => g.tabs).map((t) => t.project)).size > 1) { reached.twoWorkbooksOpen++; }
  if ([...counts.values()].some((n) => n > 1)) { reached.collision++; }

  const unfolded = ui.explorer.workbooks.flatMap((b) => b.modules.filter((m) => m.unfolded));
  check(where, "at most one module is unfolded", unfolded.length <= 1,
    unfolded.map((m) => m.name).join(","));
  check(where, "explorer.unfolded agrees with the rows",
    (ui.explorer.unfolded === null) === (unfolded.length === 0));

  check(where, "no model outlives the documents", ui.census.models <= ui.census.documents + 1,
    `models=${ui.census.models} documents=${ui.census.documents}`);

  /*
   * THE PROVIDER GATE, and it is the most valuable check here.
   *
   * Every language provider answers only for the host-active module and returns nothing
   * otherwise, so when the editor shows one module and the host believes another — or believes
   * none — hover, completions, signature help and quick fixes all go silent on a tab that looks
   * and behaves normally. Nothing else on screen is wrong, which is why it reached a developer
   * rather than a test: closing the active module's pane left the host with `active: null`, the
   * page promoted a tab to show, and the host was never told (2026-08-08).
   */
  if (!w.empty) {
    check(where, "the editor and the host agree which module is active",
      ui.focus.host !== null && ui.focus.model === ui.focus.host.model,
      `editor=${ui.focus.model} host=${ui.focus.host?.model ?? "null"} — every language provider is silent`);
  }

  // Panes are permanent or not; none should vanish from the list.
  check(where, "every pane is still listed", ui.panes.length >= 6, `${ui.panes.length} panes`);

  return ui;
}

/*
 * The walk STARTS from the interesting state rather than hoping to wander into it.
 *
 * Left to chance the collision never happened: two workbooks were open for 30 of 80 steps and
 * not once did both hold a module of the same name, so every label check passed vacuously.
 * Opening the twins up front means every subsequent step is checked against the state the
 * defects in this class actually live in.
 */
const twins = modules
  .filter((m) => modules.filter((other) => other.name.toLowerCase() === m.name.toLowerCase()).length > 1);

for (const twin of twins) {
  await api.pane("open", { module: twin.name, project: twin.workbook.toLowerCase() });
  await wait(900);
}

console.log(`seed ${SEED}, ${STEPS} steps, ${modules.length} modules, ${workbooks.length} workbooks`);
console.log(`opened ${twins.length} twin(s) up front: ${twins.map((t) => `${t.name}@${t.workbook}`).join(", ") || "none"}\n`);

// One entry per unit of weight, so `pick` stays a uniform choice over a loaded list.
const WEIGHTED = ACTIONS.flatMap(([weight, label, run]) =>
  Array.from({ length: weight }, () => [label, run]));

const history = [];
for (let step = 1; step <= STEPS; step++) {
  const [label, run] = pick(WEIGHTED);
  history.push(label);
  try {
    await run();
  } catch (error) {
    console.log(`  step ${step}: ${label} THREW: ${error.message}`);
  }
  await wait(320);

  const before = broken.length;
  const ui = await sweep(`step ${step} (${label})`);
  if (broken.length > before) {
    console.log(`  step ${step}: ${label.padEnd(26)} groups=${ui.workspace.groups.length} ` +
      `tabs=${ui.workspace.groups.reduce((n, g) => n + g.tabs.length, 0)}   ${broken.length - before} BROKEN`);
  }
}

// Leave nothing standing.
await api.act("closeDialogs");
await api.resetLayout();

console.log(`\n${checks} checks over ${STEPS} steps, ${broken.length} broken`);
console.log("states reached: " + Object.entries(reached).map(([k, n]) => `${k}=${n}`).join("  "));
const never = Object.entries(reached).filter(([, n]) => n === 0).map(([k]) => k);
if (never.length > 0) {
  console.log(`  NEVER REACHED (so any check about them passed vacuously): ${never.join(", ")}`);
}
const seen = new Set();
for (const one of broken) {
  const shape = one.replace(/^step \d+ \([^)]*\): /, "").replace(/#\d+/g, "#N").replace(/ -- .*/, "");
  if (seen.has(shape)) { continue; }
  seen.add(shape);
  console.log(`  ! ${one}`);
}
if (broken.length > 0) {
  console.log(`\n  replay with: node stress.mjs --seed ${SEED} --steps ${STEPS}`);
  console.log(`  sequence: ${history.join(" -> ")}`);
}
