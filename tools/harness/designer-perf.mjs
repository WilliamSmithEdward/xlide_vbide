/*
 * A perf walk of the designer surface: every interaction a person actually makes, timed from
 * the moment they make it to the moment the surface has ANSWERED - the canvas redrawn, the panel
 * following, the form itself carrying it - rather than to the moment a call returned.
 *
 * Each is measured five times and reported as median and worst, because one sample of a live COM
 * round trip is a story about that moment rather than about the feature.
 *
 * Run against a session holding FormFixture.xlsm:
 *
 *   tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\FormFixture.xlsm -Fresh
 *   node tools\harness\designer-perf.mjs
 *
 * It PRINTS rather than passes or fails, so it is not gate material: what counts as slow here is
 * a judgement about a person's patience, and a threshold nobody can defend becomes a number the
 * next person raises. It reports the COM ledger at the end for the one thing that IS pass/fail -
 * givenBack must equal disposed, and com-leak.mjs is what holds that line.
 *
 * The first walk (2026-08-15) found a defect rather than a cost: a TextBox whose right edge met
 * a ScrollBar could not be resized, because the selection handles were painted under the
 * neighbour. Timing a surface makes you touch every part of it, which is its own kind of test.
 */
import { open } from "./xlide-api.mjs";

const api = await open();
const { projectId: project } = await api.project();
const form = "EntryForm";
const ROUNDS = 5;

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const canvas = async () => (await api.act("designerCanvas", { module: form })).data;
const text = async () => String((await api.act("designerMarkup", { module: form })).data);

/** Runs one interaction N times, timing act -> the surface agreeing, and prints the spread. */
async function walk(label, once) {
  const runs = [];
  for (let round = 0; round < ROUNDS; round++) {
    const started = Date.now();
    await once(round);
    runs.push(Date.now() - started);
    await wait(120);
  }

  const sorted = [...runs].sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)];
  console.log(`${label.padEnd(34)} median ${String(median).padStart(5)}ms   worst `
    + `${String(sorted[sorted.length - 1]).padStart(5)}ms   [${runs.join(", ")}]`);
  return median;
}

/** Polls until the predicate holds, so a measurement ends when the SURFACE agrees. */
async function until(predicate, budgetMs = 15000) {
  const started = Date.now();
  for (;;) {
    if (await predicate()) return;
    if (Date.now() - started > budgetMs) throw new Error("timed out");
  }
}

console.log(`\nthe designer surface, against ${project}\n`);

// ---- the host's own read: the walk every projection is printed from ----
await walk("designer route (18 controls)", async () => { await api.designer(form, project); });
await walk("designer markup route", async () => { await api.designerMarkup(form, project); });

// ---- opening the tab: nothing on screen -> a drawn form ----
await api.pane("close", { module: form, face: "design" }).catch(() => {});
await walk("open the designer tab", async () => {
  await api.pane("open", { module: form, face: "design" });
  await until(async () => (await canvas()).controls.length > 10);
  await api.pane("close", { module: form, face: "design" });
});

await api.pane("open", { module: form, face: "design" });
await until(async () => (await canvas()).controls.length > 10);

/*
 * The form as it stands before any of this, and the way back to it.
 *
 * A walk that measures gestures also PERFORMS them, so by the third section the controls have
 * drifted - and a drag that has pushed a control against its parent's edge stops moving, which
 * reads as a hang rather than as a clamp. Each section starts from the same document, and the
 * walk puts the form back at the end, saved, the way the suite leaves the fixture as found.
 */
const canonical = String(await api.designerMarkup(form, project));
const reset = async () => {
  await api.act("designerSetMarkup", { module: form, markup: canonical });
  // Waits for the DOCUMENT to be that text, not for the dirty dot to clear: once the walk has
  // saved, the form has moved on, and a document put back to where this run started is honestly
  // dirty against it. What the sections need is the same starting text every time.
  await until(async () => (await text()) === canonical);
};

// ---- selection: the click, and the panel following it host-side ----
await walk("select a control -> panel", async (round) => {
  const name = round % 2 === 0 ? "RegionPick" : "NameBox";
  await api.act("designerSelect", { module: form, control: name });
  await until(async () => (await api.ui()).properties?.component === name);
});

// ---- the gestures, act -> the canvas drawing it ----
await reset();
await api.act("designerSelect", { module: form, control: "NameBox" });
await walk("nudge (arrow key)", async () => {
  const before = (await canvas()).controls.find((c) => c.name === "NameBox")?.left ?? 0;
  await api.ask('(() => { const el = document.querySelector(\'.designer-view[data-module="EntryForm"] .designer-canvas-scroll\'); el.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true })); return "sent"; })()');
  await until(async () => ((await canvas()).controls.find((c) => c.name === "NameBox")?.left ?? 0) !== before);
});

await reset();
await walk("drag a control", async (round) => {
  const step = round % 2 === 0 ? 8 : -8;
  const before = (await canvas()).controls.find((c) => c.name === "NameBox")?.left ?? 0;
  await api.act("designerDrag", { module: form, control: "NameBox", dx: step, dy: 0 });
  await until(async () => ((await canvas()).controls.find((c) => c.name === "NameBox")?.left ?? 0) !== before);
});

await reset();
await walk("resize by a handle", async (round) => {
  const step = round % 2 === 0 ? 6 : -6;
  const before = (await canvas()).controls.find((c) => c.name === "NameBox")?.width ?? 0;
  await api.act("designerResize", { module: form, control: "NameBox", edge: "se", dx: step, dy: 0 });
  await until(async () => ((await canvas()).controls.find((c) => c.name === "NameBox")?.width ?? 0) !== before);
});

await reset();
await walk("toolbox drop", async () => {
  const before = (await canvas()).controls.length;
  await api.act("designerToolbox", { module: form, kind: "Label", left: 250, top: 205 });
  await until(async () => (await canvas()).controls.length !== before);
});

// Each round needs its own victim, and making one is not part of what is being timed.
{
  const runs = [];
  for (let round = 0; round < ROUNDS; round++) {
    await api.act("designerToolbox", { module: form, kind: "Label", left: 250, top: 205 });
    const victim = (await canvas()).selected;
    const before = (await canvas()).controls.length;
    const started = Date.now();
    await api.act("designerDelete", { module: form, control: victim });
    await until(async () => (await canvas()).controls.length !== before);
    runs.push(Date.now() - started);
    await wait(120);
  }

  const sorted = [...runs].sort((a, b) => a - b);
  console.log(`${"delete".padEnd(34)} median ${String(sorted[2]).padStart(5)}ms   worst `
    + `${String(sorted[sorted.length - 1]).padStart(5)}ms   [${runs.join(", ")}]`);
}

// ---- typing: the debounced path, which SHOULD wait ----
await walk("typed edit -> draft preview", async (round) => {
  const source = await text();
  const edited = source.replace(/ToggleButton HoldToggle "Hold" at \d+,\d+/,
    `ToggleButton HoldToggle "Hold" at ${112 + (round % 2)},112`);
  await api.act("designerSetMarkup", { module: form, markup: edited });
  await until(async () => (await canvas()).draft === true);
});

// ---- the apply, and the workbook save behind it. A nudge makes the change; the timing is of
// the SAVE, from the keystroke to the form itself carrying it. ----
{
  const runs = [];
  const leftOf = async () => (await api.designer(form, project)).controls
    .find((c) => c.name === "NameBox")?.left ?? 0;
  await api.act("designerSelect", { module: form, control: "NameBox" });
  for (let round = 0; round < ROUNDS; round++) {
    const before = await leftOf();
    await api.ask('(() => { const el = document.querySelector(\'.designer-view[data-module="EntryForm"] .designer-canvas-scroll\'); el.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true })); return "sent"; })()');
    await until(async () => (await canvas()).dirty === true);
    const started = Date.now();
    await api.command("save");
    await until(async () => (await leftOf()) !== before);
    runs.push(Date.now() - started);
    await wait(150);
  }

  const sorted = [...runs].sort((a, b) => a - b);
  console.log(`${"Ctrl+S: apply + save".padEnd(34)} median ${String(sorted[2]).padStart(5)}ms   worst `
    + `${String(sorted[sorted.length - 1]).padStart(5)}ms   [${runs.join(", ")}]`);
}

// ---- what one projection COSTS, which the risks table asks for by name before anyone calls
// the designer's read path cheap: wrappers per read, with an idle tick as the floor. ----
const wrappers = async () => (await api.stats()).comWrappersGivenBack;
async function cost(label, times, once) {
  const before = await wrappers();
  const started = Date.now();
  for (let round = 0; round < times; round++) {
    await once();
  }

  const spent = Date.now() - started;
  const each = Math.round(((await wrappers()) - before) / times);
  console.log(`${label.padEnd(34)} ${String(each).padStart(5)} wrappers each, `
    + `${Math.round(spent / times)}ms each`);
}

console.log("");
await cost("COM cost: designer read", 10, () => api.designer(form, project));
await cost("COM cost: markup print", 10, () => api.designerMarkup(form, project));
await cost("COM cost: an idle tick", 10, () => api.ui());

// The form goes back as found, saved, so the next run measures the shape this one did.
await reset();
await api.command("save");
await until(async () => String(await api.designerMarkup(form, project)) === canonical);

const stats = await api.stats();
const level = stats.comWrappersGivenBack === stats.comWrappersDisposed;
console.log(`\nCOM ledger: ${stats.comWrappersGivenBack} given back, ${stats.comWrappersDisposed} `
  + `disposed - ${level ? "level" : "LEAKING, and com-leak.mjs will say where"}`);
console.log(`marshal to the host thread: ${JSON.stringify((await api.perf()).marshalMs)}`);
