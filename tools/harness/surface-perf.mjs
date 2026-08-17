/*
 * The rest of the surface, timed the way designer-perf.mjs times the designer: every interaction
 * from the act to the moment the surface has ANSWERED, five times, reported as median and worst.
 *
 * The designer walk exists because the owner felt a lag in two gestures; this one exists because
 * "the whole surface" was the ask and the designer is one half of it. Same shape deliberately:
 * a reader comparing the two tables should not have to ask whether the numbers mean the same
 * thing.
 *
 * Run against DebugFixture.xlsm, which holds real modules with real code - the language rows
 * measure an answer rather than a miss, and the probe word is taken from the module rather than
 * assumed:
 *
 *   tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\DebugFixture.xlsm -Fresh
 *   node tools\harness\surface-perf.mjs
 *
 * It PRINTS rather than passes or fails, for the reason its sibling gives: what counts as slow
 * is a judgement about a person's patience, and a threshold nobody can defend becomes a number
 * the next person raises. The COM ledger and the page's own main-thread stalls come at the end,
 * because those two ARE pass/fail and this walk is a good moment to read them.
 */
import { open } from "./xlide-api.mjs";

const api = await open();
const { projectId: project, components } = await api.project();
const ROUNDS = 5;
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function until(predicate, budgetMs = 20000) {
  const started = Date.now();
  for (;;) {
    if (await predicate()) return;
    if (Date.now() - started > budgetMs) throw new Error("timed out");
  }
}

async function walk(label, once) {
  const runs = [];
  for (let round = 0; round < ROUNDS; round++) {
    const started = Date.now();
    await once(round);
    runs.push(Date.now() - started);
    await wait(120);
  }

  const sorted = [...runs].sort((a, b) => a - b);
  console.log(`${label.padEnd(34)} median ${String(sorted[2]).padStart(5)}ms   worst `
    + `${String(sorted[sorted.length - 1]).padStart(5)}ms   [${runs.join(", ")}]`);
}

const modules = components.filter((one) => one.kind === "module" || one.kind === "class");
const first = modules[0]?.name ?? components[0].name;
const second = modules[1]?.name ?? first;

// A word the module actually holds, so the language rows measure an answer rather than a miss.
const probeText = String((await api.readModule(first, project)).text ?? "");
const probeWord = (/\b(Sub|Function|Dim|Option)\b/.exec(probeText) ?? ["Sub"])[0];

console.log(`\nthe rest of the surface, against ${project} (${first}, on "${probeWord}")\n`);

await walk("open a module tab", async () => {
  await api.pane("open", { module: first });
  await until(async () => (await api.state()).shownModule === first);
  await api.pane("close", { module: first });
  await until(async () => (await api.state()).shownModule !== first);
});

// The close on its own: the open before it is setup, not part of what is being timed.
{
  const runs = [];
  for (let round = 0; round < ROUNDS; round++) {
    await api.pane("open", { module: first });
    await until(async () => (await api.state()).shownModule === first);
    const started = Date.now();
    await api.pane("close", { module: first });
    await until(async () => (await api.state()).shownModule !== first);
    runs.push(Date.now() - started);
    await wait(120);
  }

  const sorted = [...runs].sort((a, b) => a - b);
  console.log(`${"close a module tab".padEnd(34)} median ${String(sorted[2]).padStart(5)}ms   worst `
    + `${String(sorted[sorted.length - 1]).padStart(5)}ms   [${runs.join(", ")}]`);
}

await api.pane("open", { module: first });
await api.pane("open", { module: second });
await until(async () => (await api.state()).shownModule === second);

await walk("switch between two tabs", async (round) => {
  const want = round % 2 === 0 ? first : second;
  await api.act("activate", { module: want });
  await until(async () => (await api.state()).shownModule === want);
});

await walk("read a module's text", async () => { await api.readModule(first, project); });

await walk("completions where a hand asks", async () => {
  await api.act("completions", { word: probeWord });
});

await walk("hover", async () => { await api.act("hover", { word: probeWord }); });

await walk("the project tree", async () => { await api.project(); });

await walk("a module-scope search", async () => {
  await api.act("search", { query: probeWord, scope: "module" });
  await api.act("search", { close: 1 });
});

// The panel walks a component's whole property bag through its type library, so this is the
// COM-heaviest thing on the surface that is not the designer's own projection. Measured through
// the ACTIVATE that aims it, so the honest reading is this line against the tab switch above:
// the difference is the panel's own share.
await walk("activate -> the panel follows", async (round) => {
  const want = round % 2 === 0 ? first : second;
  await api.act("activate", { module: want });
  await until(async () => (await api.ui()).properties?.component === want);
});

await walk("an idle snapshot", async () => { await api.ui(); });

const stats = await api.stats();
console.log(`\nCOM ledger: ${stats.comWrappersGivenBack} given back, ${stats.comWrappersDisposed} disposed`);
const ui = await api.ui();
console.log(`page main-thread stalls over 50ms: ${JSON.stringify(ui.longTasks ?? [])}`);
console.log(`models held: ${JSON.stringify(ui.census ?? {})}`);
