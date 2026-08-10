/*
 * How the editor scales with module size.
 *
 * Every timing taken before this came from a module of a few dozen lines, which answers the
 * wrong question. What matters is not "is hover fast" but "is hover still fast in the module I
 * actually work in", and the honest answer to that is a curve rather than a number.
 *
 *   tools\New-PerfFixture.ps1
 *   node tools\harness\perf-scaling.mjs
 *
 * Read every figure against the two FLOORS printed first. `pagecall` is a script that answers
 * immediately; the promise floor is what the door costs to collect an answer that is not, which
 * every language feature is. A figure near the promise floor is a statement about the door and
 * not about the product - and a hover in the developer's own editor never crosses the door at
 * all, so what they wait for is the ANALYZER'S SHARE plus the page's work: the last column.
 *
 * Not a gate step. It takes a while, and a timing that runs on every commit is a timing nobody
 * reads.
 */

import { open } from "./xlide-api.mjs";

const api = await open({});
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const project = await api.project();

async function until(what, predicate, budgetMs = 30000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(250);
  }
  throw new Error(`timed out waiting for ${what}`);
}

const floor = await api.trip("pagecall", { n: 10 });

// TWO floors, and the second is the one that matters here.
//
// `pagecall` is a script that answers immediately. Every language feature answers a PROMISE, and
// the door collects a promise by polling - so an async route costs the call, the waits, and the
// poll that finds it settled, whatever the feature underneath did. Asking about a word that is
// not in the module exercises exactly that path with no analyzer work behind it.
//
// Without this line the curve reads as "every feature costs the same at every size", which is a
// statement about the door and not about the product. It was 77ms before the poll learned to
// back off, and the whole curve sat at 77 (2026-08-08).
const promiseFloor = await api.tripFeature("hover", { word: "ZZZnoSuchWordZZZ" }, { n: 8 });

console.log("the floors:");
console.log(`  a script into the page and back   ${String(floor.medianMs).padStart(6)}ms`);
console.log(`  a PROMISE collected by the door   ${String(promiseFloor.medianMs).padStart(6)}ms   <- every figure below contains this\n`);

const sizes = ["Small", "Medium", "Large", "Huge"];
const rows = [];

for (const module of sizes) {
  await api.pane("open", { module, project: project.projectId });
  await until(`${module} to be shown`, async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith("/" + module.toLowerCase()) ? ui : null;
  });

  // Let the analyzer settle: the first request after an open pays for work the rest do not,
  // and a curve built from first requests measures the seeding rather than the feature.
  await wait(2500);

  const text = (await api.readModule(module, project.projectId)).text ?? "";
  const lines = text.split(/\r?\n/).length;

  // A call site DEEP in the module, so the position is past everything the analyzer must scan.
  const prefix = module[0];
  const deep = `${prefix}${Math.max(0, Math.floor((lines / 8) * 0.8))}`;
  const word = text.includes(deep) ? deep : `${prefix}0`;

  // Timed INSIDE the page, so the door appears in none of these. The through-the-door figures
  // are kept alongside for one size only, at the end, to show how much of them was the harness.
  const hover = await api.timeFeature("hover", { word }, { n: 8 });
  const completions = await api.timeFeature("completions", { word }, { n: 8 });
  const definition = await api.timeFeature("definition", { word }, { n: 8 });

  // And the analyzer's own share of it, which the counters report separately.
  await api.perf({ reset: true });
  await api.timeFeature("hover", { word }, { n: 5 });
  const engine = (await api.engineCosts()).find((one) => one.method === "textDocument/hover");

  rows.push({
    module,
    lines,
    hover: hover.medianMs,
    completions: completions.medianMs,
    definition: definition.medianMs,
    engineHover: engine ? engine.medianMs : 0,
  });

  console.log(`  ${module.padEnd(8)} ${String(lines).padStart(6)} lines   ` +
    `hover ${String(hover.medianMs).padStart(5)}ms   ` +
    `completions ${String(completions.medianMs).padStart(5)}ms   ` +
    `definition ${String(definition.medianMs).padStart(5)}ms   ` +
    `(analyzer's share of hover: ${engine ? engine.medianMs : "?"}ms)`);
}

console.log("\nwhat the curve says:");

const smallest = rows[0];
const biggest = rows[rows.length - 1];
const growth = (what) => {
  const from = Math.max(1, smallest[what]);
  const to = biggest[what];
  const sizeRatio = biggest.lines / Math.max(1, smallest.lines);
  return `${what}: ${from}ms -> ${to}ms across ${sizeRatio.toFixed(0)}x the lines ` +
    `(${(to / from).toFixed(1)}x)`;
};

for (const what of ["hover", "completions", "definition", "engineHover"]) {
  console.log(`  ${growth(what)}`);
}

console.log(`\n  the promise floor is ${promiseFloor.medianMs}ms. A figure near it is the DOOR, not the feature.`);
console.log("  A hover in the developer's editor never crosses the door at all, so what they wait for is");
console.log("  the analyzer's share plus the page's own work: the last column, not the first three.");

// A stall a developer would feel, rather than a ratio.
const felt = rows.filter((one) => one.hover > 200 || one.completions > 200);
if (felt.length > 0) {
  console.log(`\n  OVER 200ms (a fifth of a second, which is felt): ` +
    felt.map((one) => `${one.module} at ${one.lines} lines`).join(", "));
} else {
  console.log("\n  nothing over 200ms at any size measured.");
}
