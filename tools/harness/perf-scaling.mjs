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

import { open, wait } from "./xlide-api.mjs";

const api = await open({});
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

// MASSIVE IS WHY THE FIXTURE EXISTS and it was not being measured. Its generator says so in as
// many words: the four sizes below it all fit comfortably and so agree with each other about
// what is fast, and anything quadratic in module size is invisible at 11,000 lines and obvious
// at VBA's per-module ceiling. Opening it costs the EDITOR seconds of its own parse, which the
// wait below absorbs; the figures are the page's and the analyzer's, not the open's.
const sizes = ["Small", "Medium", "Large", "Huge", "Massive"];
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
  //
  // TAKEN FROM THE TEXT, not from the module's name. The old spelling built the word out of the
  // module's first letter - which holds for four of the five sizes and not for Massive, whose
  // procedures are prefixed V: every feature was asked about `M0`, a word no module has, and
  // answered `did: false` (2026-08-21). A curve is only about size if the probe exists at every
  // size, so it is read out of the source itself.
  const declared = [...text.matchAll(/^(?:Public\s+|Private\s+)?(?:Function|Sub)\s+(\w+)/gim)]
    .map((found) => found[1]);
  if (declared.length === 0) {
    throw new Error(`${module} declares no procedure to probe`);
  }

  const word = declared[Math.min(declared.length - 1, Math.floor(declared.length * 0.8))];

  // AND THE PAGE MUST ACTUALLY HOLD IT. The tab reports the new model the moment it switches,
  // seconds before a ceiling-sized module's text is in it - so at 64,802 lines every feature
  // was asked about a word the model did not have yet, declined, and printed as `undefined` in
  // a curve that then read NaN and still said "nothing over 200ms" (2026-08-21).
  await until(`${module} to hold ${word} in the page's own model`,
    async () => (await api.at(word)) ?? null, 120000);

  // Timed INSIDE the page, so the door appears in none of these. The through-the-door figures
  // are kept alongside for one size only, at the end, to show how much of them was the harness.
  // A feature that declined has no timing, and a curve that prints its absence as a number is
  // worse than one that stops: it read "undefined" and carried on concluding.
  const timed = async (what) => {
    const answer = await api.timeFeature(what, { word }, { n: 8 });
    if (typeof answer?.medianMs !== "number") {
      throw new Error(`${what} in ${module} answered no timing: ${JSON.stringify(answer)}`);
    }

    return answer;
  };

  const hover = await timed("hover");
  const completions = await timed("completions");
  const definition = await timed("definition");

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
