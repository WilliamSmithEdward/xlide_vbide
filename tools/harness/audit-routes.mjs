/*
 * Is every route the shim serves documented, reachable from the client, and DRIVEN by something?
 *
 * Asked of the SOURCE, not of anyone's memory. A route table is exactly the kind of thing that is
 * complete on the day it is written and quietly is not, six routes later - which is what it was
 * when this was written: the reference had all thirty-two, the driving guide had twenty, and one
 * route had no client method at all (2026-08-07).
 *
 * DOCUMENTED AND REACHABLE IS NOT COVERED, and this said so for two days without noticing the
 * difference. A route with a doc row and a client method can still be a route nothing has ever
 * called: measured 2026-08-09, two of the forty-nine were exactly that. So the third question is
 * asked too, and a route left out has to be left out BY NAME with a reason, which is the only kind
 * of exemption that cannot rot quietly. A name that turns out to be driven after all fails as
 * loudly as one that is not, because a stale excuse reads like a considered one.
 *
 *   node tools\harness\audit-routes.mjs
 *
 * Exits non-zero on any gap, so the gate can hold the line rather than a promise to.
 */

import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..", "..");

const read = (relative) => readFileSync(join(root, relative), "utf8");

/**
 * The cases inside the switches on request.Route, and only those. The file holds other switches -
 * the assert vocabulary, the page-message kinds - whose cases are not routes, so matching
 * `case "..."` across the file answers a different question by nine extra names.
 */
function routesOf(source) {
  const lines = source.split(/\r?\n/);
  const found = new Set();

  lines.forEach((line, at) => {
    if (!/switch \(request\.Route\)/.test(line)) { return; }

    let depth = 0;
    let opened = false;

    for (let i = at; i < lines.length; i++) {
      for (const character of lines[i]) {
        if (character === "{") { depth++; opened = true; } else if (character === "}") { depth--; }
      }

      if (opened && depth <= 0) { break; }

      const route = /^\s*case "([a-zA-Z]+)"/.exec(lines[i]);
      if (route && depth <= 1) { found.add(route[1]); }
    }
  });

  return [...found].sort();
}

/*
 * EVERY PART OF THE SESSION, not the file the switch happened to be in.
 *
 * This read AddInSession.cs alone, and the routes moved to AddInSession.DebugApi.cs when the debug
 * api was split off as a partial (2026-08-09). The check said so rather than reporting zero gaps
 * across zero routes, which is the failure it was built to avoid; reading the whole folder means
 * the next split does not need it edited at all.
 */
const sessionParts = readdirSync(join(root, "src/Xlide.Vbe.Shim/AddIn"))
  .filter((file) => file.startsWith("AddInSession") && file.endsWith(".cs"))
  .map((file) => read(`src/Xlide.Vbe.Shim/AddIn/${file}`))
  .join("\n");

const routes = routesOf(sessionParts);
if (routes.length === 0) {
  console.log("FAIL no routes found; the switch this reads may have moved");
  process.exit(1);
}

const reference = read("docs/xlide-api.md");
const driving = read("docs/driving-excel.md");
const client = read("tools/harness/xlide-api.mjs");

const escaped = (route) => route.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// Each is looked for the way that document actually writes it: a row in the reference's route
// table, a row in the driving guide's route-to-client table, and a call in the client.
const inReference = (route) => new RegExp("^\\| `" + escaped(route) + "`", "m").test(reference);
const inDriving = (route) => new RegExp("^\\| `" + escaped(route) + "`", "m").test(driving);
const inClient = (route) => new RegExp("call\\(`?\"?" + escaped(route) + "[`\"$)]").test(client);

/**
 * The client methods that reach each route.
 *
 * Taken by walking back from each `call(...)` to the nearest property or method definition above
 * it, because the client writes them three ways - `name: (args) => call(...)`, `async name(args) {`
 * and a multi-line body with the call some lines down. A regex that knew only the first shape read
 * `act` and `ui` as unreachable while four suites were calling them, which is the wrong answer in
 * the direction that matters: it would have sent somebody to write a method that already existed.
 */
function methodsByRoute(client) {
  const lines = client.split(/\r?\n/);
  const found = new Map();

  lines.forEach((line, at) => {
    const call = /call\(`?"?([a-zA-Z]+)/.exec(line);
    if (!call) { return; }

    // At the object's own indent, exactly. Anything deeper is a statement inside a method body,
    // and `if (...)` two lines above the call reads as a property named `if` otherwise.
    for (let i = at; i >= 0 && at - i < 30; i--) {
      const named = /^ {4}(?:async )?([a-zA-Z][a-zA-Z0-9]*)\s*[:(]/.exec(lines[i]);
      if (!named) { continue; }

      if (!found.has(call[1])) { found.set(call[1], new Set()); }
      found.get(call[1]).add(named[1]);
      return;
    }
  });

  return found;
}

/*
 * ROUTES NOTHING DRIVES, ON PURPOSE. Each one names why, and the name is checked: an entry here
 * for a route that IS driven fails, so this list cannot quietly become a list of things somebody
 * once meant to get round to.
 */
const NOT_DRIVEN_ON_PURPOSE = {
  trip:
    "a measurement, not an assertion. It reports wall clock from asking to observable across the "
    + "page/host boundary, and the only honest pass/fail for a latency figure is a threshold "
    + "somebody has to keep true on every machine the gate runs on. Its driver is "
    + "perf-scaling.mjs, which needs a fourth fixture and exists to be READ rather than to go "
    + "green. Run it when the surface feels slow: `node tools\\harness\\perf-scaling.mjs`.",

  drainfinalizers:
    "a bisecting tool, not an assertion. Run an operation, call it, and if the host dies THAT "
    + "operation leaked a wrapper. As a leak CHECK it is a false-negative machine - measured "
    + "against a build with 8,734 wrappers pending it reported clean and the host lived, because "
    + "releasing an apartment-threaded object from the finalizer thread is only sometimes fatal. "
    + "com-leak.mjs counts instead, which is deterministic. Nothing should call this in a suite.",
};

/*
 * DRIVEN MEANS DRIVEN BY SOMETHING THE GATE RUNS.
 *
 * This used to read every file in tools/harness, which answers a different and much easier
 * question: whether a route is mentioned in a file that exists. Seven suites in that folder are
 * run by nothing at all, so a route whose only caller was one of them passed this check for as
 * long as it took somebody to notice - and the whole purpose of the check is that nobody has to.
 *
 * The gate script is the list, read rather than restated: anything named in verify.ps1 counts, and
 * a suite added there starts counting the moment it is added.
 */
const gate = read("tools/verify.ps1");
const runByTheGate = readdirSync(join(root, "tools/harness"))
  .filter((file) => /\.(mjs|ps1)$/.test(file) && file !== "xlide-api.mjs" && file !== "audit-routes.mjs")
  .filter((file) => gate.includes(file));

const corpus = runByTheGate
  .map((file) => readFileSync(join(root, "tools/harness", file), "utf8"))
  .join("\n");

/*
 * Receivers that are not the api client. `.log(` matched 203 `console.log(` calls and 5 real ones,
 * so the `log` route counted as driven by every suite that prints anything - and would have gone
 * on counting after the last real caller was deleted. The route names here are ordinary words, so
 * this will keep mattering as the client grows.
 */
const NOT_THE_CLIENT = ["console", "JSON", "Math", "Object", "Array", "String", "Number", "Date",
  "process", "performance", "window", "document"];

const methods = methodsByRoute(client);
const isDriven = (route) =>
  [...(methods.get(route) ?? [])].some((method) =>
    new RegExp(`(?<!${NOT_THE_CLIENT.join("|")})\\.${method}\\s*\\(`).test(corpus))
  // The PowerShell probes build the URL themselves rather than going through the client.
  || new RegExp(`/${escaped(route)}[?"'\`\\s]`).test(corpus);

const gaps = [];
for (const route of routes) {
  if (!inReference(route)) { gaps.push(`${route}: not in docs/xlide-api.md`); }
  if (!inDriving(route)) { gaps.push(`${route}: not in docs/driving-excel.md`); }
  if (!inClient(route)) { gaps.push(`${route}: no method in tools/harness/xlide-api.mjs`); }

  const driven = isDriven(route);
  const excused = Object.hasOwn(NOT_DRIVEN_ON_PURPOSE, route);

  if (!driven && !excused) {
    gaps.push(`${route}: nothing in tools/harness drives it. Drive it, or name it in `
      + "NOT_DRIVEN_ON_PURPOSE with the reason");
  }

  if (driven && excused) {
    gaps.push(`${route}: excused as never driven, but something drives it. Remove the excuse`);
  }
}

if (gaps.length > 0) {
  console.log(`FAIL ${gaps.length} gap(s) across ${routes.length} routes:`);
  for (const gap of gaps) { console.log(`  ${gap}`); }
  process.exit(1);
}

const excused = Object.keys(NOT_DRIVEN_ON_PURPOSE);
console.log(`ok   all ${routes.length} routes are in both documents and reachable from the client`);
console.log(`ok   ${routes.length - excused.length} are driven by one of the ${runByTheGate.length} `
  + `suites the gate runs; ${excused.length} left out on purpose: ${excused.join(", ")}`);
