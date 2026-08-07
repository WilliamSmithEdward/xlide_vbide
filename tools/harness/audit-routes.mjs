/*
 * Is every route the shim serves documented, and reachable from the client?
 *
 * Asked of the SOURCE, not of anyone's memory. A route table is exactly the kind of thing that is
 * complete on the day it is written and quietly is not, six routes later — which is what it was
 * when this was written: the reference had all thirty-two, the driving guide had twenty, and one
 * route had no client method at all (2026-08-07).
 *
 *   node tools\harness\audit-routes.mjs
 *
 * Exits non-zero on any gap, so the gate can hold the line rather than a promise to.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..", "..");

const read = (relative) => readFileSync(join(root, relative), "utf8");

/**
 * The cases inside the switches on request.Route, and only those. The file holds other switches —
 * the assert vocabulary, the page-message kinds — whose cases are not routes, so matching
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

const routes = routesOf(read("src/Xlide.Vbe.Shim/AddIn/AddInSession.cs"));
if (routes.length === 0) {
  console.log("FAIL no routes found; the switch this reads may have moved");
  process.exit(1);
}

const reference = read("docs/debug-api.md");
const driving = read("docs/driving-excel.md");
const client = read("tools/harness/xlide-api.mjs");

const escaped = (route) => route.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// Each is looked for the way that document actually writes it: a row in the reference's route
// table, a row in the driving guide's route-to-client table, and a call in the client.
const inReference = (route) => new RegExp("^\\| `" + escaped(route) + "`", "m").test(reference);
const inDriving = (route) => new RegExp("^\\| `" + escaped(route) + "`", "m").test(driving);
const inClient = (route) => new RegExp("call\\(`?\"?" + escaped(route) + "[`\"$)]").test(client);

const gaps = [];
for (const route of routes) {
  if (!inReference(route)) { gaps.push(`${route}: not in docs/debug-api.md`); }
  if (!inDriving(route)) { gaps.push(`${route}: not in docs/driving-excel.md`); }
  if (!inClient(route)) { gaps.push(`${route}: no method in tools/harness/xlide-api.mjs`); }
}

if (gaps.length > 0) {
  console.log(`FAIL ${gaps.length} gap(s) across ${routes.length} routes:`);
  for (const gap of gaps) { console.log(`  ${gap}`); }
  process.exit(1);
}

console.log(`ok   all ${routes.length} routes are in both documents and reachable from the client`);
