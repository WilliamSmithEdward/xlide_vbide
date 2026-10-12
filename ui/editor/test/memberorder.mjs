// The explorer's member orders (#92): as in code, by name, by kind, and a property's legs together
// under every one of them. Called rather than looked at, because the order is the whole behaviour.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scratch = await mkdtemp(path.join(tmpdir(), "xlide-memberorder-"));
const out = path.join(scratch, "memberorder.mjs");

await build({
  entryPoints: [path.join(root, "src", "memberorder.ts")],
  outfile: out,
  bundle: true,
  format: "esm",
  platform: "neutral",
});

const { orderMembers } = await import(pathToFileURL(out).href);

let passed = 0;
let failed = 0;
const check = (name, got, want) => {
  try {
    assert.deepEqual(got, want);
    passed++;
    console.log(`ok   ${name}`);
  } catch {
    failed++;
    console.log(`FAIL ${name}\n     got  ${JSON.stringify(got)}\n     want ${JSON.stringify(want)}`);
  }
};

const module = [
  { name: "Total", kind: "Property Let", line: 3 },
  { name: "Colour", kind: "Enum", line: 8 },
  { name: "run", kind: "Sub", line: 14 },
  { name: "Point", kind: "Type", line: 20 },
  { name: "Area", kind: "Function", line: 25 },
  { name: "Total", kind: "Property Get", line: 30 },
  { name: "Apply", kind: "Sub", line: 36 },
  { name: "Total", kind: "Property Set", line: 40 },
];
const names = (rows) => rows.map((row) => `${row.kind} ${row.name}`);

check("as in code keeps the written order, a property's legs together where the first stood",
  names(orderMembers(module, "code")),
  ["Property Get Total", "Property Let Total", "Property Set Total", "Enum Colour", "Sub run",
    "Type Point", "Function Area", "Sub Apply"]);

check("by name is alphabetical, ignoring case",
  names(orderMembers(module, "name")),
  ["Sub Apply", "Function Area", "Enum Colour", "Type Point", "Sub run",
    "Property Get Total", "Property Let Total", "Property Set Total"]);

check("by kind puts Enums, Types, Subs, Functions, then Properties, each by name",
  names(orderMembers(module, "kind")),
  ["Enum Colour", "Type Point", "Sub Apply", "Sub run", "Function Area",
    "Property Get Total", "Property Let Total", "Property Set Total"]);

const before = JSON.stringify(module);
orderMembers(module, "name");
check("the input is not reordered", JSON.stringify(module), before);

check("a module with nothing in it lists nothing", orderMembers([], "kind"), []);

await rm(scratch, { recursive: true, force: true });
console.log(`\n${passed}/${passed + failed} passed`);
process.exit(failed === 0 ? 0 : 1);
