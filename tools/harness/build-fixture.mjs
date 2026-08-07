/*
 * Writes a fixture's components into the open workbook, through the debug api.
 *
 * The whole reason this exists rather than being six lines of PowerShell: VBComponents.Add through
 * `Workbook.VBProject` is exactly what "Trust access to the VBA project object model" gates, and
 * the add-in is already past that gate. So the components go in from INSIDE and the setting stays
 * off, which is where Microsoft ships it and where a machine is safer.
 *
 *   node build-fixture.mjs <plan.json>
 *
 * The plan is JSON — { modules: [{ name, kind, code }], sheetCode, openAtEnd } — because module
 * text is full of quotes, doubled quotes and CRLFs, and a file crosses once where a command line
 * would be escaped through two shells.
 */

import { readFileSync } from "node:fs";
import { open } from "./xlide-api.mjs";

const planPath = process.argv[2];
if (!planPath) {
  console.error("usage: node build-fixture.mjs <plan.json>");
  process.exit(1);
}

// The byte-order mark is stripped, because the usual producer is PowerShell 5.1 and its
// `-Encoding utf8` means "UTF-8 WITH a BOM". JSON.parse refuses one, with an error naming a
// character that does not appear to be there.
const plan = JSON.parse(readFileSync(planPath, "utf8").replace(/^﻿/, ""));
const api = await open({});

const health = await api.doctor();
if (!health.healthy) {
  console.error(`the door is not healthy: ${health.findings.join("; ")}`);
  process.exit(1);
}

/**
 * Writes a module and CHECKS that it took, because once it did not.
 *
 * The first write to a module that is freshly added AND currently shown fails — the editor
 * answers "Invalid procedure call or argument" — and the write route reports that in the log
 * rather than in its reply. Building a fixture on an unchecked write produced one with an empty
 * Rival, which meant the duplicate `Recalculate` was not there, which meant the fixture no longer
 * exercised the collision it exists for. It looked fine (2026-08-07).
 *
 * A second write always takes. So: write, read the line count back from the object model, and
 * write again if nothing arrived. Never report a module as written on the strength of having
 * asked for it.
 */
async function writeAndCheck(name, code) {
  const wanted = code.split(/\r\n|\n|\r/).filter((line, at, all) => at < all.length - 1 || line.length > 0).length;

  for (let attempt = 1; attempt <= 3; attempt++) {
    await api.writeModule(name, code);

    const project = await api.project();
    const lines = project.components.find((component) => component.name === name)?.lines ?? 0;

    if (lines > 0) {
      return { lines, attempts: attempt };
    }

    console.log(`  ${name.padEnd(14)} write ${attempt} did not take; asking again`);
  }

  throw new Error(`${name} would not accept its text after three attempts (wanted ${wanted} lines)`);
}

for (const module of plan.modules) {
  const added = await api.component("add", { kind: module.kind, name: module.name });
  const written = await writeAndCheck(added.name, module.code);
  const retried = written.attempts > 1 ? `  (took ${written.attempts} attempts)` : "";
  console.log(`  ${added.name.padEnd(14)} ${module.kind === 2 ? "class " : "module"}  ${written.lines} lines${retried}`);
}

/*
 * The first worksheet's own module. Its name is the CODE name, which is what the object model
 * calls it and what a document module is addressed by — not the tab caption a user sees, which
 * can differ and would address nothing.
 */
if (plan.sheetCode) {
  const project = await api.project();
  const sheet = project.components.find((component) => component.kind === "document"
    && component.name.toLowerCase() !== "thisworkbook");

  if (sheet) {
    const written = await writeAndCheck(sheet.name, plan.sheetCode);
    console.log(`  ${sheet.name.padEnd(14)} document  ${written.lines} lines`);
  } else {
    console.log("  (no worksheet document module found; the sheet case is not in this fixture)");
  }
}

/*
 * Which panes are open is part of the fixture. A module with no tab is the one a rename silently
 * misses, so leaving most of them closed is the arrangement the interesting test needs — and it
 * comes for free here, because adding a component through the door does not open a pane for it.
 */
if (plan.openAtEnd) {
  await api.caret(1, { module: plan.openAtEnd, column: 1 });
  console.log(`  opened ${plan.openAtEnd}, and left the rest closed`);
}

await api.command("save");

// Read it back from the object model rather than trusting the writes: the editor rewrites what it
// is given — it respells keywords and completes what it thinks is unfinished — so what is IN the
// workbook is the only thing worth reporting.
const built = await api.project();
const written = built.components.filter((component) => component.lines > 0);
console.log(`  saved; ${written.length} component(s) hold code`);

const missing = plan.modules
  .map((module) => module.name)
  .filter((name) => !built.components.some((component) => component.name === name));

if (missing.length > 0) {
  console.error(`these did not make it into the project: ${missing.join(", ")}`);
  process.exit(1);
}
