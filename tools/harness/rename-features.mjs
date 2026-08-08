/*
 * Rename and Go to Definition, driven through the providers the editor uses.
 *
 * Run against the rename fixture, which exists for these cases:
 *   tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\fixtures\\RenameFixture.xlsm
 *   node tools\\harness\\rename-features.mjs
 *
 * PARITY WITH THE NATIVE EDITOR is checked at every step, because a feature that touches the
 * editor is not tested until it is: rename rewrites modules, and the host's own panes are what
 * Run, Step and ToggleBreakpoint act on afterwards.
 *
 * Rename, driven for the first time.
 *
 * It is the feature that rewrites the developer's code across every module using a symbol, and
 * until now nothing but a hand on F2 could start one. The rename fixture exists for these cases:
 * a qualified call, a bare call that is ambiguous and must be LEFT ALONE, a module whose name
 * merely begins the same way, and a module with no tab open.
 *
 * Every case puts the workbook back with undoRename, and checks that it went back.
 */

import { open } from "./xlide-api.mjs";

const api = await open({});
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const project = await api.project();

const broken = [];
let checks = 0;
const check = (what, ok, detail) => {
  checks++;
  console.log(`  ${ok ? "ok  " : "FAIL"} ${what}${detail ? "  -- " + detail : ""}`);
  if (!ok) { broken.push(what); }
};

async function until(what, predicate, budgetMs = 15000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(200);
  }
  throw new Error(`timed out waiting for ${what}`);
}

const textOf = async (module) => (await api.readModule(module, project.projectId)).text ?? "";

/**
 * The native pane, the surface and the page all naming the same module.
 *
 * Checked after everything that moves the editor. The page agreeing with itself proves only
 * that the page is consistent; what a Run acts on is the pane underneath.
 */
async function parity(after) {
  checks++;
  const sync = await api.inSync();
  console.log(`  ${sync.agreed ? "ok  " : "FAIL"} native parity ${after.padEnd(28)} native=${sync.nativeModule} surface=${sync.surfaceModule} page=${sync.pageModule}`);
  if (!sync.agreed) { broken.push(`native parity ${after}`); }
}

async function showing(module) {
  await api.pane("open", { module, project: project.projectId });
  await until(`${module} to be shown`, async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith("/" + module.toLowerCase()) ? ui : null;
  });
}

console.log("rename, against the fixture built for it\n");

const before = {
  Helpers: await textOf("Helpers"),
  Consumer: await textOf("Consumer"),
  HelpersExtra: await textOf("HelpersExtra"),
  Rival: await textOf("Rival"),
};

await showing("Helpers");
await parity("with Helpers shown");

console.log("renaming Recalculate from its declaration in Helpers:");
const said = await api.act("rename", { word: "Recalculate", newName: "Recompute" });
console.log(`  ${JSON.stringify(said)}`);
check("the rename was accepted", said.did, said.detail);

await wait(3000);

const after = {
  Helpers: await textOf("Helpers"),
  Consumer: await textOf("Consumer"),
  HelpersExtra: await textOf("HelpersExtra"),
  Rival: await textOf("Rival"),
};

check("the declaration was renamed",
  after.Helpers.includes("Recompute") && !after.Helpers.includes("Sub Recalculate"),
  after.Helpers.split(/\r?\n/).find((l) => /Sub Rec/.test(l)));

check("the QUALIFIED call in Consumer followed",
  after.Consumer.includes("Helpers.Recompute"),
  after.Consumer.split(/\r?\n/).find((l) => l.includes("Helpers.")));

// The whole point of the fixture: two modules declare Recalculate, so a bare call cannot be
// proved to mean either, and must be left exactly as it was.
check("the AMBIGUOUS bare call was left alone",
  after.Consumer.includes('Recalculate "ambiguous"'),
  after.Consumer.split(/\r?\n/).find((l) => /"ambiguous"/.test(l)));

check("Rival's own Recalculate was not touched",
  after.Rival === before.Rival);

check("HelpersExtra, whose name merely starts the same, was not touched",
  after.HelpersExtra === before.HelpersExtra);

console.log("\nputting it back:");
const undone = await api.undoRename();
console.log(`  ${JSON.stringify(undone).slice(0, 200)}`);
await wait(3000);

const restored = {
  Helpers: await textOf("Helpers"),
  Consumer: await textOf("Consumer"),
};
check("undo restored the declaration", restored.Helpers.trim() === before.Helpers.trim());
check("undo restored the caller", restored.Consumer.trim() === before.Consumer.trim());
await parity("after the rename and its undo");

console.log("\ngo to definition, from the qualified call:");
await showing("Consumer");
const definition = await api.act("definition", { word: "Recalculate" });
console.log(`  ${JSON.stringify(definition).slice(0, 260)}`);
check("definition resolves from a call site", definition.did, definition.detail);
await parity("after resolving a definition");

console.log(`\n${checks} checks, ${broken.length} broken`);
for (const one of broken) { console.log("  ! " + one); }

process.exit(broken.length === 0 ? 0 : 1);
