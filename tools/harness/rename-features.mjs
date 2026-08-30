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

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});

const { check, done } = reporter();

async function until(what, predicate, budgetMs = 15000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(200);
  }
  throw new Error(`timed out waiting for ${what}`);
}

/*
 * THE WORKBOOK THAT HOLDS THE FIXTURE, FOUND, not the one in front.
 *
 * `api.project()` answers about the workbook on screen, and with two open - which is the state
 * the gate runs this in, beside TwinFixture - which is in front is a coin-flip Excel decides
 * when it opens two files on one command line. The old guard trusted project() and THREW when
 * it lost the flip ("TwinFixture has no Consumer"), which turned an ordering accident into a
 * red suite. Asking every open project which one holds these modules removes the flip: the one
 * that has Consumer, Watcher and Rival is this suite's, whichever is on screen.
 */
const NEEDED = ["Helpers", "Consumer", "Watcher", "Rival", "HelpersExtra"];
let project = null;
for (const candidate of (await api.projects()).projects) {
  const held = await Promise.all(NEEDED.map((name) =>
    api.readModule(name, candidate.projectId).then((m) => m.text != null).catch(() => false)));
  if (held.every(Boolean)) { project = candidate; break; }
}
if (project === null) {
  // Thrown rather than process.exit: exiting with the client's socket still open trips a libuv
  // assertion, which buries the one line worth reading under a crash.
  throw new Error(
    "this suite needs RenameFixture.xlsm open, and no workbook holds its modules "
    + `(${NEEDED.join(", ")}). Open it and run this again:\n`
    + "  tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\fixtures\\RenameFixture.xlsm");
}

const textOf = async (module) => (await api.readModule(module, project.projectId)).text ?? "";

/**
 * The native pane, the surface and the page all naming the same module.
 *
 * Checked after everything that moves the editor. The page agreeing with itself proves only
 * that the page is consistent; what a Run acts on is the pane underneath.
 */
async function parity(after) {
  const sync = await api.inSync();
  check(`native parity ${after.padEnd(28)}`, sync.agreed,
    `native=${sync.nativeModule} surface=${sync.surfaceModule} page=${sync.pageModule}`);
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

// The reply is read now rather than printed. It used to answer a bare true whatever happened, so
// the only way to learn that an undo had stopped halfway was the two module reads below - which
// say the text is wrong and cannot say why. `stopped` names the module that refused.
check("the undo ran to the end", undone.undone === true, undone.stopped ?? JSON.stringify(undone));

// NAMED, not counted. This asserted a count of two and the rename touches three: Helpers declares
// the procedure, Consumer qualifies it, and Watcher names it as well - Watcher exists precisely
// because a rename has to update the tab being looked at as well as the ones that are not. A count
// is the wrong assertion for a cross-module operation anyway: it passes for the right number of
// wrong modules, and it says nothing about WHICH the rename reached, which is the whole question.
const putBack = [...(undone.modules ?? [])].sort();
check("and it put back every module the rename touched",
  JSON.stringify(putBack) === JSON.stringify(["Consumer", "Helpers", "Watcher"]),
  `${JSON.stringify(undone.modules)} - expected Helpers (the declaration), Consumer (a qualified `
  + "call) and Watcher (a call in the module left open on purpose)");

await wait(3000);

const restored = {
  Helpers: await textOf("Helpers"),
  Consumer: await textOf("Consumer"),
};
check("undo restored the declaration", restored.Helpers.trim() === before.Helpers.trim());
check("undo restored the caller", restored.Consumer.trim() === before.Consumer.trim());
await parity("after the rename and its undo");

/*
 * AND PLAIN UNDO IS HOW A DEVELOPER REACHES IT.
 *
 * `undoRename` above is the door's own route. This is the path Ctrl+Z, the toolbar button and the
 * wrench menu all take, and until 2026-08-30 it was not this one: the reversal lived in a context
 * menu (the owner: "i think it should be undo like ctrl z, like other things"), and the key
 * everybody does press did something worse than nothing. A rename arrives as ordinary adopted
 * text, so it sits on each model's undo stack - Ctrl+Z reversed the SHOWN module's share and left
 * every other module renamed, which is a half-renamed project.
 *
 * So the check is not that undo did something. It is that it reached the modules that are not on
 * screen, because those are the ones a per-model undo silently leaves behind.
 */
console.log("\nand again, to take it back with plain undo:");
await showing("Helpers");
const again = await api.act("rename", { word: "Recalculate", newName: "Recompute" });
check("a second rename was accepted", again.did, again.detail);
await wait(3000);

await showing("Helpers");
await api.act("undo");
await wait(3000);

const byUndo = {
  Helpers: await textOf("Helpers"),
  Consumer: await textOf("Consumer"),
};
check("plain undo put back the module in front of the developer",
  byUndo.Helpers.trim() === before.Helpers.trim(),
  byUndo.Helpers.split(/\r?\n/).find((l) => /Sub Rec/.test(l)));
check("AND the module that was not on screen, which a per-model undo would have left renamed",
  byUndo.Consumer.trim() === before.Consumer.trim(),
  byUndo.Consumer.split(/\r?\n/).find((l) => l.includes("Helpers.")));
await parity("after the rename and a plain undo");

console.log("\ngo to definition, from the qualified call:");
await showing("Consumer");
const definition = await api.act("definition", { word: "Recalculate" });
console.log(`  ${JSON.stringify(definition).slice(0, 260)}`);
check("definition resolves from a call site", definition.did, definition.detail);
await parity("after resolving a definition");

console.log("\nfind all references, as a list and as the dialog:");
const refList = await api.act("references", { word: "Recalculate" });
console.log(`  ${JSON.stringify(refList).slice(0, 200)}`);
// The list crosses modules the way the declaration does: the qualified call in Consumer, the
// call in Watcher, and the declaration in Helpers are all uses of this symbol.
check("references lists the symbol's uses across modules",
  refList.did && Array.isArray(refList.data) && refList.data.length >= 2,
  `${(refList.data ?? []).length} reference(s)`);

// The DATA form leaves nothing on screen; the dialog is the feature, and this is what the
// audit's A16 was about - the act stopped one step short of the thing it is named for.
const dialogsBeforeOpen = (await api.ui()).dialogs.map((d) => d.id);
check("the data form leaves no dialog standing",
  !dialogsBeforeOpen.some((id) => /references/i.test(id)), dialogsBeforeOpen.join(","));

const refOpen = await api.act("references", { word: "Recalculate", open: 1 });
console.log(`  ${JSON.stringify(refOpen).slice(0, 200)}`);
check("references open=1 reports it opened the dialog", refOpen.did, refOpen.detail);
const dialogsAfterOpen = await until("the references dialog to stand", async () => {
  const open = (await api.ui()).dialogs.map((d) => d.id);
  return open.some((id) => /references/i.test(id)) ? open : null;
}).catch(() => []);
check("and ui.dialogs sees the references dialog on screen",
  dialogsAfterOpen.some((id) => /references/i.test(id)), dialogsAfterOpen.join(","));

await api.act("closeDialogs");
await until("the dialog to be gone again", async () =>
  !(await api.ui()).dialogs.some((d) => /references/i.test(d.id))).catch(() => {});
await parity("after the references dialog");

process.exit(done());
