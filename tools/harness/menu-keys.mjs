/*
 * What the editor's context menu promises beside each command, read off the menu itself.
 *
 * A key drawn beside a command is a promise, and the browser's accelerator hook breaks some of
 * them silently: F2 is the Object Browser and F8 / Shift+F8 are Step Into and Step Over before
 * the page ever sees them, while the editor's own defaults put Rename on F2 and the next problem
 * on F8, and the menu drew those keys beside commands that could never receive them (#90). This
 * reads the menu the way a developer does, and holds every key it shows to one the page gets.
 *
 * The standalone editor draws its context view inside a shadow root, so the roots are searched.
 *
 *   node tools\harness\menu-keys.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

// What the native editor's hook takes before the page sees it (VbeCommands.ForKey and
// SurfaceCommandForKey), spelled as the menu spells keys. A command the page handles itself may
// show one of these only when it IS what the hook runs for that key.
const claimedByHost = new Map([
  ["F1", "Command Palette"],
  ["F2", null],
  ["Ctrl+F2", null],
  ["F5", "Run Sub/UserForm"],
  ["Ctrl+F5", null],
  ["Shift+F5", null],
  ["F8", null],
  ["Shift+F8", null],
  ["Ctrl+F8", "Run To Cursor"],
  ["Ctrl+Shift+F8", null],
  ["F9", "Toggle Breakpoint"],
  ["Ctrl+F9", null],
  ["Ctrl+Shift+F9", null],
  ["Ctrl+S", null],
  ["Ctrl+W", null],
  ["Ctrl+F4", null],
  ["Ctrl+L", "Call Stack"],
  ["Ctrl+G", "Immediate Window"],
  ["Ctrl+R", "Project Explorer"],
  ["F4", "Properties Window"],
]);

const expected = new Map([
  ["Rename Symbol", "Ctrl+Shift+R"],
  ["Change All Occurrences", ""],
  ["Refactor...", ""],
  ["Go to Definition", "Shift+F2"],
  ["Last Position", "Ctrl+Shift+F2"],
  ["Run Sub/UserForm", "F5"],
  ["Run To Cursor", "Ctrl+F8"],
  ["Toggle Breakpoint", "F9"],
  ["Call Stack", "Ctrl+L"],
  ["Command Palette", "F1"],
]);

console.log("the context menu's keys, read off the menu\n");

await api.pane("open", { module: "Runner", project: project.projectId });
await wait(800);
await api.caret(8, { module: "Runner", project: project.projectId, column: 5 });
await wait(300);

const rows = await api.ask(`(async () => {
  const ed = window.xlideBridge.workspace.activeEditor();
  ed.focus();
  ed.trigger("probe", "editor.action.showContextMenu", null);
  await new Promise((r) => setTimeout(r, 600));
  const roots = [document, ...[...document.querySelectorAll("*")].filter((e) => e.shadowRoot).map((e) => e.shadowRoot)];
  const items = [];
  for (const root of roots) {
    for (const item of root.querySelectorAll(".monaco-menu .action-item")) {
      const label = item.querySelector(".action-label")?.textContent?.trim() ?? "";
      const key = item.querySelector(".keybinding")?.textContent?.trim() ?? "";
      if (label) { items.push({ label, key }); }
    }
  }
  document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", code: "Escape", keyCode: 27, bubbles: true }));
  document.body.click();
  return items;
})()`) ?? [];

check("the context menu opened and could be read", rows.length > 5, `${rows.length} item(s)`);
for (const row of rows) { console.log(`  ${row.label.padEnd(34)} ${row.key}`); }

const shown = new Map(rows.map((row) => [row.label, row.key]));
for (const [label, key] of expected) {
  check(`${label} shows ${key === "" ? "no key" : key}`,
    shown.has(label) && shown.get(label) === key, `menu shows "${shown.get(label) ?? "(absent)"}"`);
}

for (const row of rows) {
  if (!row.key) { continue; }
  const owner = claimedByHost.has(row.key) ? claimedByHost.get(row.key) : "unclaimed";
  check(`${row.label}'s ${row.key} is a key the page gets`,
    owner === "unclaimed" || owner === row.label,
    owner === null ? `the host takes ${row.key} for itself` : `the host takes ${row.key} for ${owner}`);
}

process.exit(done());
