/*
 * The Locals rows the editor's accessibility tree loses.
 *
 * Below an expanded object, the native Locals window paints every row, and its accessibility
 * tree stops: expanding a Workbook at the top of the list took every local after it out of the
 * panel. Its child lists and parent links are wrong there too, so the reader takes rows from the
 * list's own chain, recovers the rest by selecting them in the hidden window, and works out
 * each row's depth itself (LocalsReader, mapped 2026-10-17).
 *
 * This expands and collapses the blocks that hide rows - an object at the top, a Collection
 * recovered from below one, a member far out of view, the same object again after a collapse -
 * and checks every top-level row, every depth, and that a step reaches a value below them.
 * Nothing here expands a Worksheet: the native expand reads every property, and
 * Worksheet.MailEnvelope starts Outlook.
 */

import { open, reporter, waitFor, scratchModule } from "./xlide-api.mjs";

const api = await open({});
const project = (await api.projects()).projects[0];
const { check, done } = reporter();

const name = `LocalsTail${process.pid}`;
const scratch = scratchModule(api, project.projectId, name);
const code = [
  "Option Explicit",
  "Public Sub Inspect()",
  "    Dim wb As Workbook",
  "    Set wb = ThisWorkbook",
  "    Dim a As Long, b As String",
  "    a = 7: b = \"bee\"",
  "    Dim c As Collection: Set c = New Collection",
  "    c.Add \"first\": c.Add \"second\"",
  "    Dim n As Long: n = 3",
  "    Stop",
  "    a = 8",
  "    n = 4",
  "End Sub",
  ""].join("\r\n");

const rows = async () => (await api.locals()).rows ?? [];
// The module's own row leads the list, as it does natively; the locals follow it.
const top = (list) => list.filter((r) => r.depth === 0 && r.expression !== name).map((r) => r.expression + (r.expanded ? "+" : "")).join(",");
const value = (list, local) => list.find((r) => r.depth === 0 && r.expression === local)?.value;

async function expect(what, predicate, budgetMs = 8000) {
  try {
    await waitFor(what, async () => predicate(await rows()), { budgetMs });
    check(what, true, top(await rows()));
  } catch {
    const list = await rows();
    check(what, false, `${top(list)}, ${list.length} rows`);
  }
}

// The page draws a reading on the host's next poll, and a click on the old drawing carries an
// old row index, which the host rightly refuses - so the click waits for the page to catch up.
async function toggle(expression, depth) {
  await waitFor("the page to draw the reading", async () =>
    (await api.ask("document.querySelectorAll('#locals-table .locals-row:not(.locals-header)').length")) === (await rows()).length,
    { budgetMs: 5000 });
  const clicked = await api.ask(`(() => {
    const row = [...document.querySelectorAll('#locals-table .locals-row:not(.locals-header)')].find((r) =>
      r.querySelector('.locals-toggle')?.getAttribute('aria-label')?.split(' ')[1] === ${JSON.stringify(expression)}
      && r.querySelector('.locals-name')?.style.paddingLeft === '${depth * 14}px');
    const button = row?.querySelector('.locals-toggle');
    button?.click();
    return !!button;
  })()`);
  if (!clicked) throw new Error(`no expander for ${expression} at depth ${depth}`);
}

console.log("Locals rows below an expanded object\n");

try {
  await api.component("add", { name, kind: "module", project: project.projectId });
  await api.writeModule(name, code, project.projectId);
  await waitFor("the break", async () => {
    await api.caret(3, { module: name, project: project.projectId });
    await new Promise((resolve) => setTimeout(resolve, 600));
    await api.command("run");
    return waitFor("break mode", async () => (await api.state()).debugMode === "break", { budgetMs: 5000 }).catch(() => false);
  }, { budgetMs: 30000 });

  const all = "wb,a,b,c,n";
  await expect("the break shows every local", (list) => top(list) === all);
  await toggle("wb", 0);
  await expect("expanding the first local keeps every local after it",
    (list) => top(list) === "wb+,a,b,c,n" && list.filter((r) => r.depth === 1).length > 100);
  await toggle("c", 0);
  await expect("a row recovered from below it expands, with its members nested",
    (list) => top(list) === "wb+,a,b,c+,n"
      && list.filter((r) => r.depth === 1).slice(-2).map((r) => `${r.expression}=${r.value}`).join(",") === "Item 1=\"first\",Item 2=\"second\"");
  await api.command("stepInto");
  await api.command("stepInto");
  await expect("a step reaches a value below the expanded blocks", (list) => value(list, "a") === "8");
  await toggle("c", 0);
  await expect("the recovered row collapses", (list) => top(list) === "wb+,a,b,c,n");
  await toggle("Worksheets", 1);
  await expect("a member far out of view expands, and the members after it keep their depth", (list) => {
    const at = list.findIndex((r) => r.depth === 1 && r.expression === "Worksheets");
    const after = list.slice(at + 1).filter((r) => r.depth < 2);
    return list[at]?.expanded && list[at + 1]?.depth === 2
      && after[0]?.depth === 1 && top(list) === "wb+,a,b,c,n";
  });
  await toggle("Worksheets", 1);
  await expect("it collapses", (list) => !list.some((r) => r.depth === 2));
  await toggle("wb", 0);
  await expect("the first local collapses", (list) => top(list) === all && list.length === 6);
  await toggle("c", 0);
  await expect("the Collection alone nests its members, the row after it stays",
    (list) => top(list) === "wb,a,b,c+,n" && list.filter((r) => r.depth === 1).length === 2);
  await toggle("wb", 0);
  await expect("the first local expands again after a collapse, the Collection still open below",
    (list) => top(list) === "wb+,a,b,c+,n" && list.filter((r) => r.depth === 1).length > 100);
  await toggle("wb", 0);
  await expect("it collapses again", (list) => top(list) === "wb,a,b,c+,n");
  await toggle("c", 0);
  await expect("everything is collapsed", (list) => top(list) === all && list.length === 6);
  check("the break held throughout", (await api.state()).debugMode === "break");
} finally {
  await api.command("reset").catch(() => {});
  await scratch.dispose();
}

done();
