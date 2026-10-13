/*
 * The code window's Object and Procedure lists above the editor (#101), driven the way a click
 * drives them and read back from the module the host holds. A document module lists its own
 * object; picking it writes its default event's handler, as the native list does; picking an
 * event writes that handler in the place the VBE's own CreateEventProc puts it; a handler that
 * stands is gone to, never written twice; and the lists follow the caret.
 *
 * The full comparison against the VBE - every object and event of a sheet, a workbook, a class
 * with WithEvents variables and a form with each toolbox control, handler for handler with
 * CreateEventProc's - was run once by hand on 2026-10-12 and is what the engine's tables were
 * corrected from; this suite holds the behaviour on the gate.
 *
 * Run against the debug fixture:
 *   node tools\harness\code-bar.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

async function until(what, predicate, budgetMs = 20000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(250);
  }
  throw new Error(`timed out waiting for ${what}`);
}

const lists = async () => (await api.act("codeBar")).data;
const sheetText = async () => (await api.readModule("Sheet1", project.projectId)).text ?? "";
const original = await sheetText();

console.log("the code window's Object and Procedure lists\n");

try {
  await api.guard(true);
  await api.pane("open", { module: "Sheet1", project: project.projectId });
  const shown = await until("Sheet1's objects", async () => {
    const now = await lists();
    return now?.objects?.includes("Worksheet") ? now : null;
  });
  check("a sheet's module lists (General) and Worksheet", JSON.stringify(shown.objects) === '["(General)","Worksheet"]',
    JSON.stringify(shown.objects));
  check("the caret above every procedure reads (General) / (Declarations)",
    shown.object === "(General)" && shown.procedure === "(Declarations)", `${shown.object} / ${shown.procedure}`);

  // ---- the object itself: its default event, as the native list writes it ----------------
  const picked = await api.act("codeBar", { object: "Worksheet" });
  check("picking Worksheet lists its events", (picked.data?.procedures ?? []).includes("Change"),
    (picked.data?.procedures ?? []).slice(0, 6).join(", "));
  const withDefault = await until("Worksheet_SelectionChange written",
    async () => ((await sheetText()).includes("Worksheet_SelectionChange(") ? sheetText() : null));
  check("and writes the default handler, Worksheet_SelectionChange, with the type library's parameters",
    withDefault.includes("Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n\r\nEnd Sub"));
  const following = await until("the lists on the new handler", async () => {
    const now = await lists();
    return now.procedure === "SelectionChange" ? now : null;
  }).catch(() => lists());
  check("the lists follow the caret into it", following.object === "Worksheet" && following.procedure === "SelectionChange",
    `${following.object} / ${following.procedure}`);

  // ---- an event: written where CreateEventProc writes it ---------------------------------
  await api.act("codeBar", { object: "Worksheet", procedure: "Activate" });
  const withActivate = await until("Worksheet_Activate written",
    async () => ((await sheetText()).includes("Worksheet_Activate(") ? sheetText() : null));
  const order = ["Worksheet_Activate", "Worksheet_SelectionChange"].map((name) => withActivate.indexOf(name));
  check("an earlier event goes before the object's later handler, as the VBE places it", order[0] >= 0 && order[0] < order[1],
    JSON.stringify(order));

  await api.act("codeBar", { object: "Worksheet", procedure: "Activate" });
  await wait(1500);
  const again = await sheetText();
  check("a handler that stands is gone to, not written again",
    again.split("Sub Worksheet_Activate(").length === 2, `${again.split("Sub Worksheet_Activate(").length - 1} declaration(s)`);

  // ---- (General): the module's own procedures --------------------------------------------
  await api.pane("open", { module: "Helper", project: project.projectId });
  await until("Helper on screen", async () => (await lists())?.procedures?.includes("Describe"));
  await api.act("codeBar", { object: "(General)", procedure: "Describe" });
  await wait(300);
  const general = await lists();
  check("(General) lists the module's procedures and a pick goes there",
    general.object === "(General)" && general.procedure === "Describe", `${general.object} / ${general.procedure}`);

  // ---- the look: inset, sized, popup rows in the theme -----------------------------------
  const look = await api.ask(`(() => {
    const bar = [...document.querySelectorAll(".code-bar")].find((one) => !one.hidden);
    const select = bar && bar.querySelector("select");
    const option = select && select.querySelector("option");
    if (!select || !option) { return null; }
    const box = bar.getBoundingClientRect(), field = select.getBoundingClientRect();
    return { height: field.height, inset: field.left - box.left, optionBackground: getComputedStyle(option).backgroundColor };
  })()`);
  check("the lists are 24px fields inset from the bar's edge, their popup rows coloured",
    look && look.height === 24 && look.inset >= 8 && look.optionBackground !== "rgba(0, 0, 0, 0)", JSON.stringify(look));
} finally {
  await api.writeModule("Sheet1", original, project.projectId).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
}

process.exit(done());
