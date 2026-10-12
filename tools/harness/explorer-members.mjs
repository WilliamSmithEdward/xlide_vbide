/*
 * The explorer's members of an unfolded module (#92): Enums and Types listed beside the
 * procedures, the three orders the header's Sort picks, and a property's Get, Let and Set kept
 * together under every one. Read off the drawn tree, the way a developer reads it.
 *
 *   node tools\harness\explorer-members.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();
const NAME = `MembersWitness${process.pid % 10000}`;
const CRLF = "\r\n";

const rows = async () => await api.ask(`[...document.querySelectorAll('.tree-proc')]
  .filter((row) => row.dataset.procModule === ${JSON.stringify(NAME)})
  .map((row) => row.textContent.trim())`) ?? [];

async function sortBy(sort) {
  const asked = await api.act("explorerSort", { sort });
  const deadline = Date.now() + 5000;
  while (Date.now() < deadline && (await api.ui()).explorer?.sort !== sort) { await wait(150); }
  await wait(400);
  return asked;
}

console.log("the explorer's module members, read off the tree\n");
const before = (await api.ui()).explorer?.sort ?? "code";

try {
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => {});
  await api.component("add", { kind: "module", name: NAME, project: project.projectId });
  await api.writeModule(NAME, [
    "Option Explicit", "",
    "Private mTotal As Long", "",
    "Public Property Let Total(ByVal value As Long)", "    mTotal = value", "End Property", "",
    "Public Enum Colour", "    Red", "    Green", "End Enum", "",
    "Public Sub run()", "End Sub", "",
    "Private Type Point", "    X As Double", "    Y As Double", "End Type", "",
    "Public Function Area() As Double", "End Function", "",
    "Public Property Get Total() As Long", "    Total = mTotal", "End Property", "",
    "Public Sub Apply()", "End Sub",
  ].join(CRLF), project.projectId);
  await wait(1200);
  await api.pane("open", { module: NAME, project: project.projectId });
  await wait(800);
  // A TOGGLE, and opening the pane may already have unfolded it: the tree follows the editor.
  await wait(800);
  if ((await api.ui()).explorer?.unfolded?.module !== NAME) {
    await api.act("unfoldModule", { module: NAME });
  }

  await sortBy("code");
  let drawn = await rows();
  for (let tries = 0; drawn.length < 7 && tries < 20; tries += 1) { await wait(250); drawn = await rows(); }
  check("as in code: Enums and Types are listed, a property's legs together where the first stood",
    JSON.stringify(drawn) === JSON.stringify([
      "Property Get Total", "Property Let Total", "Enum Colour", "Sub run", "Type Point", "Function Area", "Sub Apply"]),
    JSON.stringify(drawn));

  await sortBy("name");
  drawn = await rows();
  check("by name is alphabetical",
    JSON.stringify(drawn) === JSON.stringify([
      "Sub Apply", "Function Area", "Enum Colour", "Type Point", "Sub run", "Property Get Total", "Property Let Total"]),
    JSON.stringify(drawn));

  await sortBy("kind");
  drawn = await rows();
  check("by kind: Enums, Types, Subs, Functions, Properties",
    JSON.stringify(drawn) === JSON.stringify([
      "Enum Colour", "Type Point", "Sub Apply", "Sub run", "Function Area", "Property Get Total", "Property Let Total"]),
    JSON.stringify(drawn));

  const settings = await api.settings();
  check("the order is a setting the host holds", settings.explorerSort === "kind", JSON.stringify(settings.explorerSort));

  const glyphs = await api.ask(`[...document.querySelectorAll('.tree-proc')]
    .filter((row) => row.dataset.procModule === ${JSON.stringify(NAME)})
    .map((row) => row.querySelector('.codicon')?.className.replace('codicon ', ''))`);
  check("each kind wears its own glyph",
    JSON.stringify(glyphs) === JSON.stringify([
      "codicon-symbol-enum", "codicon-symbol-structure", "codicon-symbol-method", "codicon-symbol-method",
      "codicon-symbol-method", "codicon-symbol-property", "codicon-symbol-property"]),
    JSON.stringify(glyphs));
} finally {
  await sortBy(before).catch(() => {});
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => {});
}

process.exit(done());
