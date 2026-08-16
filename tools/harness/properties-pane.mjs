/*
 * The Properties panel, which writes real component state and could not be reached at all.
 *
 * WHY THIS EXISTS. It was the only user-visible surface in the product with no api presence in
 * EITHER direction: no route, no act, no client method, no snapshot field, no suite, and no doc
 * row. The panel is a view over the object model's own property bag, so its "(Name)" row renames
 * a component and every other row writes through `Properties.Item(name)` into the host - the
 * shim's only property-write path - and none of that had ever been driven by anything but a hand.
 *
 * WHAT IT GUARDS. Three things the panel has to get right and nothing was checking. A write must
 * reach the object model, which is asserted by reading the component back through a different
 * route than the one that wrote it. A REFUSED write must not leave the grid showing a value the
 * host does not have, which is the failure mode that makes a properties grid worse than no grid:
 * the developer reads the row and believes it. And the panel must not go on describing a
 * component that has been removed, which is the same failure a step further - not one wrong row
 * but a whole pane of them, under a name the tree no longer lists.
 *
 *   node tools\harness\properties-pane.mjs
 */
import { open, reporter, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const name = `Props${process.pid}`;

const { check, done } = reporter();

const panel = async () => (await api.ui()).properties;

let made = false;
let currentName = name;

try {
  await api.component("add", { kind: "module", name, project: project.projectId });
  made = true;
  await api.writeModule(name, "Option Explicit\r\n", project.projectId);

  // Opening a module selects it, which is what puts it in the panel - the same thing the tree
  // does on a click.
  await api.pane("open", { module: name, project: project.projectId });
  await waitFor("the panel to show the module", async () =>
    (await panel()).component?.toLowerCase() === name.toLowerCase());

  const shown = await panel();
  check("the panel names the component it is showing", shown.component === name,
    JSON.stringify(shown).slice(0, 200));

  check("and says what kind of thing it is", (shown.kind ?? "").length > 0,
    `kind is ${JSON.stringify(shown.kind)}`);

  // A standard module offers exactly one property, and it is the one that renames it.
  const nameRow = shown.rows.find((row) => row.name === "(Name)");
  check("a standard module offers its (Name) row", nameRow !== undefined,
    `rows: ${shown.rows.map((r) => r.name).join(", ")}`);
  check("and the row is writable", nameRow?.writable === true, JSON.stringify(nameRow));
  check("and it holds the module's current name", nameRow?.value === name,
    `the row says ${JSON.stringify(nameRow?.value)}`);

  // ---- a write that must reach the object model ----

  const renamed = `${name}Renamed`;
  const said = await api.act("editProperty", { name: "(Name)", value: renamed });
  check("the panel accepts an edit to (Name)", said.did, said.detail);

  await waitFor("the rename to reach the object model", async () =>
    (await api.readModule(renamed, project.projectId).catch(() => ({ text: null }))).text !== null);
  currentName = renamed;

  // Read back through a DIFFERENT route than the one that wrote it: the panel agreeing with
  // itself proves nothing about the host.
  const readBack = await api.readModule(renamed, project.projectId);
  check("the component answers to the new name through the module route",
    (readBack.text ?? "").includes("Option Explicit"), JSON.stringify(readBack).slice(0, 160));

  await waitFor("the panel to follow the rename", async () =>
    (await panel()).component?.toLowerCase() === renamed.toLowerCase());
  check("and the panel follows it", (await panel()).component === renamed);

  // ---- a write the host will refuse ----

  // A name with a space is not a legal VBA identifier, so the editor refuses it. What matters is
  // that the grid does not go on showing it afterwards.
  const refused = await api.act("editProperty", { name: "(Name)", value: "not a legal name" });
  check("a write the host refuses is reported as refused, not as done",
    refused.did === false && /refused/i.test(refused.detail ?? ""),
    `answered ${JSON.stringify(refused)}. The panel posts an edit and does not wait, so an act `
    + "that returns as soon as it has posted reports every refusal as a success - which is what "
    + "this said until the check was written");

  await waitFor("the panel to settle after the refused write", async () =>
    (await panel()).rows.some((row) => row.name === "(Name)"));

  const afterRefusal = await panel();
  const stillNamed = afterRefusal.rows.find((row) => row.name === "(Name)")?.value;
  check("a refused write does not leave the grid showing a value the host rejected",
    stillNamed === renamed,
    `the grid says ${JSON.stringify(stillNamed)} but the component is ${renamed}. A properties `
    + "grid the developer cannot believe is worse than no grid at all");

  check("and the component still answers to the name it had",
    ((await api.readModule(renamed, project.projectId)).text ?? "").includes("Option Explicit"));

  // ---- a property that does not exist is refused, not invented ----

  const nonsense = await api.act("editProperty", { name: "NoSuchProperty", value: "1" });
  check("a property the component does not have is refused by name",
    !nonsense.did && /no writable property/i.test(nonsense.detail ?? ""),
    JSON.stringify(nonsense));

  // ---- a document component speaks the language too ----
  //
  // A worksheet's values live in EXCEL's type library, not the editor's, and a worksheet has no
  // designer to point at one. So these rows read `Visible -1` and `EnableSelection 0` while every
  // form in the product had been reading its enums by name since the morning. The panel goes to
  // the host object for them now - the sheet itself, matched by code name.

  const sheet = ((await api.project(project.projectId)).components ?? [])
    .find((one) => one.type === 100 && !/^ThisWorkbook$/i.test(one.name ?? ""));
  check("the project offers a worksheet document to read", sheet !== undefined,
    "no document component other than ThisWorkbook, so the enum rows below cannot be checked");

  await api.pane("open", { module: sheet.name, project: project.projectId });
  await waitFor("the panel to show the worksheet", async () =>
    (await panel()).component?.toLowerCase() === sheet.name.toLowerCase());

  const sheetRows = await panel();
  const visible = sheetRows.rows.find((row) => row.name === "Visible");
  check("a worksheet's Visible row is spelled the way the language spells it",
    visible?.value === "xlSheetVisible",
    `the row says ${JSON.stringify(visible?.value)}; -1 is what the object model stores and `
    + "xlSheetVisible is what the developer writes");

  check("and the row offers the members of its enum",
    (visible?.options ?? []).includes("xlSheetVeryHidden"),
    `options: ${JSON.stringify(visible?.options)}`);

  // A write by NAME must reach the model as the number. EnableSelection rather than Visible,
  // because a workbook refuses to hide its last visible sheet and the fixture may have one.
  const chosen = await api.act("editProperty", { name: "EnableSelection", value: "xlUnlockedCells" });
  check("an enum written by name is accepted", chosen.did, chosen.detail);

  await waitFor("the panel to show the written value", async () =>
    (await panel()).rows.find((row) => row.name === "EnableSelection")?.value !== "xlNoRestrictions");

  check("and the row comes back as the member, not as its number",
    (await panel()).rows.find((row) => row.name === "EnableSelection")?.value === "xlUnlockedCells",
    JSON.stringify((await panel()).rows.find((row) => row.name === "EnableSelection")));

  const restored = await api.act("editProperty", { name: "EnableSelection", value: "xlNoRestrictions" });
  check("and the sheet goes back to how it was found", restored.did, restored.detail);

  // Back to the module, which is what the rest of this suite is about.
  await api.pane("open", { module: renamed, project: project.projectId });
  await waitFor("the panel to return to the module", async () =>
    (await panel()).component?.toLowerCase() === renamed.toLowerCase());

  // ---- the panel must not outlive what it describes ----
  //
  // The panel holds the name of whatever it was last aimed at. Nothing told it when that thing
  // was removed, so the tree lost the row and the panel kept the header, and the two disagreed
  // on screen about what the workbook contains (reported 2026-08-15, over a form the tree no
  // longer listed). It is the panel's own reads that are the harm: every row under that header
  // described a component the project does not have.

  check("the panel is aimed at the component that is about to go",
    (await panel()).component === renamed);

  await api.component("remove", { name: renamed, project: project.projectId });
  made = false;

  await waitFor("the component to leave the project", async () =>
    !((await api.project(project.projectId)).components ?? [])
      .some((one) => (one.name ?? "").toLowerCase() === renamed.toLowerCase()));

  // Readiness on a DIFFERENT surface than the one under test: the removal takes the module's tab
  // with it, and that publish is sent before the panel's. Waiting on the panel itself would turn
  // the defect into an aborted run with no check attributed to it.
  await waitFor("the strip to drop the removed module's tab", async () =>
    !((await api.ui()).workspace?.groups ?? []).some((group) =>
      (group.tabs ?? []).some((tab) => (tab.module ?? "").toLowerCase() === renamed.toLowerCase())));

  // Either the shown module or nothing at all, depending on what is left open; what matters is
  // that it is not the component that was removed.
  const orphaned = await panel();
  check("the panel stops naming a component the project no longer holds",
    (orphaned.component ?? "").toLowerCase() !== renamed.toLowerCase(),
    `the panel says ${JSON.stringify(orphaned.component)} and the removed component was ${renamed}`);

  check("and it is not still showing that component's rows",
    !orphaned.rows.some((row) => row.value === renamed),
    `rows: ${JSON.stringify(orphaned.rows).slice(0, 200)}`);
} catch (error) {
  // A throw is a verdict, not an accident. Without this, a precondition failing at the first
  // await - the add refused, the session gone - fell straight to the finally below, which
  // printed "0 passed, 0 failed": a summary the gate reads as green, for a run that checked
  // nothing (2026-08-12). The throw becomes a named failure so the count cannot lie.
  check("the suite ran to its last check", false, error.message);
} finally {
  await api.pane("close", { module: currentName, project: project.projectId, answer: "discard" })
    .catch(() => {});

  if (made) {
    await api.component("remove", { name: currentName, project: project.projectId }).catch(() => {});
    await api.component("remove", { name, project: project.projectId }).catch(() => {});
  }

  process.exitCode = done();
}
