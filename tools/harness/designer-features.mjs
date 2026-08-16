/*
 * The designer route, held to the plan that built the form: every control present with its
 * type, container and geometry, then the three mutations round-tripped, then everything gone.
 *
 * The suite builds its own form through the api against whatever session is live - the same
 * calls tools\New-FormFixture.ps1 makes - so it needs no fixture of its own and costs the
 * gate no extra Excel launch. It removes the component at the end, whatever happened, because
 * the suites after it (write-rollback above all) make claims about what the fixture contains.
 *
 * This is docs/userform-designer.md's M1 landing: the read route is the instrument every
 * later designer claim is measured against, and this suite is what keeps the instrument
 * honest.
 */
import { open, reporter, waitFor } from "./xlide-api.mjs";
import { FORM_CONTROLS, FORM_MODULE as PLANNED_FORM, FORM_PROPERTIES, buildForm } from "./form-plan.mjs";

const api = await open();
// projectId is the workbook display name, which is what the routes resolve; `project` is the
// VBA project's own name and with one workbook open the two are easy to confuse.
const { projectId: project, components } = await api.project();
const { check, done } = reporter();

// The suite runs against whatever fixture is live, so the not-a-form refusal is checked
// against a module that fixture actually has rather than a name from someone else's.
const plainModule = components.find((component) => component.kind !== "form")?.name;

const near = (a, b) => a !== null && a !== undefined && Math.abs(a - b) < 0.01;

console.log(`the designer route, against ${project}\n`);

// The fixture's own from-disk form, read FIRST - before any designer tab or markup request
// touches it. A 2026-08-13 session answered "has no designer to read" exactly here while
// the designer-tab path read the same form seconds later; the split has not reproduced
// since (probed first-touch on fresh sessions, 2026-08-14), so this row stands where the
// reproduction would begin and names it if it ever returns. On a session whose fixture form
// is already gone (a crashed run's aftermath) the condition is not exercisable - the gate's
// fresh session is where it always is.
const fixtureFormStands = components.some((component) => component.name === PLANNED_FORM);
if (fixtureFormStands) {
  const fromDisk = await api.designer(PLANNED_FORM, project).catch((why) => ({ failed: why.message }));
  check("the fixture's own from-disk form answers the route first-touch",
    fromDisk.failed === undefined && fromDisk.form?.caption === "Quarter Entry",
    JSON.stringify(fromDisk.failed ?? fromDisk.form ?? null));
} else {
  check("the fixture's own from-disk form answers the route first-touch", true,
    `not exercisable: ${PLANNED_FORM} is not in this session`);
}

/*
 * The suite's form gets a name the session will actually take. A form NAME can be refused for
 * the rest of a session once it has been used - added and removed, or even refused once - and
 * a workbook that LOADED holding the planned name (FormFixture itself) burns it the moment
 * anything removes that form. So: never remove a form this suite did not create, and walk to
 * the first name that adds cleanly. The PLANNED name itself is never in the walk: it is the
 * fixture's own form, load-bearing for the first-touch row above - on a session where some
 * mishap already removed it, walking onto that name built a suite form there and the cleanup
 * then removed the fixture's name for good (measured 2026-08-14: a crashed run left the next
 * one to do exactly that).
 */
let form = `${PLANNED_FORM}2`;
{
  let built = false;
  for (let attempt = 0; attempt < 4 && !built; attempt++) {
    form = `${PLANNED_FORM}${attempt + 2}`;
    try {
      await buildForm(api, project, form);
      built = true;
    } catch (why) {
      if (!/refused as a name/.test(String(why?.message))) {
        throw why;
      }
    }
  }

  if (!built) {
    throw new Error(`no usable form name: the session has burned ${PLANNED_FORM}2 through ${form}`);
  }
}

/** What the cleanup's workbook write answered, checked after the try closes. */
let written = null;

/*
 * Every raw-DOM row below scopes itself to THIS form's view.
 *
 * A live session can hold more than one designer tab - another form's, or one a developer left
 * open by hand - and a bare `querySelector(".designer-canvas-scroll")` takes whichever stands
 * first in the DOM, which is a coin toss about which document a keystroke edits. The view root
 * carries `data-module`, so aiming is one attribute; every gesture row below is aimed.
 */
const inView = (selector) => `document.querySelector('.designer-view[data-module="${form}"] ${selector}')`;
const inViewAll = (selector) => `[...document.querySelectorAll('.designer-view[data-module="${form}"] ${selector}')]`;

/**
 * A real keydown on THIS view's canvas: the gestures the keyboard owns - nudge, resize, undo,
 * delete - all arrive this way, through the listener a developer's own key would reach.
 *
 * It THROWS when the canvas is not mounted, and that is half the point of the helper: the page
 * keeps one designer view attached at a time, so a canvas whose tab is not showing is not in
 * the DOM at all, and a press into nothing would leave the row waiting out its budget for a
 * change nobody asked for. An instrument that quietly does nothing is worse than a missing one.
 */
const press = async (key, extra = "") => {
  const answer = await api.ask(
    `(() => { const el = ${inView(".designer-canvas-scroll")}; if (!el) return "no canvas"; `
    + `el.dispatchEvent(new KeyboardEvent("keydown", { key: ${JSON.stringify(key)}, ${extra} `
    + `bubbles: true, cancelable: true })); return "sent"; })()`);
  if (answer !== "sent") {
    throw new Error(`${key} went nowhere: ${form}'s canvas is not mounted (${answer})`);
  }

  return answer;
};

try {
  // ---- built through the model, as the fixture generator builds it ----

  const design = await api.designer(form, project);

  const caption = FORM_PROPERTIES.find((p) => p.property === "Caption")?.value;
  check("the form carries its caption", design.form.caption === caption, design.form.caption);
  check("the form carries its size, through the component's own properties",
    near(design.form.width, 360) && near(design.form.height, 320),
    `${design.form.width}x${design.form.height}`);
  check("the canvas area is the inside pair, smaller than the outside by the chrome",
    design.form.insideWidth > 0 && design.form.insideWidth < design.form.width,
    `inside ${design.form.insideWidth} vs ${design.form.width}`);

  // ---- every planned control, by identity, kind, container and geometry ----

  const byName = new Map(design.controls.map((control) => [control.name, control]));
  for (const wanted of FORM_CONTROLS) {
    const found = byName.get(wanted.name);
    check(`${wanted.name} is on the form`, found !== undefined);
    if (!found) {
      continue;
    }

    check(`  ...as a ${wanted.type}`, found.type === wanted.type, found.type);
    check(`  ...in ${wanted.parent ?? "the form"}`,
      wanted.parent ? found.parent === wanted.parent : found.parent === form,
      `parent=${found.parent}`);
    check("  ...where it was put",
      near(found.left, wanted.left) && near(found.top, wanted.top)
      && near(found.width, wanted.width) && near(found.height, wanted.height),
      `${found.left},${found.top} ${found.width}x${found.height}`);

    const askedCaption = wanted.set?.Caption;
    if (askedCaption) {
      check(`  ...captioned "${askedCaption}"`, found.caption === askedCaption, found.caption);
    }
  }

  const pages = design.controls.filter((control) => control.type === "Page");
  check("the MultiPage's pages are rows of the tree, parented to it",
    pages.length >= 2 && pages.every((page) => page.parent === "Wizard"),
    pages.map((page) => `${page.name} in ${page.parent}`).join(", "));

  check("every control carries a font or is a shape that has none",
    design.controls.every((control) => control.font === null || typeof control.font.name === "string"));

  // ---- the mutations, round-tripped ----

  const probe = await api.designerEdit("add", {
    module: form, project, type: "commandButton", name: "ProbeButton",
    left: 10, top: 280, width: 60, height: 20,
  });
  check("add answers the name and the kind it made", probe.ok === true && probe.name === "ProbeButton"
    && probe.type === "CommandButton", JSON.stringify(probe));

  const captioned = await api.designerEdit("set", {
    module: form, project, name: "ProbeButton", property: "Caption", value: "Probe", as: "text",
  });
  check("set answers what the property reads back", /Caption is Probe/.test(captioned.detail ?? ""),
    captioned.detail);

  const sized = await api.designerEdit("set", {
    module: form, project, name: "ProbeButton", property: "Font.Size", value: "12",
  });
  check("a dotted set reaches the font", /Font\.Size is 12/.test(sized.detail ?? ""), sized.detail);

  const afterAdd = await api.designer(form, project);
  const probeRow = afterAdd.controls.find((control) => control.name === "ProbeButton");
  check("the added control reads back where it was put, with its caption and font",
    probeRow !== undefined && near(probeRow.left, 10) && probeRow.caption === "Probe"
    && near(probeRow.font?.size, 12),
    JSON.stringify(probeRow ?? null));

  const removed = await api.designerEdit("remove", { module: form, project, name: "ProbeButton" });
  const afterRemove = await api.designer(form, project);
  check("remove takes it back out",
    removed.ok === true && !afterRemove.controls.some((control) => control.name === "ProbeButton"));

  const nested = await api.designerEdit("remove", { module: form, project, name: "PickAir" });
  const afterNested = await api.designer(form, project);
  check("a control inside a Frame is removed through its own container",
    nested.ok === true && !afterNested.controls.some((control) => control.name === "PickAir")
    && afterNested.controls.some((control) => control.name === "PickGround"));

  // ---- the controls resolve in the code-behind (xlide_vscode#17, fed by the seed) ----

  const { FORM_CODE } = await import("./form-plan.mjs");
  await api.writeModule(form, FORM_CODE, project);
  await api.pane("open", { module: form, project });

  // FORM_CODE line 4 is `    RegionPick.AddItem "North"`; column 16 sits after the dot.
  await waitFor("the seed to carry the controls into completion", async () =>
    ((await api.act("completions", { line: 4, column: 16 })).data ?? []).length > 0,
    { budgetMs: 20000 });

  const comboMembers = ((await api.act("completions", { line: 4, column: 16 })).data ?? [])
    .map((item) => item.label ?? item.insertText);
  check("a control receiver offers its own type's members",
    comboMembers.includes("AddItem") && comboMembers.length > 30,
    `${comboMembers.length} member(s): ${comboMembers.slice(0, 5).join(", ")}`);

  // Line 6 is `    Taxable.Value = True`; column 13 sits after the dot. A CheckBox has no
  // AddItem, which is what proves the controls are typed individually rather than as one set.
  const checkMembers = ((await api.act("completions", { line: 6, column: 13 })).data ?? [])
    .map((item) => item.label ?? item.insertText);
  check("and the types are per control, not one shared set",
    checkMembers.includes("Value") && !checkMembers.includes("AddItem"),
    `${checkMembers.length} member(s)`);

  // Line 11 is `        NameBox.SetFocus`; column 17 sits after the dot. SetFocus is in no
  // per-control dump - it lives on MSForms.Control, the base every placed control extends,
  // and the analyzer merges that base into each control class (xlide_vscode#20's side find).
  // Before the merge this list had no SetFocus, no Visible, no Left.
  const textMembers = ((await api.act("completions", { line: 11, column: 17 })).data ?? [])
    .map((item) => item.label ?? item.insertText);
  check("a control offers its Control base class: SetFocus and Visible on a TextBox",
    textMembers.includes("SetFocus") && textMembers.includes("Visible")
    && !textMembers.includes("AddItem"),
    `${textMembers.length} member(s)`);

  await waitFor("the code-behind to analyse clean", async () =>
    ((await api.problems(form)).findings ?? [])
      .every((finding) => finding.code !== "undeclared-variable"), { budgetMs: 20000 });
  const codeBehind = (await api.problems(form)).findings ?? [];
  check("no control reference reads as undeclared - the analyzer's own resolution, no filter",
    codeBehind.every((finding) => finding.code !== "undeclared-variable"),
    JSON.stringify(codeBehind.map((finding) => finding.code)));

  const controlHover = await api.act("hover", { word: "RegionPick" });
  check("a control answers hover", controlHover.did === true, controlHover.detail);
  const memberHover = await api.act("hover", { word: "AddItem" });
  check("and so does its member", memberHover.did === true, memberHover.detail);
  const baseHover = await api.act("hover", { word: "SetFocus" });
  check("and so does a member of the Control base", baseHover.did === true, baseHover.detail);

  // FORM_CODE line 14 is `    Me.Hide`; column 8 sits after the dot. Me composes the
  // controls, the form surface (Show rides the analyzer's own VBA-wrapper table, not
  // MSForms), and the module's code - the three sources xlide_vscode#18's canary names.
  const meItems = ((await api.act("completions", { line: 14, column: 8 })).data ?? [])
    .map((item) => item.label ?? item.insertText);
  check("Me. offers the controls and the form surface",
    meItems.includes("RegionPick") && meItems.includes("Show"), `${meItems.length} member(s)`);

  // ---- the form as text ----

  const markup = await api.designerMarkup(form, project);
  check("the markup opens with the form line",
    markup.startsWith(`Form ${form} "`), markup.split("\n")[0]);
  check("a nested control prints inside its container",
    /\r?\n    Frame Options "Freight" at 12,112 size 92x66\r?\n        OptionButton PickGround/.test(markup),
    markup.slice(0, 400));
  check("a page prints under its MultiPage, a control under the page",
    /\r?\n    MultiPage Wizard[^\r\n]*\r?\n        Page Page1 "Page1"\r?\n            CheckBox Agree/.test(markup));

  const idempotent = await api.applyMarkup(form, markup, project);
  check("applying the form's own markup adds and removes nothing",
    idempotent.ok === true && idempotent.added.length === 0 && idempotent.removed.length === 0,
    JSON.stringify({ added: idempotent.added, removed: idempotent.removed, set: idempotent.set }));

  const withButton = `${markup.trimEnd()}\r\n    CommandButton MarkupBtn "Go" at 8,282 size 60x20\r\n`;
  const applied = await api.applyMarkup(form, withButton, project);
  check("a line added to the document adds the control",
    applied.ok === true && applied.added.includes("MarkupBtn"), JSON.stringify(applied));

  const afterMarkup = await api.designer(form, project);
  const btn = afterMarkup.controls.find((control) => control.name === "MarkupBtn");
  check("and it reads back placed, captioned, the kind the line said",
    btn !== undefined && btn.type === "CommandButton" && btn.caption === "Go"
    && near(btn.left, 8) && near(btn.top, 282),
    JSON.stringify(btn ?? null));

  const backAgain = await api.applyMarkup(form, markup, project);
  check("re-applying the original document removes it again",
    backAgain.ok === true && backAgain.removed.includes("MarkupBtn"), JSON.stringify(backAgain));

  const refusedMarkup = await api.applyMarkup(form, "Form X\n  Label L at banana\n", project)
    .catch((why) => why.message);
  check("a document that does not parse applies nothing, naming the line",
    /did not parse/.test(String(refusedMarkup)) && /line 2/.test(String(refusedMarkup)),
    String(refusedMarkup));
  const untouched = await api.designer(form, project);
  check("and the form is untouched by it",
    untouched.controls.length === afterMarkup.controls.length - 1,
    `${untouched.controls.length} controls vs ${afterMarkup.controls.length} with the probe button`);

  // ---- the designer tab: the form worn as an editor tab, markup beside the visual ----

  await api.pane("open", { module: form, face: "design" });
  const withDesigner = await waitFor("the designer tab to stand in the strip", async () =>
    ((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .find((tab) => tab.module === form && tab.face === "design"), { budgetMs: 15000 });
  check("the designer tab stands, active, labelled with its face",
    withDesigner.active === true && /\[Design\]/.test(withDesigner.label),
    JSON.stringify(withDesigner));

  await waitFor("the panel to target the designer tab's form", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });
  check("and opening it targets the form in the Properties panel, the native designer's own selection", true);

  // The divider's chevrons are real SVG paths now - one rendered as an empty button in a
  // state the headless frame could not reproduce (the owner, 2026-08-15) - and a glyph
  // that is an element can be counted where a border-trick pseudo could not be.
  const chevrons = await api.ask(`${inViewAll(".designer-collapse svg")}.length`);
  check("both divider chevrons carry their glyph", Number(chevrons) === 2, `${chevrons} glyph(s)`);

  // ---- Run with the designer tab active launches the form - the editor's own F5 ----

  // The command POSTS the editor's own action and answers promptly (measured 2026-08-14:
  // the action reads its aim ~30ms after Execute returns); the form stands a beat later.
  // The userform verb stays off the host thread so it answers whatever that thread is in.
  // The caption on the REAL window doubles as the proof that a form-level property set
  // through the api reaches the surface the runtime paints - the component Properties bag,
  // not the designer dispatch's dead copy (the owner's side-by-side found the difference).
  const running = api.command("run");
  const launched = await waitFor("the form to stand running", async () =>
    ((await api.userforms()).forms ?? []).find((title) => title.includes("Quarter Entry")),
  { budgetMs: 20000 });
  check("Run with the designer tab active launches the form, wearing the form's own caption",
    launched !== undefined, JSON.stringify(launched ?? null));
  const ranAnswer = await running;
  check("and the Run command answered that it executed", ranAnswer.ran === true, ranAnswer.detail);

  const closedForm = await api.userforms("close", "Quarter Entry");
  check("the running form closes the way its X would", closedForm.ran === true, closedForm.detail);

  await waitFor("design mode to return", async () =>
    (await api.state()).debugMode === "design", { budgetMs: 15000 });
  await waitFor("the native designer window to go back down", async () => {
    const afterRun = await api.state();
    const rows = (await api.windows()).windows ?? [];
    // Type 1 is the designer window, type 10 the Toolbox - the palette the owner has now
    // reported three times, total suppression pinned here by the editor's own window list.
    return afterRun.paletteVisible === false
      && !rows.some((w) => ((w.type === 1 && w.caption.includes(form)) || w.type === 10) && w.visible);
  }, { budgetMs: 15000 });
  check("and the native designer window went back down, no toolbox standing", true);

  // The round trip a click makes: away to the code tab, back to the designer tab. A tab
  // rebuilt from the strip WITHOUT its face is the code identity, so clicking back onto
  // the designer activated the code pane instead and the designer tab never took the slot
  // back (the developer, 2026-08-13).
  const awayToCode = await api.act("activate", { module: form });
  check("activating the code tab takes the active slot from the designer",
    awayToCode.did === true, awayToCode.detail);
  const backToDesign = await api.act("activate", { module: form, face: "design" });
  check("and activating the designer tab takes it back",
    backToDesign.did === true, backToDesign.detail);

  // The tab's own apply - the path Ctrl+S takes, through the view's document and the host
  // round trip - proven against the live form, not just against the route it shares a
  // service with.
  const tabMarkup = await waitFor("the tab's document to hold the form's markup", async () => {
    const read = await api.act("designerMarkup", { module: form });
    return read.did === true && String(read.data).startsWith(`Form ${form}`) ? read : null;
  }, { budgetMs: 15000 });
  check("the tab's document holds the form's markup", true, tabMarkup.detail);

  const tabApplied = await api.act("designerApply", {
    module: form,
    markup: `${String(tabMarkup.data).trimEnd()}\r\n    CommandButton TabBtn "Go" at 8,282 size 60x20\r\n`,
  });
  check("applying an edited document from the tab lands on the form",
    tabApplied.did === true && (tabApplied.data?.added ?? []).includes("TabBtn"),
    JSON.stringify({ detail: tabApplied.detail, data: tabApplied.data }));

  const afterTabApply = await api.designer(form, project);
  check("and the form truly carries what the tab applied",
    afterTabApply.controls.some((control) => control.name === "TabBtn"
      && control.type === "CommandButton" && control.caption === "Go"));

  const putBack = await api.act("designerApply", { module: form, markup: String(tabMarkup.data) });
  check("re-applying the original document from the tab removes it again",
    putBack.did === true && (putBack.data?.removed ?? []).includes("TabBtn"),
    JSON.stringify(putBack.data));

  const refusedTab = await api.act("designerApply", { module: form, markup: "Form X\n  Label L at banana\n" });
  check("a document that does not parse is refused at the tab, naming the line",
    refusedTab.did === false && /line 2/.test(refusedTab.detail ?? ""), refusedTab.detail);
  const untouchedByTab = await api.designer(form, project);
  check("and the form is untouched by the refusal",
    !untouchedByTab.controls.some((control) => control.name === "TabBtn"));

  // ---- form properties in the markup, linked to the native panel's source, and LIVE ----

  // The refusal above deliberately left the BAD text in the tab's document, dirty - and a
  // dirty document is protected from every push. Restore it first, which also pins that an
  // ok apply reopens the document to pushes.
  await api.act("designerApply", { module: form, markup: String(tabMarkup.data) });

  // A change made OUTSIDE the tab - the api's set, same source the native Properties window
  // writes - reaches the OPEN tab's document without anyone re-activating it: the routes
  // re-project the tab after every mutation.
  await api.designerEdit("set", { module: form, project, property: "BackColor", value: "12632256" });
  await waitFor("the open tab's document to grow the BackColor line", async () => {
    const read = await api.act("designerMarkup", { module: form });
    return /BackColor = 12632256/.test(String(read.data ?? ""));
  }, { budgetMs: 15000 });
  check("a form property set through the api appears in the OPEN tab's document, live", true);

  const routeMarkup = await api.designerMarkup(form, project);
  check("and in the projection route's document",
    /\r?\n    BackColor = 12632256\r?\n/.test(routeMarkup));

  // The dialect's rule, held from the other side: a document WITHOUT the line cannot erase
  // the colour - an unspoken property is one an apply never touches.
  await api.act("designerApply", { module: form, markup: String(tabMarkup.data) });
  const kept = await api.designer(form, project);
  check("an apply of a document without the property line leaves the colour standing",
    kept.form.backColor === 12632256, String(kept.form.backColor));

  // A document WITH the line changes it - the write goes where the native panel's would.
  const recoloured = `${String(tabMarkup.data).replace(/\r?\n/, "\r\n    BackColor = 12639424\r\n")}`;
  await api.act("designerApply", { module: form, markup: recoloured });
  const changed = await api.designer(form, project);
  check("an apply of a document with the line writes the native panel's own property",
    changed.form.backColor === 12639424, String(changed.form.backColor));

  await api.designerEdit("set", { module: form, project, property: "BackColor", value: "-2147483633" });
  const defaulted = await api.designerMarkup(form, project);
  check("back at the default, the line leaves the document - defaults stay unspoken",
    !/BackColor/.test(defaulted));

  // ---- the squiggles: Core's tolerant parse, drawn on the document as it is typed ----

  await api.act("designerSetMarkup", {
    module: form,
    markup: `Form ${form} size 100x100\r\n    Label A at banana\r\n    Gadget G at 1,1 size 2x2\r\n`,
  });
  const lint = await waitFor("the squiggles to arrive", async () => {
    const read = (await api.act("designerLint", { module: form })).data ?? [];
    return read.length >= 2 ? read : null;
  }, { budgetMs: 15000 });
  check("a bad document wears its squiggles, each at its line, warnings apart from errors",
    lint.some((f) => f.line === 2 && f.severity === "error")
    && lint.some((f) => f.line === 3 && f.severity === "warning" && /ProgId/.test(f.message)),
    JSON.stringify(lint));

  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });
  await waitFor("the squiggles to clear on the canonical text", async () =>
    ((await api.act("designerLint", { module: form })).data ?? []).length === 0, { budgetMs: 15000 });
  check("and the canonical document wears none", true);

  // ---- the canvas follows the document: the draft previews, the form untouched ----

  const draftMarkup = `${String(tabMarkup.data).trimEnd()}\r\n    CommandButton DraftBtn "Soon" at 8,282 size 60x20\r\n`;
  await api.act("designerSetMarkup", { module: form, markup: draftMarkup });
  const draftCanvas = await waitFor("the canvas to preview the draft", async () => {
    const read = (await api.act("designerCanvas", { module: form })).data;
    return read?.draft === true && read.controls.some((c) => c.name === "DraftBtn") ? read : null;
  }, { budgetMs: 15000 });
  check("a control typed into the document appears on the canvas as a DRAFT, before any apply",
    true, `${draftCanvas.controls.length} control(s), draft`);
  check("and the form itself is untouched by the preview",
    !(await api.designer(form, project)).controls.some((c) => c.name === "DraftBtn"));

  // A half-typed line must not blank the picture: broken text keeps the last good draft.
  await api.act("designerSetMarkup", { module: form, markup: `${draftMarkup}    Label Broken at banana\r\n` });
  await waitFor("the squiggles to arrive on the broken text", async () =>
    ((await api.act("designerLint", { module: form })).data ?? []).length > 0, { budgetMs: 15000 });
  const heldCanvas = (await api.act("designerCanvas", { module: form })).data;
  check("a document that stops parsing keeps the last good picture",
    heldCanvas.draft === true && heldCanvas.controls.some((c) => c.name === "DraftBtn"));

  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });
  await waitFor("the applied projection to return", async () => {
    const read = (await api.act("designerCanvas", { module: form })).data;
    return read?.draft === false && !read.controls.some((c) => c.name === "DraftBtn");
  }, { budgetMs: 15000 });
  check("and a document back at canonical puts the form's own picture back", true);

  // ---- the canvas scrolls, and Ctrl+S is the product's save ----

  // A draft tall enough to guarantee overflow, whatever the frame's height; the wheel is
  // synthesised, which is exactly what the view's own wheel handler exists to serve - a
  // dispatched WheelEvent never triggers native scrolling, so this path had no driver.
  await api.act("designerSetMarkup", {
    module: form,
    markup: String(tabMarkup.data).replace(/size \d+x[\d.]+/, "size 360x900"),
  });
  await waitFor("the tall draft to render", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.draft) === true, { budgetMs: 15000 });
  const wheeled = await api.ask(`(() => { const el = ${inView(".designer-canvas-scroll")}; el.scrollTop = 0; el.dispatchEvent(new WheelEvent("wheel", { deltaY: 240, bubbles: true, cancelable: true })); return el.scrollTop; })()`);
  check("a wheel over the canvas scrolls the form", Number(wheeled) > 0, `scrollTop ${wheeled}`);

  // A draft picture is said twice and no more: the tab's own unsaved dot, and the outline
  // around the form. The banner that said it a third time in words was retired 2026-08-15
  // (the owner: the dot on the tab is fine), so this row holds the two cues that remain -
  // and the absence of the third, because the strip it occupied is the point.
  const draftCues = await api.ask(`(() => { const scroll = ${inView(".designer-canvas-scroll")}; const face = ${inView(".dc-form")}; return [scroll?.classList.contains("draft") ? "outlined" : "no outline", ${inView(".designer-draft-note")} ? "banner stands" : "no banner", face ? "form drawn" : "no form"].join("|"); })()`);
  check("the draft picture wears the form outline and no banner",
    draftCues === "outlined|no banner|form drawn", String(draftCues));
  const draftTab = ((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
    .find((tab) => tab.module === form && tab.face === "design");
  check("and the tab carries the unsaved dot, which is where the state is said plainly",
    draftTab?.dirty === true, JSON.stringify(draftTab ?? null));

  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });

  // The apply the tab's Ctrl+S makes is followed by the HOST'S save - command 3, File
  // Save - because Ctrl+S means "save the workbook" everywhere else in the product and
  // the designer must not quietly mean less.
  const saveMark = (await api.log({ max: 20000 })).next;
  await api.act("designerApply", { module: form, markup: String(tabMarkup.data) });
  const savedLine = await api.waitForLog("command: 3 executed", { since: saveMark, timeout: 15000 })
    .catch(() => null);
  check("an apply through the tab is followed by the host's own save - Ctrl+S saves here too",
    savedLine !== null && (savedLine.lines ?? []).length > 0,
    savedLine ? savedLine.lines[0] : "no File Save in the log after the apply");

  // The REAL Ctrl+S is a HOST accelerator - the page never sees the key - so the host's
  // Save, finding the designer tab active, asks the tab to apply and the page calls back
  // for the raw save. `command save` drives the exact entry the keystroke takes.
  await api.act("activate", { module: form, face: "design" });
  await api.act("designerSetMarkup", { module: form, markup: draftMarkup });
  const acceleratorMark = (await api.log({ max: 20000 })).next;
  await api.command("save");
  await waitFor("the draft to land on the form through the host's save", async () =>
    (await api.designer(form, project)).controls.some((c) => c.name === "DraftBtn"), { budgetMs: 15000 });
  check("the host's own Ctrl+S with a designer tab active applies the draft first", true);
  const acceleratorSave = await api.waitForLog("command: 3 executed", { since: acceleratorMark, timeout: 15000 })
    .catch(() => null);
  check("and the raw save follows it",
    acceleratorSave !== null && (acceleratorSave.lines ?? []).length > 0,
    acceleratorSave ? acceleratorSave.lines[0] : "no File Save in the log");
  await api.act("designerApply", { module: form, markup: String(tabMarkup.data) });

  // ---- M3 opens: selection on the canvas, and the double-click's event stub ----

  await api.act("designerSelect", { module: form, control: "RegionPick" });
  const selectedCanvas = (await api.act("designerCanvas", { module: form })).data;
  check("clicking a control selects it on the canvas", selectedCanvas.selected === "RegionPick",
    JSON.stringify(selectedCanvas.selected));
  const markupLine = Number(selectedCanvas.markupLine ?? 0);
  const markupLines = String((await api.act("designerMarkup", { module: form })).data).split(/\r?\n/);
  check("and the markup caret lands on the selected control's line",
    markupLine > 1 && /RegionPick/.test(markupLines[markupLine - 1] ?? ""), `line ${markupLine}`);

  await api.act("designerSelect", { module: form });
  const formSelected = (await api.act("designerCanvas", { module: form })).data;
  check("clicking the form's own ground selects the form, caret to the Form line",
    formSelected.selected === "" && Number(formSelected.markupLine) === 1,
    JSON.stringify({ selected: formSelected.selected, line: formSelected.markupLine }));

  // The double-click: the control's DEFAULT event handler, written into the code-behind
  // when absent (RegionPick is a ComboBox, so Change), and only navigated to when it
  // stands - the native designer's gesture, duplication and all... minus the duplication.
  await api.act("designerEventStub", { module: form, control: "RegionPick" });
  await waitFor("the default event stub to land in the code-behind", async () =>
    /Private Sub RegionPick_Change\(\)/.test((await api.readModule(form, project)).text ?? ""), { budgetMs: 15000 });
  check("double-click writes the control's default event stub", true);

  await api.act("designerEventStub", { module: form, control: "RegionPick" });
  await new Promise((resolve) => setTimeout(resolve, 1500));
  const stubs = ((await api.readModule(form, project)).text ?? "").match(/Private Sub RegionPick_Change\(\)/g) ?? [];
  check("and a second double-click navigates rather than duplicating", stubs.length === 1,
    `${stubs.length} stub(s)`);

  // ---- the M4 bridgehead: selection flows into the Properties panel as control rows ----

  await api.act("activate", { module: form, face: "design" });

  // The REAL click's path: hit-testing at the control's own centre must answer the
  // control. designerSelect bypasses hit-testing, which is how an invisible full-canvas
  // overlay (the notice, display-clobbered past its hidden attribute) ate every real
  // click while 157 checks stayed green.
  const hitAnswer = await api.ask(`(() => { const el = ${inViewAll(".dc")}.find(e => e.dataset.control === "RegionPick"); if (!el) return "no element"; const r = el.getBoundingClientRect(); const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2); return el === hit || el.contains(hit) ? "the control" : (hit?.className || "nothing"); })()`);
  check("a real click's hit-test lands on the control, not an invisible overlay",
    hitAnswer === "the control", String(hitAnswer));

  // The HAND across the whole form face (the owner's call, 2026-08-15): every inch of it
  // responds to a press, which is what a hand means. The gesture cursors live elsewhere - the
  // move cursor while a drag runs, each handle's own resize cursor.
  const ergonomics = await api.ask(`(() => { const el = ${inViewAll(".dc")}.find(e => e.dataset.control === "RegionPick"); const face = ${inView(".dc-form")}; return getComputedStyle(el).cursor + "|" + getComputedStyle(face).cursor + "|" + (el.title || "no title"); })()`);
  check("hovering the canvas shows the hand, and the name-and-kind tooltip stands",
    /^pointer\|pointer\|RegionPick \(ComboBox\)$/.test(String(ergonomics)), String(ergonomics));

  await api.act("designerSelect", { module: form, control: "RegionPick" });

  // Selected, it offers the move - and says so with the cursor, which an unselected control
  // beside it does not (the owner's rule, 2026-08-15).
  const selectedCursor = await api.ask(`(() => { const of = (n) => { const el = ${inViewAll(".dc")}.find(e => e.dataset.control === n); return el ? getComputedStyle(el).cursor : "missing"; }; return of("RegionPick") + "|" + of("NameBox"); })()`);
  check("the SELECTED control wears the move cursor, its neighbours keep the hand",
    selectedCursor === "move|pointer", String(selectedCursor));

  const controlPanel = await waitFor("the panel to show the selected control", async () => {
    const shown = (await api.ui()).properties;
    return shown?.component === "RegionPick" ? shown : null;
  }, { budgetMs: 15000 });
  check("selecting a control targets the panel at it, kind and all",
    controlPanel.kind === "ComboBox", controlPanel.kind);
  const leftRow = controlPanel.rows.find((row) => row.name === "Left");
  check("and its geometry reads through the rows", Number(leftRow?.value) === 84, leftRow?.value);

  const panelMove = await api.act("editProperty", { name: "Left", value: "90" });
  check("a control row's edit answers through the model", panelMove.did === true, panelMove.detail);
  await waitFor("the form to carry the panel's move", async () =>
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 90) < 0.01, { budgetMs: 15000 });
  check("the form truly moved", true);
  await waitFor("and the open tab's document to follow it", async () =>
    /RegionPick at 90,38/.test(String((await api.act("designerMarkup", { module: form })).data)),
  { budgetMs: 15000 });
  check("the markup follows the panel's move - liveness end to end", true);

  await api.act("editProperty", { name: "Left", value: "84" });
  await waitFor("and back where the plan puts it", async () =>
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 84) < 0.01, { budgetMs: 15000 });

  await api.act("designerSelect", { module: form });
  await waitFor("the form's ground to return the panel to the form", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });
  check("selecting the form returns the panel to the component", true);

  // ---- the defaults inventory: what a kind holds before anyone touches it ----

  // Measured from a bare instance of the coclass, so this is MSForms' own answer rather than a
  // table of ours. It is what the projection will compare against to print only what changed.
  const buttonDefaults = await api.controlDefaults("commandButton");
  const byProperty = new Map(buttonDefaults.properties.map((row) => [row.name, row.value]));
  check("a kind's untouched values are measured, not written down",
    buttonDefaults.count > 10 && byProperty.get("BackColor") === "-2147483633"
    && byProperty.get("Enabled") === "True",
    `${buttonDefaults.count} properties, BackColor=${byProperty.get("BackColor")}`);

  // The inventory's own honesty: a FONT on a bare control is not the font it would wear on a
  // form, which inherits. So the projection compares fonts against the form, not against this.
  check("and it says so where a bare control cannot know - a font is inherited, not defaulted",
    byProperty.get("FontName") !== undefined && byProperty.get("FontName") !== "Tahoma",
    `bare says ${byProperty.get("FontName")}, a form's control wears Tahoma`);

  const foreign = await api.controlDefaults("gizmo");
  check("a kind with no coclass of ours answers nothing rather than a guess",
    foreign.count === 0, JSON.stringify(foreign));

  // ---- the panel speaks the developer's language: members, not their ints ----

  const formPanel = await waitFor("the panel to show the form", async () => {
    const shown = (await api.ui()).properties;
    return shown?.component === form ? shown : null;
  }, { budgetMs: 15000 });

  const cycleRow = formPanel.rows.find((row) => row.name === "Cycle");
  check("an enum-valued property shows its member name, and offers the members",
    cycleRow?.value === "fmCycleAllForms" && (cycleRow?.options ?? []).includes("fmCycleCurrentForm"),
    JSON.stringify(cycleRow));

  const colourRow = formPanel.rows.find((row) => row.name === "BackColor");
  check("a colour shows the hex the VBE spells, not a signed integer",
    /^&H[0-9A-F]{8}&$/.test(String(colourRow?.value)), String(colourRow?.value));

  // The honest limit, pinned: names come from the TYPE LIBRARY, never from the property's own
  // name. StartUpPosition is a plain Integer in MSForms however much it looks like an enum, so
  // it keeps its number rather than being given a vocabulary this product invented.
  const plainRow = formPanel.rows.find((row) => row.name === "StartUpPosition");
  check("a property the library does not name keeps its number - nothing is guessed",
    /^-?\d+$/.test(String(plainRow?.value)) && !plainRow?.options, JSON.stringify(plainRow));

  const wroteName = await api.act("editProperty", { name: "Cycle", value: "fmCycleCurrentForm" });
  check("a member NAME writes through to the model", wroteName.did === true, wroteName.detail);

  // ...and the number still writes, which is what keeps the row a text field rather than a
  // list: the developer types either, and the panel answers in the language it reads back.
  const wroteNumber = await api.act("editProperty", { name: "Cycle", value: "0" });
  check("the raw number writes too, and reads back as the member's name",
    wroteNumber.did === true && /fmCycleAllForms/.test(wroteNumber.detail ?? ""), wroteNumber.detail);

  const wroteHex = await api.act("editProperty", { name: "BackColor", value: "&H00C0FFC0&" });
  check("a colour written as hex lands as the model's own number",
    wroteHex.did === true, wroteHex.detail);
  await waitFor("the form to carry the colour", async () =>
    ((await api.designer(form, project)).form?.backColor ?? 0) === 12648384, { budgetMs: 15000 });
  check("12648384, which is what &H00C0FFC0& means", true);
  await api.act("editProperty", { name: "BackColor", value: "&H8000000F&" });

  // A CONTROL's rows read the same way, from that control's own library.
  await api.act("designerSelect", { module: form, control: "RegionPick" });
  const controlNamed = await waitFor("the panel to show the control's named values", async () => {
    const shown = (await api.ui()).properties;
    return shown?.component === "RegionPick"
      && shown.rows.some((row) => row.name === "TextAlign") ? shown : null;
  }, { budgetMs: 15000 });
  const alignRow = controlNamed.rows.find((row) => row.name === "TextAlign");
  check("a control's enum rows are named too, from its own type library",
    alignRow?.value === "fmTextAlignLeft" && (alignRow?.options ?? []).includes("fmTextAlignCenter"),
    JSON.stringify(alignRow));

  // The controls the rows wear: a caret that is ALWAYS there on a named-value row (an
  // affordance that appears on hover is one you have to know about first), a swatch on a
  // colour, and the platform's own arrow on a True/False - three rows, one visual language.
  const rowControls = await api.ask('(() => { const of = (name) => { const row = [...document.querySelectorAll(".prop-row")].find(r => r.querySelector(".prop-name")?.textContent === name); if (!row) return name + "=missing"; if (row.querySelector(".prop-caret")) return name + "=caret"; if (row.querySelector(".prop-swatch")) return name + "=swatch"; if (row.querySelector("select.prop-value")) return name + "=select"; return name + "=plain"; }; return [of("TextAlign"), of("BackColor"), of("Enabled"), of("Left")].join(" "); })()');
  check("a named-value row wears a caret, a colour a swatch, a flag the platform's arrow",
    String(rowControls) === "TextAlign=caret BackColor=swatch Enabled=select Left=plain",
    String(rowControls));

  // The caret shows EVERY member, not the one the field already holds: a datalist filtered by
  // the field's own text offered exactly one choice, which is no dropdown at all (the owner,
  // 2026-08-15: "I only see one option for most").
  const listed = await api.ask('(() => { const row = [...document.querySelectorAll(".prop-row")].find(r => r.querySelector(".prop-name")?.textContent === "TextAlign"); const caret = row?.querySelector(".prop-caret"); if (!caret) return "no caret"; caret.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, cancelable: true })); const options = [...row.querySelectorAll(".prop-option")].map(o => o.textContent); document.activeElement?.blur(); return options.join(","); })()');
  check("and its caret opens the whole list, not just what the field holds",
    String(listed) === "fmTextAlignLeft,fmTextAlignCenter,fmTextAlignRight", String(listed));

  const alignWrite = await api.act("editProperty", { name: "TextAlign", value: "fmTextAlignCenter" });
  check("and a control's enum takes a member name on the way in",
    alignWrite.did === true && /fmTextAlignCenter/.test(alignWrite.detail ?? ""), alignWrite.detail);
  await api.act("editProperty", { name: "TextAlign", value: "fmTextAlignLeft" });

  // ---- M5 opens: the canvas moves controls, and the DOCUMENT is the transaction log ----

  // The real pointer sequence - press on what the hit test answers, past the threshold, drop -
  // so the row proves the gesture, not the arithmetic behind it.
  const dragged = await api.act("designerDrag", { module: form, control: "RegionPick", dx: 24, dy: 12 });
  check("a pointer drag on the canvas moves a control", dragged.did === true, dragged.detail);
  await waitFor("the drop to reach the document", async () =>
    /RegionPick at 108,50/.test(String((await api.act("designerMarkup", { module: form })).data)),
  { budgetMs: 15000 });
  check("the drop rewrites the control's line - at 84,38 becomes at 108,50", true);

  const dragCanvas = await waitFor("the canvas to preview the drop", async () => {
    const read = (await api.act("designerCanvas", { module: form })).data;
    return read?.draft === true ? read : null;
  }, { budgetMs: 15000 });
  const dropped = dragCanvas.controls.find((c) => c.name === "RegionPick");
  check("the picture is the DRAFT, and the control stands where it was dropped",
    dragCanvas.dirty === true && Math.abs(Number(dropped?.left) - 108) < 0.51,
    JSON.stringify({ dirty: dragCanvas.dirty, left: dropped?.left }));
  check("and the FORM is untouched until the save - the drag writes the document, not the form",
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 84) < 0.01);

  // Ctrl+S is the road to the form, the same one a hand-typed edit takes.
  await api.command("save");
  await waitFor("the drag to land on the form", async () =>
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 108) < 0.01, { budgetMs: 15000 });
  check("Ctrl+S carries the drag to the form", true);

  // One Ctrl+Z takes the whole drag back, from the CANVAS half: proof the drop is a single
  // edit rather than a stream of them, and that the canvas reaches the document's undo.
  await press("z", "ctrlKey: true,");
  await waitFor("the document to come back", async () =>
    /RegionPick at 84,38/.test(String((await api.act("designerMarkup", { module: form })).data)),
  { budgetMs: 15000 });
  check("one Ctrl+Z on the canvas undoes the whole drag - the document is the transaction log", true);
  await api.command("save");
  await waitFor("the undo to reach the form too", async () =>
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 84) < 0.01, { budgetMs: 15000 });
  check("and the undo travels the same road back to the form", true);

  // The arrow keys: one point a press, through the same commit the drop takes.
  await api.act("designerSelect", { module: form, control: "RegionPick" });
  await press("ArrowRight");
  await waitFor("the nudge to reach the document", async () =>
    /RegionPick at 85,38/.test(String((await api.act("designerMarkup", { module: form })).data)),
  { budgetMs: 15000 });
  check("an arrow nudges the selection by a single point", true);

  // A second nudge is a SECOND undo step. The stack stop before each move is what keeps one
  // gesture from swallowing the one before it, and one Ctrl+Z gives back exactly one.
  await press("ArrowRight");
  await waitFor("the second nudge", async () =>
    /RegionPick at 86,38/.test(String((await api.act("designerMarkup", { module: form })).data)),
  { budgetMs: 15000 });
  await press("z", "ctrlKey: true,");
  await waitFor("one gesture back", async () =>
    /RegionPick at 85,38/.test(String((await api.act("designerMarkup", { module: form })).data)),
  { budgetMs: 15000 });
  check("each move is its own undo step - one Ctrl+Z gives back one gesture, not the pair", true);

  // A drag with nowhere to go cannot lose the control behind an edge: it stops at the corner.
  await api.act("designerDrag", { module: form, control: "RegionPick", dx: -400, dy: -400 });
  await waitFor("the clamped drop", async () =>
    /RegionPick at 0,0/.test(String((await api.act("designerMarkup", { module: form })).data)),
  { budgetMs: 15000 });
  check("a drag past the top-left stops at the parent's corner rather than going negative", true);

  // ---- and the handles resize: the same commit, the same undo, the size clause too ----

  // From canonical, so every number below is exact rather than inherited from the rows above.
  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });
  await waitFor("the document back at canonical before the resize rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });

  const tabText = async () => String((await api.act("designerMarkup", { module: form })).data);
  const grew = await api.act("designerResize",
    { module: form, control: "RegionPick", edge: "se", dx: 24, dy: 12 });
  check("a pull on the south-east handle resizes a control", grew.did === true, grew.detail);
  await waitFor("the size clause to follow the handle", async () =>
    /RegionPick at 84,38 size 144x32/.test(await tabText()), { budgetMs: 15000 });
  check("120x20 becomes 144x32, and the position is untouched", true);

  await api.act("designerResize", { module: form, control: "RegionPick", edge: "nw", dx: 12, dy: 12 });
  await waitFor("the north-west pull", async () =>
    /RegionPick at 96,50 size 132x20/.test(await tabText()), { budgetMs: 15000 });
  check("a north-west pull moves the origin and the extent together", true);

  const drawn = (await api.act("designerCanvas", { module: form })).data.controls
    .find((c) => c.name === "RegionPick");
  check("and the canvas draws the box it just wrote",
    Math.abs(Number(drawn?.width) - 132) < 0.51 && Math.abs(Number(drawn?.height) - 20) < 0.51,
    JSON.stringify(drawn));

  await api.act("designerResize", { module: form, control: "RegionPick", edge: "se", dx: -500, dy: -500 });
  await waitFor("the floored box", async () =>
    /RegionPick at 96,50 size 4x4/.test(await tabText()), { budgetMs: 15000 });
  check("a pull past the far edge stops at the floor size rather than inverting the box", true);

  // Shift+arrow is the keyboard's resize, the native designer's own pairing with a bare arrow.
  await press("ArrowRight", "shiftKey: true,");
  await waitFor("the keyboard resize", async () =>
    /RegionPick at 96,50 size 5x4/.test(await tabText()), { budgetMs: 15000 });
  check("Shift+arrow resizes by a point where a bare arrow moves", true);

  // The FORM's own frame resizes too - and its line takes a size and never a position.
  const formGrew = await api.act("designerResize", { module: form, edge: "se", dx: 20, dy: 12 });
  check("the form's own frame resizes by its handles", formGrew.did === true, formGrew.detail);
  await waitFor("the form's line to carry the new size", async () =>
    new RegExp(`^Form ${form} .* size 380x332$`, "m").test(await tabText()), { budgetMs: 15000 });
  check("360x320 becomes 380x332 on the Form line, with no position added", true);

  // The promise the handles now keep: each wears the cursor of the pull it makes, and the
  // pointer can actually land on it. Asked of a CONTROL near the top of the form rather than
  // of the form's own frame, whose far corner sits below the canvas viewport on a short tab -
  // where elementFromPoint honestly answers the markup editor underneath (2026-08-15, the
  // first run after the toolbox strip took its 30px).
  // A handle that overlaps a NEIGHBOUR must still be what a press reaches. The handles hang
  // 3px outside the box they dress and their overlay is only a sibling of it, so a control
  // later in the document painted straight over them: NameBox's right edge meets the ScrollBar,
  // and that edge could not be grabbed at all until the overlay was lifted above the controls
  // (found by the perf walk, 2026-08-15). NameBox is chosen for exactly that adjacency.
  await api.act("designerSelect", { module: form, control: "NameBox" });
  const handleCursor = await api.ask(`(() => { const h = ${inView(".dc-handle-se")}; if (!h) return "no handle"; const r = h.getBoundingClientRect(); const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2); return getComputedStyle(h).cursor + "|" + (hit === h ? "reachable" : "covered by " + (hit?.className || "nothing")); })()`);
  check("a handle wears its own resize cursor and the pointer can reach it",
    handleCursor === "nwse-resize|reachable", String(handleCursor));

  // ...and the resize that handle promises actually runs, over the neighbour and all.
  const overNeighbour = await api.act("designerResize",
    { module: form, control: "NameBox", edge: "se", dx: 6, dy: 0 });
  check("and the pull it promises runs, even where a neighbour overlaps the handle",
    overNeighbour.did === true, overNeighbour.detail);

  // ---- and Delete takes a control out, children and all ----

  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });
  await waitFor("the document back at canonical before the delete rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });

  // A Frame, on purpose: its child is indented UNDER it, so a delete that only took the header
  // line would leave an orphaned OptionButton the parser has nowhere to put.
  //
  // The Frame holds ONE child here, not the two the plan draws: the nested-remove row above
  // took PickAir off the form for good, and every projection since has been without it. A
  // first draft of these rows asserted all three names came back and spent an hour looking for
  // a fault in undo - the document simply never held the third. Read what the state IS.
  const deleted = await api.act("designerDelete", { module: form, control: "Options" });
  check("Delete takes the selected control out of the document", deleted.did === true, deleted.detail);
  const afterDelete = await tabText();
  check("and its whole block goes with it - properties and children included",
    !/Frame Options/.test(afterDelete) && !/PickGround/.test(afterDelete)
    && /ToggleButton HoldToggle/.test(afterDelete),
    afterDelete.split(/\r?\n/).length + " line(s) left");
  check("the FORM still holds it until the save",
    (await api.designer(form, project)).controls.some((c) => c.name === "PickGround"));
  check("and the selection lands back on the form, where the panel follows it",
    ((await api.act("designerCanvas", { module: form })).data?.selected) === "");

  await press("z", "ctrlKey: true,");
  const blockBack = await waitFor("the block to come back", async () => {
    const text = await tabText();
    return /Frame Options/.test(text) && /PickGround/.test(text) ? "restored" : null;
  }, { budgetMs: 8000 }).catch(async () => {
    // A bare timeout says nothing about WHY, and chasing this one without the state cost an
    // hour. The document's own stack answers first: an undo that found nothing to give back
    // reads undoable=false, and a document already at canonical reads dirty=false.
    const snap = (await api.act("designerCanvas", { module: form })).data;
    return `not back: undoable=${snap?.undoable} dirty=${snap?.dirty} `
      + `selected=${JSON.stringify(snap?.selected)} drawn=${snap?.controls?.length}`;
  });
  check("one Ctrl+Z brings the whole block back - a mistaken Delete costs one keystroke",
    blockBack === "restored", String(blockBack));

  // ---- the xlide toolbox: a kind dragged out of the palette and dropped on the form ----

  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });
  await waitFor("the document back at canonical before the toolbox rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });

  // The palette must offer exactly what the apply can add. The route's own refusal names its
  // kinds, so the two lists are compared rather than both being written down twice.
  const kindsOffered = await api.ask(`${inViewAll(".designer-tool")}.map(b => b.dataset.kind.toLowerCase()).sort().join(",")`);
  const kindsRefusal = String((await api.designerEdit("add", { module: form, project, type: "gizmo" })
    .catch((why) => why.message)));
  const kindsTaken = (kindsRefusal.match(/one of ([^:]+)$/)?.[1] ?? "")
    .split(",").map((one) => one.trim().toLowerCase()).filter(Boolean).sort().join(",");
  check("the toolbox offers exactly the kinds the apply can add",
    String(kindsOffered) === kindsTaken, `palette=${kindsOffered} route=${kindsTaken}`);

  // An empty patch of the form's ground: right of the Frame, below the Image and the two
  // spinner shapes, so the drop lands on the FORM rather than on something already there.
  const fromPalette = await api.act("designerToolbox",
    { module: form, kind: "CommandButton", left: 230, top: 120 });
  check("a kind dragged out of the toolbox lands on the form",
    fromPalette.did === true, fromPalette.detail);
  const withDrop = await tabText();
  check("the drop writes a whole line - named, placed and sized like the native toolbox's",
    /^ {4}CommandButton CommandButton1 at 230,120 size 72x24$/m.test(withDrop),
    withDrop.split(/\r?\n/).find((line) => /CommandButton1/.test(line)) ?? "no line");
  check("and the new control is SELECTED, the way one dropped from the native palette is",
    ((await api.act("designerCanvas", { module: form })).data?.selected) === "CommandButton1");
  check("the FORM does not have it until the save",
    !(await api.designer(form, project)).controls.some((c) => c.name === "CommandButton1"));

  // Dropped INSIDE a Frame, the container under the pointer wins: the line nests under it and
  // its position is the frame's own, not the form's.
  const nestedDrop = await api.act("designerToolbox",
    { module: form, kind: "Label", left: 40, top: 150 });
  check("a drop over a Frame goes INTO the frame", nestedDrop.did === true, nestedDrop.detail);
  const withNested = await tabText();
  check("and its line is indented under it, placed in the frame's own coordinates",
    /^ {8}Label Label1 at \d+,\d+ size 66x16$/m.test(withNested),
    withNested.split(/\r?\n/).find((line) => /Label1/.test(line)) ?? "no line");

  // The whole loop: Ctrl+S puts both on the FORM, through the same add the api route makes.
  await api.command("save");
  const landed = await waitFor("the drops to reach the form", async () => {
    const walk = await api.designer(form, project);
    const button = walk.controls.find((c) => c.name === "CommandButton1");
    const label = walk.controls.find((c) => c.name === "Label1");
    return button && label ? { button, label } : null;
  }, { budgetMs: 20000 });
  check("Ctrl+S adds them to the form where the drop put them",
    near(landed.button.left, 230) && near(landed.button.top, 120)
    && landed.button.type === "CommandButton", JSON.stringify(landed.button));
  check("and the one dropped in the Frame is parented to it",
    landed.label.parent === "Options", `parent=${landed.label.parent}`);

  // Off the form entirely is not a drop: the ghost goes home and the document is untouched.
  const missed = await api.act("designerToolbox", { module: form, kind: "Label", left: -400, top: -400 });
  check("a drop outside the form adds nothing rather than a control at the origin",
    missed.did === false, missed.detail);

  await api.act("designerDelete", { module: form, control: "CommandButton1" });
  await api.act("designerDelete", { module: form, control: "Label1" });
  await api.command("save");
  await waitFor("the toolbox's controls to leave the form again", async () => {
    const walk = await api.designer(form, project);
    return !walk.controls.some((c) => c.name === "CommandButton1" || c.name === "Label1");
  }, { budgetMs: 20000 });
  check("and Delete takes them off the form the same way, leaving the plan's own controls", true);

  // Undo all the way to the FLOOR of the stack. The document's own arrival must not be a step
  // on it: as an edit it was, and one Ctrl+Z in a tab whose only edit was the projection
  // landing blanked the document while the canvas kept showing the form - found by hand and by
  // the owner in the same minute, with ten green drag rows standing (2026-08-15).
  for (let i = 0; i < 12; i++) { await press("z", "ctrlKey: true,"); }
  const floor = String((await api.act("designerMarkup", { module: form })).data);
  check("undone to the floor of the stack, the document is still the form's text - never blank",
    /^Form /.test(floor) && /CommandButton OkButton/.test(floor), `${floor.length} char(s)`);

  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });
  const restored = await waitFor("the document back at canonical", async () => {
    const read = (await api.act("designerCanvas", { module: form })).data;
    return read?.dirty === false ? read : null;
  }, { budgetMs: 15000 });
  check("and the unapplied moves let go, leaving the form as the plan draws it",
    restored.draft === false);
  await api.act("designerSelect", { module: form });

  // The stub gesture pulled the CODE tab active; the rename rows below need the face back.
  await api.act("activate", { module: form, face: "design" });

  // ---- a rename carries the designer tab: re-keyed in place, never a corpse ----

  const designTabs = async () =>
    ((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .filter((tab) => tab.face === "design");

  // The product rename, with the designer tab ACTIVE - the hardest case, because the shown
  // module's name changes underneath and the publish must not read that as a native move
  // that takes the active slot away from the tab.
  const renamedForm = `${form}R`;
  await api.act("activate", { module: form, face: "design" });
  const away = await api.act("renameModule", { module: form, newName: renamedForm });
  check("the product rename answers", away.did === true, away.detail);

  const followed = await waitFor("the designer tab to follow the rename", async () =>
    (await designTabs()).find((tab) => tab.module === renamedForm), { budgetMs: 15000 });
  check("the designer tab stands under the NEW name, still active",
    followed.active === true && followed.label === `${renamedForm} [Design]`,
    JSON.stringify(followed));
  check("and no corpse stands under the old one",
    !(await designTabs()).some((tab) => tab.module === form));

  const followedMarkup = await waitFor("the followed tab's document to answer under the new name",
    async () => {
      const read = await api.act("designerMarkup", { module: renamedForm });
      return read.did === true && String(read.data).startsWith(`Form ${renamedForm}`) ? read : null;
    }, { budgetMs: 15000 });
  check("and its document answers under the new name, first line and all", true, followedMarkup.detail);

  // There and back - which also proves a renamed-away name is reusable in the session,
  // unlike a removed one's (probed 2026-08-13: a rename does not burn the name).
  const back = await api.act("renameModule", { module: renamedForm, newName: form });
  check("the rename back answers, reusing the renamed-away name", back.did === true, back.detail);
  await waitFor("the designer tab to follow back", async () =>
    (await designTabs()).find((tab) => tab.module === form), { budgetMs: 15000 });

  // The Properties panel's own (Name) row renames through the same adoption, so the tab
  // follows there too - it used to close instead. The panel targets whatever was last
  // opened or selected, so the code tab aims it first.
  await api.act("activate", { module: form });
  await waitFor("the panel to target the form", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });
  const panelAway = await api.act("editProperty", { name: "(Name)", value: renamedForm });
  check("the panel's (Name) edit answers", panelAway.did === true, panelAway.detail);
  await waitFor("the designer tab to follow the panel's rename", async () =>
    (await designTabs()).find((tab) => tab.module === renamedForm), { budgetMs: 15000 });
  check("the panel's rename carries the tab the same way, no corpse either",
    !(await designTabs()).some((tab) => tab.module === form));

  await waitFor("the panel to follow its own rename", async () =>
    (await api.ui()).properties?.component === renamedForm, { budgetMs: 15000 });
  await api.act("editProperty", { name: "(Name)", value: form });
  await waitFor("the tab to come back once more", async () =>
    (await designTabs()).find((tab) => tab.module === form), { budgetMs: 15000 });
  check("and back once more through the panel", true);

  await api.pane("close", { module: form, face: "design" });
  await waitFor("the designer tab to leave the strip", async () =>
    !((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === form && tab.face === "design"), { budgetMs: 15000 });
  check("and its close takes only the designer face, not the code tab",
    ((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === form && !tab.face));

  const notFormTab = await api.pane("open", { module: plainModule, face: "design" })
    .catch((why) => why.message);
  check("a module that is not a form is refused a designer tab",
    /no designer tab/.test(String(notFormTab)), String(notFormTab));

  // ---- refusals are answers ----

  const missing = await api.designer("NoSuchForm", project).catch((why) => why.message);
  check("a form that is not there is refused by name", /no component named NoSuchForm/.test(String(missing)),
    String(missing));

  const notAForm = await api.designer(plainModule, project).catch((why) => why.message);
  check(`a module that is not a form (${plainModule}) is refused as one`,
    /not a UserForm/.test(String(notAForm)), String(notAForm));

  const badKind = await api.designerEdit("add", { module: form, project, type: "gizmo" })
    .catch((why) => why.message);
  check("an unknown control kind is refused with the list", /not a control kind/.test(String(badKind)),
    String(badKind));
} finally {
  await api.component("remove", { name: form, project }).catch(() => {});
  // And the FILE too, not only the session. This suite saves the workbook on purpose - the
  // Ctrl+S rows exist to prove the designer's save reaches it - so every run left its temporary
  // form written into the fixture, where the next run's other suites found it. The gate's own
  // discard probe takes the project's FIRST module as its subject, and on 2026-08-15 that was
  // a leftover UserForm for the first time; it timed out waiting for the form's findings to
  // clear. The save after the removal is what makes "as found" true on disk.
  written = await api.command("save").catch((why) => ({ ran: false, detail: why.message }));
}

const swept = await api.project();
check("the form is gone again, leaving the fixture as found",
  !swept.components.some((component) => component.name === form));
check("and the workbook was written after the removal, so the FILE is as found too",
  written?.ran === true, JSON.stringify(written ?? null));

done();
