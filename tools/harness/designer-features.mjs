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
  const wheeled = await api.ask('(() => { const el = document.querySelector(".designer-canvas-scroll"); el.scrollTop = 0; el.dispatchEvent(new WheelEvent("wheel", { deltaY: 240, bubbles: true, cancelable: true })); return el.scrollTop; })()');
  check("a wheel over the canvas scrolls the form", Number(wheeled) > 0, `scrollTop ${wheeled}`);

  // With the canvas scrolled and the draft note up, the note must be what the point under
  // its centre HITS - as a flow sibling of the positioned scroll box it painted UNDER the
  // scrolled form until it became a pinned overlay.
  const bannerHit = await api.ask('(() => { const note = document.querySelector(".designer-draft-note"); if (!note || note.hidden) return "no note standing"; const r = note.getBoundingClientRect(); const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2); return note.contains(hit) ? "note on top" : `covered by ${hit?.className ?? "nothing"}`; })()');
  check("the draft banner stays above the scrolled form", bannerHit === "note on top", String(bannerHit));

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
}

const swept = await api.project();
check("the form is gone again, leaving the fixture as found",
  !swept.components.some((component) => component.name === form));

done();
