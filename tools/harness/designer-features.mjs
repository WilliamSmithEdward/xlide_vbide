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
import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { open, reporter, waitFor } from "./xlide-api.mjs";
import { FORM_CONTROLS, FORM_MODULE as PLANNED_FORM, FORM_PICTURES, FORM_PROPERTIES, buildForm } from "./form-plan.mjs";

const api = await open();
// projectId is the workbook display name, which is what the routes resolve; `project` is the
// VBA project's own name and with one workbook open the two are easy to confuse.
const { projectId: project, components } = await api.project();
const { check, done } = reporter();

// The suite runs against whatever fixture is live, so the not-a-form refusal is checked
// against a module that fixture actually has rather than a name from someone else's.
const plainModule = components.find((component) => component.kind !== "form")?.name;

const near = (a, b) => a !== null && a !== undefined && Math.abs(a - b) < 0.01;

/**
 * A CLASSIC icon file - the kind whose frames are BMPs rather than PNGs, which is the kind OLE's
 * own picture loader reads and hands back as an icon rather than as a bitmap.
 *
 * Written rather than committed because nothing in the repository is one (the product's own .ico
 * is PNG-framed throughout, so it loads through GDI+ and arrives flattened), and because the
 * ICON road in the shim - DrawIconEx onto a DIB section, which is a different piece of code from
 * the bitmap road's GetDIBits - is worth proving on purpose rather than by luck.
 *
 * The layout is the format's own: a directory of one entry, then a BITMAPINFOHEADER whose height
 * is DOUBLED because an icon stacks its image over its mask, the pixels bottom-up, and the mask
 * left at zero, which means "show every pixel".
 */
function classicIcon(side, bgra) {
  const pixels = side * side * 4;
  const mask = side * 4;
  const file = Buffer.alloc(6 + 16 + 40 + pixels + mask);

  file.writeUInt16LE(0, 0);
  file.writeUInt16LE(1, 2);
  file.writeUInt16LE(1, 4);
  file[6] = side;
  file[7] = side;
  file.writeUInt16LE(1, 10);
  file.writeUInt16LE(32, 12);
  file.writeUInt32LE(40 + pixels + mask, 14);
  file.writeUInt32LE(22, 18);

  file.writeUInt32LE(40, 22);
  file.writeInt32LE(side, 26);
  file.writeInt32LE(side * 2, 30);
  file.writeUInt16LE(1, 34);
  file.writeUInt16LE(32, 36);
  file.writeUInt32LE(pixels, 42);

  for (let at = 0; at < pixels; at += 4) {
    file[62 + at] = bgra[0];
    file[63 + at] = bgra[1];
    file[64 + at] = bgra[2];
    file[65 + at] = bgra[3];
  }

  return file;
}

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
 * A control's placed geometry, read out of the DOCUMENT's own attributes.
 *
 * Every gesture row used to match the line whole - `/RegionPick at 85,38/` - which the tagged
 * dialect made impossible twice over: the four numbers are attributes now, and their ORDER is the
 * printer's business rather than anything a row should pin. Reading them by name says what the
 * row means and survives the next thing the printer decides to put first.
 */
const placed = async (name) => {
  const doc = String((await api.act("designerMarkup", { module: form })).data);
  const line = doc.split(/\r?\n/).find((text) => new RegExp(`\\bName="${name}"`).test(text)) ?? "";
  const number = (attribute) => {
    const found = new RegExp(`\\b${attribute}="(-?[\\d.]+)"`).exec(line);
    return found === null ? null : Number(found[1]);
  };

  return {
    left: number("Left"), top: number("Top"),
    width: number("Width"), height: number("Height"),
    line: line.trim(),
  };
};

/** `placed`, as the "84,38 size 144x32" a row wants to print when it fails. */
const placedAt = async (name) => {
  const box = await placed(name);
  return `${box.left},${box.top} size ${box.width}x${box.height}`;
};

/** All four at once, for the resize rows: a gesture that moves an edge changes an origin AND an
 * extent, and a wait that watched only one of them would pass halfway through. */
const sizedAt = async (name, left, top, width, height) => {
  const box = await placed(name);
  return box.left === left && box.top === top && box.width === width && box.height === height;
};

/*
 * The FORM's own text as it stands NOW - what a restore has to write to leave the document
 * clean. The projection captured at the top of the run is not it: rows since have moved
 * controls and saved them, so writing that back leaves a document that differs from the form
 * and a dirty dot nothing can clear (three rows waited fifteen seconds for that, 2026-08-16).
 */
const canonicalNow = async () => await api.designerMarkup(form, project);

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
  // ---- A FORM'S DOCUMENT IS ONLY COMPLETE ONCE THE WORKBOOK HAS BEEN SAVED ----
  //
  // Everything past identity, geometry and the two inherited colours reaches the document by
  // reading what MSForms PERSISTED - the PropMask in the saved workbook, off the file, with
  // nothing exported (FormDesignService.ChangedProperties says why that road was taken). A form
  // built in this session is not in the file yet, so it has no mask, so those properties are
  // absent from its markup however plainly they are set on the control.
  //
  // This is pinned rather than merely worked around, because three language-service rows below
  // aim at `PicturePosition` and `SpecialEffect` on THIS form and had been passing on a stale
  // artifact: a run that dies between building the form and the cleanup leaves its design in the
  // fixture FILE, and the next run's fresh form then finds the dead run's mask under its own
  // name. They passed for a year of runs and failed the first time the file was clean
  // (2026-08-18). A row that only passes when a previous run crashed is worse than no row.
  const beforeSaving = await api.designerMarkup(form, project);
  check("a form built this session spells nothing the saved file has not seen yet",
    !/PicturePosition=/.test(beforeSaving) && !/SpecialEffect=/.test(beforeSaving),
    beforeSaving.split(/\r?\n/).find((line) => /OkButton/.test(line))?.trim());

  // WAITED FOR, not read in the same breath as the save. The baseline is re-read off the
  // WORKBOOK FILE, and the file is not readable the instant the command returns: measured
  // 2026-08-18, a markup print immediately after a save still spells nothing and the same print
  // 1.5s later has the attributes. This row read at once and passed twice before failing twice,
  // which is what a race looks like from the outside. The wait is also the honest statement of
  // the behaviour - a developer's document grows these a beat after Ctrl+S, not during it.
  await api.command("save");
  const afterSaving = await waitFor("the saved mask to reach the document", async () => {
    const text = await api.designerMarkup(form, project);
    return /<CommandButton\b[^>]*\bName="OkButton"[^>]*\bPicturePosition="1"/.test(text)
      && /<Frame\b[^>]*\bName="Options"[^>]*\bSpecialEffect="3"/.test(text) ? text : null;
  }, { budgetMs: 20000 }).catch(() => null);
  check("and the save is what puts them there - the mask is read off the workbook",
    afterSaving !== null,
    (await api.designerMarkup(form, project)).split(/\r?\n/).find((line) => /OkButton/.test(line))?.trim());

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

    // The picture the plan hung on it, described rather than carried: what it is, that it has
    // real pixels, and that the placement property went in beside it.
    if (wanted.picture) {
      // (Bitmap) for both, ICO included - form-plan says why: a PNG-framed icon is not a classic
      // icon and comes back through GDI+ as pixels. The icon road is proved further down, on an
      // icon built to be one.
      check("  ...holding a picture that reads (Bitmap)",
        found.picture?.kind === "(Bitmap)", JSON.stringify(found.picture ?? null));
      check("  ...with real pixels behind it",
        (found.picture?.width ?? 0) > 0 && (found.picture?.height ?? 0) > 0,
        `${found.picture?.width}x${found.picture?.height}`);
      for (const [property, value] of Object.entries(wanted.picture.place ?? {})) {
        const shown = property === "PictureSizeMode" ? found.picture?.sizeMode : found.picture?.position;
        check(`  ...placed with ${property} ${value}`, shown === value, String(shown));
      }
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

  /*
   * Where a receiver's dot is, found in the plan's own text rather than counted by hand.
   *
   * These were four hard-coded line numbers until 2026-08-16, when the fixture grew a
   * Views_Change handler and two of them quietly pointed at blank lines - the rows then failed
   * saying the seed carried no members, which is a true statement about the wrong line.
   */
  const afterDot = (receiver) => {
    const lines = FORM_CODE.split(/\r?\n/);
    const at = lines.findIndex((line) => line.includes(`${receiver}.`));
    return {
      line: at + 1,
      column: (lines[at] ?? "").indexOf(`${receiver}.`) + receiver.length + 2,
    };
  };

  const comboSpot = afterDot("RegionPick");
  await waitFor("the seed to carry the controls into completion", async () =>
    ((await api.act("completions", comboSpot)).data ?? []).length > 0,
    { budgetMs: 20000 });

  const comboMembers = ((await api.act("completions", comboSpot)).data ?? [])
    .map((item) => item.label ?? item.insertText);
  check("a control receiver offers its own type's members",
    comboMembers.includes("AddItem") && comboMembers.length > 30,
    `${comboMembers.length} member(s): ${comboMembers.slice(0, 5).join(", ")}`);

  // A CheckBox has no AddItem, which is what proves the controls are typed individually rather
  // than as one set.
  const checkMembers = ((await api.act("completions", afterDot("Taxable"))).data ?? [])
    .map((item) => item.label ?? item.insertText);
  check("and the types are per control, not one shared set",
    checkMembers.includes("Value") && !checkMembers.includes("AddItem"),
    `${checkMembers.length} member(s)`);

  // SetFocus is in no per-control dump - it lives on MSForms.Control, the base every placed
  // control extends, and the analyzer merges that base into each control class
  // (xlide_vscode#20's side find). Before the merge this list had no SetFocus, no Visible, no
  // Left.
  const textMembers = ((await api.act("completions", afterDot("NameBox"))).data ?? [])
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

  // Me composes the controls, the form surface (Show rides the analyzer's own VBA-wrapper
  // table, not MSForms), and the module's code - the three sources xlide_vscode#18's canary
  // names.
  const meItems = ((await api.act("completions", afterDot("Me"))).data ?? [])
    .map((item) => item.label ?? item.insertText);
  check("Me. offers the controls and the form surface",
    meItems.includes("RegionPick") && meItems.includes("Show"), `${meItems.length} member(s)`);

  // ---- the form as text ----

  const markup = await api.designerMarkup(form, project);
  check("the markup opens with the form line",
    markup.startsWith(`<Form Name="${form}"`), markup.split("\n")[0]);
  check("a nested control prints inside its container",
    /<Frame Name="Options"[^>]*>\r?\n\s+<OptionButton Name="PickGround"/.test(markup)
    && /<\/Frame>/.test(markup),
    markup.slice(0, 400));
  check("a page prints under its MultiPage, a control under the page",
    /<MultiPage Name="Wizard"[^>]*>\r?\n\s+<Page Name="Page1"[^>]*>\r?\n\s+<CheckBox Name="Agree"/
      .test(markup));

  const idempotent = await api.applyMarkup(form, markup, project);
  check("applying the form's own markup adds and removes nothing",
    idempotent.ok === true && idempotent.added.length === 0 && idempotent.removed.length === 0,
    JSON.stringify({ added: idempotent.added, removed: idempotent.removed, set: idempotent.set }));

  // An element goes in before the form's own close, which is what "adding a line" means once the
  // document is tagged: the close is a line rather than the end of the file.
  const addElement = (document, element) =>
    `${document.trimEnd().replace(/<\/Form>\s*$/, "")}    ${element}\r\n</Form>\r\n`;
  const withButton = addElement(markup,
    '<CommandButton Name="MarkupBtn" Caption="Go" Left="8" Top="282" Width="60" Height="20" />');
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

  const refusedMarkup = await api.applyMarkup(
    form, '<Form Name="X">\n    <Label Name="L" Left="banana" Top="1" />\n</Form>\n', project)
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
  // THE DESIGNER ITSELF is what the next rows need, and the tab's document is not evidence
  // about it: the view holds the last projection, so it answers happily while the component
  // is still in the ~400ms after a close where `Designer` is gone rather than empty. Waiting
  // on the document and then applying is waiting on the wrong observable, and it failed here
  // exactly that way (2026-08-16) with ten green rows above it.
  await waitFor("the form's designer to be readable again after the run", async () => {
    try {
      await api.designer(form, project);
      return true;
    } catch {
      return false;
    }
  }, { budgetMs: 15000 });

  const tabMarkup = await waitFor("the tab's document to hold the form's markup", async () => {
    const read = await api.act("designerMarkup", { module: form });
    return read.did === true && String(read.data).startsWith(`<Form Name="${form}"`) ? read : null;
  }, { budgetMs: 15000 });
  check("the tab's document holds the form's markup", true, tabMarkup.detail);

  const tabApplied = await api.act("designerApply", {
    module: form,
    markup: addElement(String(tabMarkup.data),
      '<CommandButton Name="TabBtn" Caption="Go" Left="8" Top="282" Width="60" Height="20" />'),
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

  const refusedTab = await api.act("designerApply", {
    module: form,
    markup: '<Form Name="X">\n    <Label Name="L" Left="banana" Top="1" />\n</Form>\n',
  });
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
    // #c0c0c0 is 12632256 with the blue-green-red order the model stores: the document spells
    // a colour the way the panel does now, and both spell it through one conversion in Core.
    return /BackColor="#c0c0c0"/.test(String(read.data ?? ""));
  }, { budgetMs: 15000 });
  check("a form property set through the api appears in the OPEN tab's document, live", true);

  const routeMarkup = await api.designerMarkup(form, project);
  check("and in the projection route's document",
    /<Form [^>]*BackColor="#c0c0c0"/.test(routeMarkup),
    routeMarkup.split(/\r?\n/).find((line) => line.includes("BackColor")));

  // The dialect's rule, held from the other side: a document WITHOUT the line cannot erase
  // the colour - an unspoken property is one an apply never touches.
  await api.act("designerApply", { module: form, markup: String(tabMarkup.data) });
  const kept = await api.designer(form, project);
  check("an apply of a document without the property line leaves the colour standing",
    kept.form.backColor === 12632256, String(kept.form.backColor));

  // A document WITH the property changes it - the write goes where the native panel's would.
  // An ATTRIBUTE on the Form's own tag, which is where the tagged dialect puts it: the old line
  // form (`BackColor = 12639424` under the header) is not a thing the parser will take, and a
  // document it refuses applies nothing at all - which is how this row sat failing.
  const recoloured = String(tabMarkup.data).replace(/^(<Form\b[^>]*?)(\s*>)/, '$1 BackColor="12639424"$2');
  await api.act("designerApply", { module: form, markup: recoloured });
  const changed = await api.designer(form, project);
  check("an apply of a document with the line writes the native panel's own property",
    changed.form.backColor === 12639424, String(changed.form.backColor));
  check("12639424, which is what #c0dcc0 means once the bytes are the other way round", true);

  await api.designerEdit("set", { module: form, project, property: "BackColor", value: "-2147483633" });
  const defaulted = await api.designerMarkup(form, project);
  check("back at the default, the line leaves the document - defaults stay unspoken",
    !/BackColor/.test(defaulted));

  // ---- a CONTROL's changed colours ride the document too ----
  //
  // Until 2026-08-16 the document carried identity, containment, geometry and caption and
  // nothing else, so a colour set in the Properties panel was invisible in the text, absent from
  // the draft preview and outside the document's undo. Two baselines decide what "changed"
  // means, and it takes both: a Label inherits the FORM's button face, while a TextBox is born
  // with the WINDOW colours and keeps them on any form.
  // What the document said about colour BEFORE either edit, so the row below measures what these
  // two added rather than assuming the form arrived saying nothing. A control can carry a colour
  // honestly - the saved-design baseline narrows on what the .frx records as changed - and an
  // absolute count would then be reporting the fixture rather than this gesture.
  const colourLines = (text) => text.split(/\r?\n/).filter((line) => /Color=/.test(line));
  const beforeColours = colourLines(await api.designerMarkup(form, project)).length;
  await api.designerEdit("set", { module: form, project, name: "NameBox", property: "BackColor", value: "13434879" });
  await api.designerEdit("set", { module: form, project, name: "Taxable", property: "ForeColor", value: "255" });
  const coloured = await waitFor("the document to carry both colours", async () => {
    const read = await api.designerMarkup(form, project);
    return /BackColor="#ffffcc"/.test(read) && /ForeColor="#ff0000"/.test(read) ? read : null;
  }, { budgetMs: 15000 });
  // ON the control's own tag now, not on a line beneath it: an attribute belongs to the element
  // it is written in, which is the whole reason for the tagged dialect.
  check("a colour set on a control appears on that control, spelled #rrggbb",
    /<[A-Za-z]+\b[^>]*\bName="NameBox"[^>]*\bBackColor="#ffffcc"/.test(coloured)
    && /<[A-Za-z]+\b[^>]*\bName="Taxable"[^>]*\bForeColor="#ff0000"/.test(coloured),
    coloured.split(/\r?\n/).filter((line) => /Color=/.test(line)).join(" | "));

  check("and the controls nobody touched say nothing about colour",
    colourLines(coloured).length === beforeColours + 2,
    `${colourLines(coloured).length} against ${beforeColours} before: `
    + colourLines(coloured).join(" | "));

  await api.designerEdit("set", { module: form, project, name: "NameBox", property: "BackColor", value: "-2147483643" });
  await api.designerEdit("set", { module: form, project, name: "Taxable", property: "ForeColor", value: "-2147483630" });
  await waitFor("the colours to leave when they go back", async () =>
    colourLines(await api.designerMarkup(form, project)).length === beforeColours, { budgetMs: 15000 });
  check("and putting a colour back where it was takes its attribute out again", true);

  // ---- the squiggles: Core's tolerant parse, drawn on the document as it is typed ----

  await api.act("designerSetMarkup", {
    module: form,
    markup: `<Form Name="${form}" Width="100" Height="100">\r\n`
      + '    <Label Name="A" Left="banana" Top="1" />\r\n'
      + '    <Gadget Name="G" Left="1" Top="1" Width="2" Height="2" />\r\n</Form>\r\n',
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

  // ---- the language service: what the document offers, and what it says about it ----
  //
  // Every aim here is computed from the document rather than written down: a fixture that
  // grows a control must not silently re-point these at the wrong line.

  const canonical = String(tabMarkup.data).replace(/\r\n/g, "\n").split("\n");
  const lineOf = (pattern) => canonical.findIndex((text) => pattern.test(text)) + 1;

  const vocabulary = await api.markupVocabulary(form, project);
  const kindsMeasured = vocabulary.kinds.map((one) => one.kind);
  check("the vocabulary route answers every toolbox kind, the form and the page",
    kindsMeasured.includes("Form") && kindsMeasured.includes("Page")
    && kindsMeasured.includes("CommandButton") && kindsMeasured.length >= 16,
    `${kindsMeasured.length}: ${kindsMeasured.join(", ")}`);

  const button = vocabulary.kinds.find((one) => one.kind === "CommandButton");
  check("a kind carries the coclass an apply would create it from",
    button?.progId === "Forms.CommandButton.1", String(button?.progId));
  // `FontSize`, not `Font.Size`: a control's extender carries the font flat, which is what a
  // bare instance measures and what an apply writes. It is Currency on the wire, so it is also
  // the row that would go quiet if the printable kinds ever lost VT_CY again.
  check("and its properties are MEASURED - a bare control's own, with their types and defaults",
    (button?.properties.length ?? 0) > 15
    && button?.properties.some((one) => one.name === "BackStyle" && (one.members?.length ?? 0) > 0)
    && button?.properties.some((one) => one.name === "FontSize" && one.default !== null),
    `${button?.properties.length} properties`);
  check("a colour property says so, so the document's own spelling can be offered",
    button?.properties.some((one) => one.name === "BackColor" && one.colour === true));

  // The Form entry is described from the LIVE form - it has no coclass to instantiate bare -
  // which is the one part of the vocabulary that needs a workbook open.
  const formKind = vocabulary.kinds.find((one) => one.kind === "Form");
  check("the Form's own properties come from the live form, enums and all",
    (formKind?.properties.length ?? 0) > 10
    && formKind?.properties.some((one) => one.name === "Cycle" && (one.members?.length ?? 0) > 0),
    `${formKind?.properties.length} properties`);

  /*
   * THE LANGUAGE SERVICE, ASKED IN THE TAGGED DIALECT.
   *
   * Every position below moved when the documents did, and not by a regex: BETWEEN tags a caret
   * is choosing a child element, INSIDE one it is choosing an attribute, and inside an attribute's
   * quotes it is choosing that attribute's value. Under the line dialect those were one place -
   * a line - which is exactly the ambiguity the tags were adopted to remove.
   */
  const okLine = lineOf(/<CommandButton Name="OkButton"/);
  // A fresh indented line INSIDE the form, before its close: after `</Form>` is outside the
  // document's one root element, and nothing is offered there.
  const freshDoc = addElement(String(tabMarkup.data), "").replace(/    \r\n<\/Form>/, "    \r\n</Form>");
  await api.act("designerSetMarkup", { module: form, markup: freshDoc });
  const freshLine = freshDoc.replace(/\r\n/g, "\n").split("\n").findIndex((text) => text === "    ") + 1;
  const fresh = await api.act("designerComplete", { module: form, line: freshLine, column: 5 });
  const offered = (fresh.data ?? []).map((one) => one.label);
  check("a fresh line inside the form completes the control kinds",
    offered.includes("CommandButton") && offered.includes("Frame") && !offered.includes("Page"),
    offered.slice(0, 8).join(", "));
  check("and a kind arrives scaffolded - a whole element, named, captioned, placed and sized",
    /^<CommandButton Name="\$\{1:[^}]+\}" Caption="\$\{2:[^}]+\}" Left="\$\{3:\d+\}" Top="\$\{4:\d+\}" Width="\$\{5:72\}" Height="\$\{6:24\}" \/>$/
      .test((fresh.data ?? []).find((one) => one.label === "CommandButton")?.insert ?? ""),
    (fresh.data ?? []).find((one) => one.label === "CommandButton")?.insert);

  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });

  // Inside a MultiPage, only a Page: the parser refuses anything else there, and the
  // completions say so early rather than offering an element that cannot land.
  const wizard = lineOf(/^\s+<MultiPage\b[^>]*\bName="Wizard"/);
  const underWizard = await api.act("designerComplete", { module: form, line: wizard + 1, column: 9 });
  const pageOnly = (underWizard.data ?? []).map((one) => one.label);
  check("inside a MultiPage the only kind offered is a Page",
    pageOnly.includes("Page") && !pageOnly.includes("Label"),
    pageOnly.slice(0, 8).join(", "));

  // The two halves the tags separate. INSIDE the Frame's body, the child kinds it may hold;
  // INSIDE its own tag, its own properties. Under the line dialect both arrived at one caret,
  // and a developer could not tell which of the two a suggestion would become.
  const options = lineOf(/^\s+<Frame\b[^>]*\bName="Options"/);
  const frameText = canonical[options - 1] ?? "";
  const inFrameBody = (await api.act("designerComplete", { module: form, line: options + 1, column: 9 })).data ?? [];
  check("inside a Frame's body the completions offer the controls it may hold",
    inFrameBody.some((one) => one.label === "OptionButton")
    && !inFrameBody.some((one) => one.label === "SpecialEffect"),
    inFrameBody.map((one) => one.label).slice(0, 8).join(", "));

  const inFrameTag = (await api.act("designerComplete",
    { module: form, line: options, column: frameText.lastIndexOf(">") + 1 })).data ?? [];
  check("and inside its own tag they offer its properties, the ones it is not already spelling",
    inFrameTag.some((one) => one.label === "BorderStyle")
    && !inFrameTag.some((one) => one.label === "SpecialEffect")
    && !inFrameTag.some((one) => one.label === "OptionButton"),
    `${inFrameTag.length}: ${inFrameTag.map((one) => one.label).slice(0, 8).join(", ")}`);

  // Inside an attribute's own quotes: that attribute's enum members, by the name a developer
  // writes rather than the number the model keeps.
  const okText = canonical[okLine - 1] ?? "";
  const valueColumn = okText.indexOf('PicturePosition="') + 'PicturePosition="'.length + 1;
  const values = (await api.act("designerComplete",
    { module: form, line: okLine, column: valueColumn })).data ?? [];
  check("inside an attribute's quotes the completions offer that property's enum members by name",
    values.some((one) => one.label === "fmPicturePositionLeftCenter")
    && values.some((one) => one.label === "fmPicturePositionCenter"),
    values.map((one) => one.label).slice(0, 5).join(", "));

  const valueHover = (await api.act("designerHover",
    { module: form, line: okLine, column: okText.indexOf("PicturePosition") + 3 })).data ?? [];
  check("and hovering an attribute says what it takes and what an untouched control holds",
    valueHover.some((block) => /PicturePosition As fmPicturePosition/.test(block))
    && valueHover.some((block) => /Default `fmPicturePosition/.test(block)),
    JSON.stringify(valueHover));

  // Hover on a kind: what that class of control IS, and nothing the tag already says - which is
  // where the geometry used to be read back, and now also where the coclass used to be. For a
  // standard kind `Forms.Frame.1` is the word under the pointer with a prefix and a suffix, so
  // the card was spending a line restating its own heading (the owner, 2026-08-16).
  const kindHover = (await api.act("designerHover",
    { module: form, line: options, column: frameText.indexOf("Frame") + 3 })).data ?? [];
  check("hovering a kind describes the class, and does not read the tag back at the reader",
    kindHover.some((block) => /groups controls/.test(block))
    && !kindHover.some((block) => /Forms\.Frame\.1/.test(block))
    && !kindHover.some((block) => /\bat\b|points/.test(block)),
    JSON.stringify(kindHover));

  const nameHover = (await api.act("designerHover",
    { module: form, line: okLine, column: okText.indexOf("OkButton") + 3 })).data ?? [];
  check("hovering a control's name declares it the way VBA would, class and all",
    nameHover.some((block) => /OkButton As MSForms\.CommandButton/.test(block)),
    JSON.stringify(nameHover));

  // The hint: the tag's grammar, following the ATTRIBUTE the hand is on. It follows the name
  // standing rather than a count, because attributes may be written in any order.
  const afterLeft = okText.indexOf('Left="262"') + 'Left="262"'.length + 1;
  const hint = (await api.act("designerHint", { module: form, line: okLine, column: afterLeft })).data;
  check("the tag hint shows the grammar and points at the attribute being typed",
    hint?.label === 'Name="..." [Caption="..."] [Left="0" Top="0"] [Width="60" Height="20"]'
    && hint?.parameter === '[Left="0" Top="0"]',
    JSON.stringify(hint));

  const formHint = (await api.act("designerHint", { module: form, line: 1, column: 7 })).data;
  check("and the Form's own tag has its own grammar, which takes no position",
    formHint?.label === 'Name="..." [Caption="..."] [Width="240" Height="180"]',
    JSON.stringify(formHint));

  // ---- the canvas follows the document: the draft previews, the form untouched ----

  const draftMarkup = addElement(String(tabMarkup.data),
    '<CommandButton Name="DraftBtn" Caption="Soon" Left="8" Top="282" Width="60" Height="20" />');
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
  await api.act("designerSetMarkup", {
    module: form,
    markup: `${draftMarkup}    <Label Name="Broken" Left="banana" Top="1" />\r\n`,
  });
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
    markup: String(tabMarkup.data).replace(/\bHeight="[\d.]+"/, 'Height="900"'),
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
  // The raw text is kept beside the split: restoring from a re-joined split puts the document
  // back with the WRONG line endings, and a document that differs from canonical by \r alone is
  // dirty for ever - which is a wait that never lands rather than a row that fails.
  const markupText = String((await api.act("designerMarkup", { module: form })).data);
  const markupLines = markupText.split(/\r?\n/);
  check("and the markup caret lands on the selected control's line",
    markupLine > 1 && /RegionPick/.test(markupLines[markupLine - 1] ?? ""), `line ${markupLine}`);

  // The WHOLE BLOCK is washed, not just the header: a control is a block in this language -
  // its properties and, for a container, its children - so a caret alone said where it started
  // and left the reader to find where it stopped.
  // Asked of the DOCUMENT through the snapshot, not counted off the screen: monaco renders a
  // decoration on the next frame and only for the lines in view, so a DOM count measures its
  // renderer and the scroll position - which is how the first cut of this row read two lines
  // for a three-line block and failed while the feature worked.
  const blockOn = async (control) => {
    await api.act("designerSelect", { module: form, ...(control ? { control } : {}) });
    const block = (await api.act("designerCanvas", { module: form })).data?.markupBlock;
    return block ? block.to - block.from + 1 : 0;
  };

  // The expectation is read off the DOCUMENT rather than written down: an M1 row above removes
  // one of the Frame's option buttons for good, so a literal "3" here is a row that passes
  // until somebody reorders the suite and then fails for a reason that is not a defect.
  const framedLines = (() => {
    const at = markupLines.findIndex((text) => /<Frame\b[^>]*\bName="Options"/.test(text));
    if (at < 0) {
      return 0;
    }

    const indent = (text) => text.length - text.trimStart().length;
    let last = at;
    while (last + 1 < markupLines.length
      && markupLines[last + 1].trim().length > 0
      && indent(markupLines[last + 1]) > indent(markupLines[at])) {
      last += 1;
    }

    // AND ITS CLOSING TAG, which the indent walk cannot reach: `</Frame>` sits at the SAME indent
    // as the opening, so the loop stops one line short of the element the selection covers.
    if (last + 1 < markupLines.length && /^\s*<\/Frame>/.test(markupLines[last + 1])) {
      last += 1;
    }

    return last - at + 1;
  })();

  const framedBlock = await blockOn("Options");
  check("selecting a container covers its whole block in the markup, children included",
    framedBlock === framedLines && framedBlock > 1,
    `${framedBlock} line(s) where the document indents ${framedLines} under the Frame`);

  // A CHILDLESS CONTROL IS ONE LINE, and that is the tagged dialect's own answer rather than a
  // convenience: its properties ride as ATTRIBUTES on its element instead of as lines beneath it,
  // so a ComboBox carrying MatchEntry and ShowDropButtonWhen from the saved baseline is still a
  // single self-closing element. Under the line dialect this counted the property lines under the
  // header, which is the shape that no longer exists.
  const plainElement = markupLines.find((text) => /<ComboBox\b[^>]*\bName="RegionPick"/.test(text)) ?? "";

  const plainBlock = await blockOn("RegionPick");
  check("and a plain control covers its own element, which is a single line",
    plainBlock === 1 && /\/>\s*$/.test(plainElement)
    && /\bMatchEntry="/.test(plainElement) && /\bShowDropButtonWhen="/.test(plainElement),
    `${plainBlock} line(s): ${plainElement.trim()}`);

  // Waited for, not read at once: monaco paints a decoration on the NEXT frame, so a count taken
  // in the same tick as the selection is a race the row loses whenever the surface is busy.
  const washed = await waitFor("the block wash to be painted", async () =>
    Number(await api.ask(`${inViewAll(".designer-block-mark")}.length`)) >= 1 ? true : null,
  { budgetMs: 5000 }).catch(() => false);
  check("and the wash is drawn on the lines the block names", washed === true,
    "monaco draws a decoration for the lines in view, so this counts what is on screen");

  // Markup to canvas, the other direction: a caret INSIDE a control's block selects that
  // control, because a property line belongs to the thing it describes.
  const caretPicks = async (line) => {
    await api.act("designerCaret", { module: form, line });
    return (await api.act("designerCanvas", { module: form })).data?.selected;
  };

  const framedLine = markupLines.findIndex((text) =>
    /<OptionButton\b[^>]*\bName="PickGround"/.test(text)) + 1;
  check("a caret on a control's own line selects it on the canvas",
    await caretPicks(framedLine) === "PickGround", `line ${framedLine}`);

  const frameLine = markupLines.findIndex((text) =>
    /<Frame\b[^>]*\bName="Options"/.test(text)) + 1;
  check("and a caret on the container's line selects the container, not its child",
    await caretPicks(frameLine) === "Options", `line ${frameLine}`);

  check("and a caret on the Form line selects the form", await caretPicks(1) === "");

  // A PENDING RENAME DOES NOT DROP THE PANEL ONTO THE FORM. The panel reads the live control, so
  // a name the document has only just invented is one it cannot find, and its answer for a
  // control it cannot find is the form. Renaming in the markup therefore left the caret and the
  // canvas on the control while the panel jumped to the UserForm (the owner, 2026-08-17).
  {
    const renamed = markupLines.map((line) =>
      line.replace(/(<CheckBox\b[^>]*\bName=")Taxable(")/, "$1TaxablePending$2")).join("\n");
    await api.act("designerSelect", { module: form, control: "Taxable" });
    await waitFor("the panel on the control before the rename", async () =>
      (await api.ui()).properties?.component === "Taxable", { budgetMs: 15000 });

    await api.act("designerSetMarkup", { module: form, markup: renamed });
    await waitFor("the canvas to preview the renamed control", async () =>
      ((await api.act("designerCanvas", { module: form })).data?.controls ?? [])
        .some((one) => one.name === "TaxablePending"), { budgetMs: 15000 });
    const onRenamed = await caretPicks(markupLines.findIndex((line) =>
      /<CheckBox\b[^>]*\bName="Taxable"/.test(line)) + 1);
    check("a caret on a control the markup has RENAMED still selects it on the canvas",
      onRenamed === "TaxablePending", String(onRenamed));
    check("and the panel stays on the control rather than falling back to the form",
      (await api.ui()).properties?.component === "Taxable",
      String((await api.ui()).properties?.component));

    await api.act("designerSetMarkup", { module: form, markup: markupText });
    await waitFor("the document back after the rename row", async () =>
      ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });
  }

  await api.act("designerSelect", { module: form });
  const formSelected = (await api.act("designerCanvas", { module: form })).data;
  check("clicking the form's own ground selects the form, caret to the Form line",
    formSelected.selected === "" && Number(formSelected.markupLine) === 1,
    JSON.stringify({ selected: formSelected.selected, line: formSelected.markupLine }));

  // A press on the canvas but OFF the form selects the form too. It used to select nothing and
  // leave the last control dressed, so the panel and the handles went on describing something
  // the developer had clicked away from.
  await api.act("designerSelect", { module: form, control: "RegionPick" });
  const offForm = await api.ask(`(() => {
    const view = document.querySelector('.designer-view[data-module=${JSON.stringify(form)}]');
    const scroll = view.querySelector('.designer-canvas-scroll');
    const box = scroll.getBoundingClientRect();
    const shape = view.querySelector('.dc-form').getBoundingClientRect();
    // Below the form, inside the scroll box: canvas ground, not the form's.
    const y = Math.min(box.bottom - 4, shape.bottom + 12);
    scroll.dispatchEvent(new PointerEvent('pointerdown', {
      bubbles: true, cancelable: true, pointerId: 1, isPrimary: true, button: 0, buttons: 1,
      clientX: box.left + 12, clientY: y,
    }));
    return y > shape.bottom ? 'pressed below the form' : 'no room below the form';
  })()`);
  check("a press on the canvas but off the form selects the form",
    (await api.act("designerCanvas", { module: form })).data?.selected === "",
    `${offForm}; the canvas says `
    + JSON.stringify((await api.act("designerCanvas", { module: form })).data?.selected));

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
  // Scrolled to first, because a hand cannot click what the canvas is not showing: on a narrow
  // window the form is wider than the box that holds it, and the rect of a clipped control still
  // reads as a position - one that belongs to whatever paints there instead.
  const hitAnswer = await api.ask(`(() => { const el = ${inViewAll(".dc")}.find(e => e.dataset.control === "RegionPick"); if (!el) return "no element"; el.scrollIntoView({ block: "nearest", inline: "nearest" }); const r = el.getBoundingClientRect(); const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2); return el === hit || el.contains(hit) ? "the control" : (hit?.className || "nothing"); })()`);
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

  /*
   * A PANEL ROW ON A DESIGNER TAB EDITS THE DOCUMENT AND WAITS FOR Ctrl+S, since 2026-08-17
   * (task #68). These rows watched the FORM immediately, which was the old contract: the panel
   * wrote the form over COM and the text caught up by echo. That put its edits outside the undo
   * stack, and the owner reported both halves of it ("ctrl+z is not undoing it", then "ctrl+z
   * updates the markdown editor, but not the designer"). So what is asserted now is the deal the
   * rest of the tab already makes - document first, form on the save - and the form NOT moving
   * before the save is as much the claim as its moving after.
   */
  const panelMove = await api.act("editProperty", { name: "Left", value: "90" });
  check("a control row's edit answers through the model", panelMove.did === true, panelMove.detail);
  await waitFor("the open tab's document to carry the panel's move", async () =>
    (await placed("RegionPick")).left === 90, { budgetMs: 15000 });
  check("a panel row writes the DOCUMENT, the way every other gesture on the tab does", true);
  check("and the FORM is untouched until the save",
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 84) < 0.01,
    String((await api.designer(form, project)).controls.find((c) => c.name === "RegionPick")?.left));

  await api.command("save");
  await waitFor("the save to carry the panel's move to the form", async () =>
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 90) < 0.01, { budgetMs: 15000 });
  check("and Ctrl+S carries it to the form - liveness end to end", true);

  await api.act("editProperty", { name: "Left", value: "84" });
  await api.command("save");
  await waitFor("and back where the plan puts it", async () =>
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 84) < 0.01, { budgetMs: 15000 });

  await api.act("designerSelect", { module: form });
  await waitFor("the form's ground to return the panel to the form", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });
  check("selecting the form returns the panel to the component", true);

  // AND A CONTROL REMOVED UNDER THE PANEL takes the panel back to the form with it. Only the TAB
  // was re-projected on a removal, so the panel went on describing a control that was gone -
  // name, geometry, font and all (found in the 2026-08-16 hunt). The panel's own publish already
  // knew what to do; nothing was asking it.
  await api.designerEdit("add", {
    module: form, project, type: "label", name: "Doomed", left: 4, top: 4, width: 30, height: 12,
  });
  await api.act("designerSelect", { module: form, control: "Doomed" });
  await waitFor("the panel to aim at the doomed control", async () =>
    (await api.ui()).properties?.component === "Doomed", { budgetMs: 15000 });
  await api.designerEdit("remove", { module: form, project, name: "Doomed" });
  await waitFor("the panel to let go of a control that is gone", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });
  check("a control removed under the panel takes the panel back to the form", true);

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

  // A SYSTEM colour is not a colour but a question - what does this machine call a button face -
  // so the row shows the question by name and the swatch paints today's answer. Spelling it
  // `#f0f0f0` would freeze the answer into the form.
  const colourRow = formPanel.rows.find((row) => row.name === "BackColor");
  check("a system colour shows its name, with the machine's own answer on the swatch",
    colourRow?.value === "Button Face" && /^#[0-9a-f]{6}$/.test(String(colourRow?.swatch)),
    JSON.stringify(colourRow));

  // AN OBJECT IS NOT A VALUE. A form's panel used to list six rows reading "[object]" - Font,
  // Picture, MouseIcon, Controls, ActiveControl, Selected - each of them a row asking to be
  // ignored. The font is the one this panel can serve, and it serves it as its parts.
  check("no row says [object]; an object is not a value a panel can show",
    !formPanel.rows.some((row) => String(row.value).includes("[object]")),
    formPanel.rows.filter((row) => String(row.value).includes("[object]")).map((row) => row.name).join(", "));
  check("a form's FONT shows as its parts, the same rows a control's selection draws",
    ["Font.Name", "Font.Size", "Font.Bold", "Font.Italic"]
      .every((name) => formPanel.rows.some((row) => row.name === name)),
    formPanel.rows.filter((row) => row.name.startsWith("Font")).map((row) => row.name).join(", "));
  check("and the library's own hidden members are not listed at all",
    !formPanel.rows.some((row) => row.name.startsWith("_") || row.name === "DesignMode"));

  // ---- the font rows are pickers, from the machine's own list ----
  //
  // Measured rather than written down, for the reason the system colours are: no list here could
  // be true on another machine. So the row is checked for the PROPERTIES of a real enumeration -
  // sorted, no vertical variants, and holding the face the model itself reported - rather than
  // against a copy of the list, which would be the same measurement written twice.
  const faceRow = formPanel.rows.find((row) => row.name === "Font.Name");
  const faces = faceRow?.options ?? [];
  // Non-decreasing under the host's own comparison (ordinal, ignoring case) rather than re-sorted
  // here: a locale collator disagrees with it about spaces and digits, and the row's claim is that
  // the list came back in order, not that two sorting rules agree.
  const outOfOrder = faces.findIndex((one, at) =>
    at > 0 && faces[at - 1].toUpperCase() > one.toUpperCase());
  check("the Font.Name row offers this machine's own faces, in order, the current one among them",
    faces.length > 10 && faces.includes(String(faceRow?.value)) && outOfOrder === -1,
    `${faces.length} face(s), showing ${faceRow?.value}`
    + (outOfOrder > 0 ? `; ${faces[outOfOrder - 1]} before ${faces[outOfOrder]}` : ""));
  check("and no '@' vertical variants, which every native font list hides",
    !faces.some((one) => one.startsWith("@")),
    faces.filter((one) => one.startsWith("@")).join(", "));

  const sizeRow = formPanel.rows.find((row) => row.name === "Font.Size");
  check("the Font.Size row offers the usual ramp and still holds a number",
    (sizeRow?.options ?? []).includes("12") && /^\d+(\.\d+)?$/.test(String(sizeRow?.value)),
    JSON.stringify(sizeRow));

  // Picked from what the row OFFERS rather than from a face named here: a machine without Courier
  // New would fail a row that spelled one, and the point of the list is that it is this machine's.
  // On a CONTROL, because the walk carries a control's font and the form's box does not - so the
  // write can be read back off the model rather than off the panel that sent it.
  await api.act("designerSelect", { module: form, control: "OkButton" });
  const buttonPanel = await waitFor("the panel to show the button", async () => {
    const shown = (await api.ui()).properties;
    return shown?.component === "OkButton" ? shown : null;
  }, { budgetMs: 15000 });
  const buttonFace = buttonPanel.rows.find((row) => row.name === "Font.Name");
  const pick = (buttonFace?.options ?? [])
    .find((one) => one !== buttonFace?.value && /^[A-Za-z][A-Za-z ]+$/.test(one));
  const wroteFace = await api.act("editProperty", { name: "Font.Name", value: pick });
  check("choosing a face from the list writes it through",
    wroteFace.did === true && (wroteFace.detail ?? "").includes(pick), wroteFace.detail);
  await waitFor("the control to wear it", async () =>
    (await api.designer(form, project)).controls.find((one) => one.name === "OkButton")?.font?.name === pick,
  { budgetMs: 15000 });
  check(`${pick}, off the machine's own list and onto the control`, true);

  await api.act("editProperty", { name: "Font.Name", value: String(buttonFace?.value) });
  await api.act("designerSelect", { module: form });
  await waitFor("the panel back on the form for the rows that follow", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });

  // And the DOCUMENT offers the same list, because it is the same measurement: one walk, two
  // surfaces, so the panel and the markup cannot hold different ideas of what fonts exist.
  const fontVocabulary = (await api.markupVocabulary(form, project)).kinds
    ?.find((kind) => kind.kind === "CommandButton")?.properties
    ?.find((one) => one.name === "FontName");
  check("the markup's vocabulary carries the same faces the panel offers",
    (fontVocabulary?.values ?? []).length === faces.length, `${(fontVocabulary?.values ?? []).length} vs ${faces.length}`);

  // ...and the document offers them where a face goes. A face is a STRING in the dialect, so what
  // a suggestion REPLACES matters as much as what it inserts: accepting one where the developer
  // has already typed `"Tah` has to take the quote with it, or the line gains a second string
  // inside the first.
  //
  // IN THE ATTRIBUTE'S OWN QUOTES, which is where a face goes now: `FontName` is written inside
  // the button's tag rather than on a line beneath it, so the probe puts a half-typed attribute
  // there and asks what the caret is offered.
  const canonicalText = String(tabMarkup.data);
  const withFont = canonicalText.split(/\r?\n/);
  const buttonLine = withFont.findIndex((line) => /<CommandButton Name="OkButton"/.test(line));
  const tryAttribute = async (partial) => {
    const lines = [...withFont];
    const original = lines[buttonLine];
    // Spliced in just before the tag closes, so the rest of the element is untouched.
    const at = original.lastIndexOf("/>");
    const probe = `${original.slice(0, at)}${partial}${original.slice(at)}`;
    lines[buttonLine] = probe;
    await api.act("designerSetMarkup", { module: form, markup: lines.join("\n") });
    await waitFor("the document to hold the probe attribute", async () =>
      String((await api.act("designerMarkup", { module: form })).data).includes(partial.trim()),
    { budgetMs: 15000 });
    const done = await api.act("designerComplete",
      { module: form, line: buttonLine + 1, column: at + partial.length + 1 });
    const offer = (done.data ?? []).find((one) => one.label === pick);
    return offer
      ? probe.slice(0, offer.replaces.from - 1) + offer.insert + probe.slice(offer.replaces.to - 1)
      : `no ${pick} among ${(done.data ?? []).length}`;
  };

  const spelled = (partial) => {
    const original = withFont[buttonLine];
    const at = original.lastIndexOf("/>");
    return `${original.slice(0, at)}FontName="${pick}"${original.slice(at)}`;
  };

  check("a face completes inside FontName's own quotes",
    (await tryAttribute('FontName="')) === spelled('FontName="'),
    await tryAttribute('FontName="'));
  check("and accepting one replaces the value already begun, closing quote and all",
    (await tryAttribute('FontName="Cou')) === spelled('FontName="Cou'),
    await tryAttribute('FontName="Cou'));

  await api.act("designerSetMarkup", { module: form, markup: canonicalText });
  await waitFor("the document back at canonical after the font rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });

  // WRITABLE is a fact about the type, not about the value that came back - and a property with
  // no setter is now LEFT OUT rather than drawn grey (the owner, 2026-08-16: "if theyre jot
  // settable; dint suow them"). CanPaste and InsideWidth were the two shapes of it: a question
  // about the editing session, and a measurement of the form. The code name stays, because it is
  // the VBE's rename gesture rather than a property the designer's library has any say over.
  check("a property with no setter is not shown at all, and the code name still is",
    !formPanel.rows.some((row) => ["CanPaste", "CanUndo", "CanRedo", "InsideWidth", "InsideHeight"]
      .includes(row.name))
    && formPanel.rows.find((row) => row.name === "(Name)")?.writable === true,
    formPanel.rows.filter((row) => !row.writable).map((row) => row.name).join(", ") || "every row writes");

  // The honest limit, pinned: names come from the TYPE LIBRARY, never from the property's own
  // name. StartUpPosition is a plain Integer in MSForms however much it looks like an enum, so
  // it keeps its number rather than being given a vocabulary this product invented.
  const plainRow = formPanel.rows.find((row) => row.name === "StartUpPosition");
  check("a property the library does not name keeps its number - nothing is guessed",
    /^-?\d+$/.test(String(plainRow?.value)) && !plainRow?.options, JSON.stringify(plainRow));

  // Back on the FORM before the rows that write its own properties: the font probes above set the
  // document, and setting it re-selects whatever the caret then lands in.
  await api.act("designerSelect", { module: form });
  await waitFor("the panel on the form for the property-writing rows", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });

  const wroteName = await api.act("editProperty", { name: "Cycle", value: "fmCycleCurrentForm" });
  check("a member NAME writes through to the model", wroteName.did === true, wroteName.detail);

  // ...and the number still writes, which is what keeps the row a text field rather than a
  // list: the developer types either. 0 is fmCycleAllForms and fmCycleAllForms is the DEFAULT, so
  // what the document does with it is take the attribute back out - "defaults stay unspoken" on
  // the write side as well as the print side, and the act says so rather than echoing a value.
  const wroteNumber = await api.act("editProperty", { name: "Cycle", value: "0" });
  check("the raw number writes too, and a default takes the attribute out again",
    wroteNumber.did === true && /default/.test(wroteNumber.detail ?? ""), wroteNumber.detail);
  check("and the document says nothing about Cycle once it is back at the default",
    !/\bCycle=/.test(String((await api.act("designerMarkup", { module: form })).data)),
    String((await api.act("designerMarkup", { module: form })).data).split(/\r?\n/)[0]);

  const wroteHex = await api.act("editProperty", { name: "BackColor", value: "&H00C0FFC0&" });
  check("a colour written as the VBE's hex lands as the model's own number",
    wroteHex.did === true, wroteHex.detail);
  await api.command("save");
  await waitFor("the form to carry the colour", async () =>
    ((await api.designer(form, project)).form?.backColor ?? 0) === 12648384, { budgetMs: 15000 });
  check("12648384, which is what &H00C0FFC0& means", true);
  // The row reads back CANONICALLY, not as the developer spelled it. Waited for rather than read
  // at once: since the panel writes the document (#68) the correction arrives with the host's
  // republish after the apply, where it used to come straight back from the COM write.
  const readBack = await waitFor("the panel row to read the colour back canonically", async () => {
    const row = (await api.ui()).properties.rows.find((one) => one.name === "BackColor");
    return row?.value === "#c0ffc0" ? row : null;
  }, { budgetMs: 15000 }).catch(() => null);
  check("and a plain colour reads back the way a developer writes one",
    readBack !== null,
    JSON.stringify((await api.ui()).properties.rows.find((row) => row.name === "BackColor")));

  // The spelling the owner asked for, in both directions.
  const wroteCss = await api.act("editProperty", { name: "BackColor", value: "#ff8000" });
  check("a colour written as #rrggbb lands too - byte order and all",
    wroteCss.did === true && wroteCss.detail?.includes("#ff8000"), wroteCss.detail);
  await api.command("save");
  await waitFor("the form to carry that one", async () =>
    ((await api.designer(form, project)).form?.backColor ?? 0) === 33023, { budgetMs: 15000 });
  check("33023, which is #ff8000 with the blue-green-red order the model stores", true);

  const wroteSystem = await api.act("editProperty", { name: "BackColor", value: "Highlight" });
  check("and a system colour written by NAME asks the question rather than freezing the answer",
    wroteSystem.did === true, wroteSystem.detail);
  await api.command("save");
  await waitFor("the form to hold the system colour", async () =>
    ((await api.designer(form, project)).form?.backColor ?? 0) === -2147483635, { budgetMs: 15000 });
  check("-2147483635, the system's own highlight index with the high bit set", true);

  // ---- the picker itself: xlide's own, not the platform's ----

  const opened = await api.act("colourPicker", { property: "BackColor" });
  check("the swatch opens a picker with both halves - a palette, and this machine's system colours",
    opened.did === true && (opened.data?.palette ?? 0) >= 64 && (opened.data?.system ?? 0) >= 24,
    opened.detail);

  // Aimed at what the picker ACTUALLY offers rather than at a colour written down here: the
  // palette is generated, so a suite that names a hex is a suite that breaks when the ramp moves.
  const paletteOffers = String(await api.ask(
    "(() => Array.from(document.querySelectorAll('.colour-swatch')).map(s => s.dataset.value).join(','))()"))
    .split(",");
  const wanted = paletteOffers[35] ?? "#000000";
  const picked = await api.act("colourPicker", { choose: wanted });
  check("picking a palette colour writes it through the row's own commit",
    picked.did === true, picked.detail);
  // Ctrl+S, because the picker commits through the same row the typed value does and that row
  // now writes the DOCUMENT (#68): the form takes it at the save like every other gesture.
  await api.command("save");
  await waitFor("the form to take the picked colour", async () => {
    const back = (await api.designer(form, project)).form?.backColor ?? 0;
    const css = `#${(back & 0xFF).toString(16).padStart(2, "0")}`
      + `${((back >> 8) & 0xFF).toString(16).padStart(2, "0")}`
      + `${((back >> 16) & 0xFF).toString(16).padStart(2, "0")}`;
    return css === wanted;
  }, { budgetMs: 15000 });
  check(`${wanted} reached the form, by the same path a typed value takes`, true);

  check("and the picker closes behind the choice",
    (await api.act("colourPicker", {})).did === false);

  const pickedSystem = await api.act("colourPicker", { property: "BackColor", choose: "Button Face" });
  check("picking a SYSTEM colour writes the question, not the colour it resolves to today",
    pickedSystem.did === true, pickedSystem.detail);
  await api.command("save");
  await waitFor("the form to hold the system colour again", async () =>
    ((await api.designer(form, project)).form?.backColor ?? 0) === -2147483633, { budgetMs: 15000 });
  check("-2147483633: the row reads Button Face, and the form asks the machine every time it paints",
    (await api.ui()).properties.rows.find((row) => row.name === "BackColor")?.value === "Button Face");

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
  check("every dropdown in the pane is the same one - an enum, a colour and a flag all match",
    String(rowControls) === "TextAlign=caret BackColor=swatch Enabled=caret Left=plain",
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

  // THE GRID GOES OFF FOR THE GESTURE ROWS, and has its own section after them.
  //
  // It ships ON, which is the right default and the wrong thing for a row that drags by twelve
  // points and names where it should land: with snapping the drop goes to the nearest six, and
  // the row would be measuring the grid rather than the commit path it was written for. Off
  // here, on there, and the restore is at the end of the grid's own section.
  const shipped = await api.settings();
  check("snapping ships set to the grid, at six points - the editor's own Align Controls to Grid",
    shipped.designerSnap === "grid" && shipped.designerGridSize === 6,
    JSON.stringify({ snap: shipped.designerSnap, size: shipped.designerGridSize }));

  await api.settings({ designerSnap: "off" });
  await waitFor("the page to hear that snapping is off", async () =>
    (await api.ui()).settings?.designerSnap === "off", { budgetMs: 15000 });

  // The real pointer sequence - press on what the hit test answers, past the threshold, drop -
  // so the row proves the gesture, not the arithmetic behind it.
  const dragged = await api.act("designerDrag", { module: form, control: "RegionPick", dx: 24, dy: 12 });
  check("a pointer drag on the canvas moves a control", dragged.did === true, dragged.detail);
  await waitFor("the drop to reach the document", async () =>
    (await placed("RegionPick")).left === 108 && (await placed("RegionPick")).top === 50,
  { budgetMs: 15000 });
  check("the drop rewrites the control's element - 84,38 becomes 108,50", true);

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
    (await placed("RegionPick")).left === 84 && (await placed("RegionPick")).top === 38,
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
    (await placed("RegionPick")).left === 85 && (await placed("RegionPick")).top === 38,
  { budgetMs: 15000 });
  check("an arrow nudges the selection by a single point", true);

  // A second nudge is a SECOND undo step. The stack stop before each move is what keeps one
  // gesture from swallowing the one before it, and one Ctrl+Z gives back exactly one.
  await press("ArrowRight");
  await waitFor("the second nudge", async () =>
    (await placed("RegionPick")).left === 86 && (await placed("RegionPick")).top === 38,
  { budgetMs: 15000 });
  await press("z", "ctrlKey: true,");
  await waitFor("one gesture back", async () =>
    (await placed("RegionPick")).left === 85 && (await placed("RegionPick")).top === 38,
  { budgetMs: 15000 });
  check("each move is its own undo step - one Ctrl+Z gives back one gesture, not the pair", true);

  // A drag with nowhere to go cannot lose the control behind an edge: it stops at the corner.
  await api.act("designerDrag", { module: form, control: "RegionPick", dx: -400, dy: -400 });
  await waitFor("the clamped drop", async () =>
    (await placed("RegionPick")).left === 0 && (await placed("RegionPick")).top === 0,
  { budgetMs: 15000 });
  check("a drag past the top-left stops at the parent's corner rather than going negative", true);

  // ---- F5 runs the DOCUMENT, and touches nothing else ----
  //
  // Two pins with one run. The form catches up on an APPLY, so a run that skipped it launched
  // the last one: undo a move, press F5, and the form standing on screen still holds the move
  // while the canvas beside it does not (the owner, 2026-08-16, with a screenshot of a Hold
  // button in two places). Run applies first now - and applies ONLY: the save that once rode
  // under F5 broke parity with the native editor, and on a never-saved workbook raised the
  // Save As dialog behind the modal form, where cancelling it after the form closed unwound
  // two interleaved modal loops and sometimes killed the host (the owner, 2026-08-27). So the
  // second pin is the SAVED baseline read before and after the run, byte-identical: the
  // launched window is the document, and the file never felt it.

  const stale = (await api.designer(form, project)).controls.find((c) => c.name === "RegionPick");
  check("the form has not caught up with the document yet - which IS the hazard",
    Math.abs(Number(stale?.left)) > 0.5 || Math.abs(Number(stale?.top)) > 0.5,
    `the form holds ${stale?.left},${stale?.top} while the document says 0,0; if these already `
    + "agree the row below proves nothing");

  // THE PRECONDITION, said out loud. A form whose code-behind no longer compiles does not run,
  // and the Run command answers "executed" either way - so a broken project reads here as a run
  // that silently never happened (2026-08-17, chasing exactly that). The code-behind names
  // RegionPick, Taxable, Views, ViewNote and NameBox, and Option Explicit makes a missing one a
  // compile error, so a row above that removed or renamed one of them lands here rather than
  // where it happened.
  const compiles = await api.compile();
  check("the project still compiles, so a run that does not stand means what it says",
    compiles.compiled === true,
    JSON.stringify({ compiled: compiles.compiled, errors: compiles.errors, detail: compiles.detail }));

  // A DOOR INTO THE POISONED SESSION. This row fails only inside a full run (#71) and never
  // standalone, so the way to see it is to stop here and drive the live session by hand:
  //   set XLIDE_PAUSE=1 && node tools\harness\designer-features.mjs
  // The suite then holds everything exactly as it is until the pause runs out, and the api
  // client can be pointed at it from another shell.
  if (process.env.XLIDE_PAUSE) {
    console.log(`  .... paused before the dirty run; the session is yours (form ${form})`);
    await new Promise((resume) => setTimeout(resume, 600000));
  }

  // The file's word BEFORE the run, through the workbook's own storage. Read here rather
  // than at the section top so a slow apply from the rows above cannot sit between the two
  // reads and blur what the run itself did.
  const fileBefore = JSON.stringify(await api.designerBaseline(form, project));

  const ranDirty = api.command("run");
  // SAYS WHY WHEN IT DOES NOT STAND, rather than throwing a bare timeout. A run that fails here
  // fails for one of a few reasons - the mode never left design, a dialog is sitting in front of
  // it, the project stopped compiling - and a row that reports none of them sends the next reader
  // to the window-event log to guess (2026-08-17, twice).
  const stood = await waitFor("the form to stand running", async () =>
    ((await api.userforms()).forms ?? []).find((title) => title.includes("Quarter Entry")),
  { budgetMs: 20000 }).catch(() => null);
  if (!stood) {
    const why = {
      mode: (await api.state()).debugMode,
      forms: (await api.userforms()).forms ?? [],
      dialogs: ((await api.dialogs()).dialogs ?? []).map((one) => one.title ?? one),
      compiles: (await api.compile()).compiled,
    };
    check("the form stands after a run over a dirty designer", false, JSON.stringify(why));
  } else {
    check("the form stands after a run over a dirty designer", true, String(stood));
  }

  const dirtyRunAnswer = await ranDirty;
  check("F5 over a dirty designer answers that it ran", dirtyRunAnswer.ran === true,
    dirtyRunAnswer.detail);

  /*
   * A CLOSED FORM IS NOT YET AN UNLOADED ONE, AND THE DESIGNER GOES WITH IT (measured 2026-08-16).
   *
   * The close is posted, and for roughly 400ms afterwards the component answers "no designer to
   * read" - the object is not there to be asked, so the route throws and the whole group aborts.
   * `debugMode` is no guide at all: the log has it back at design while the form still stands.
   * Nor is the running list, which empties before the designer returns. So this waits for the
   * designer to come back, which is a different question from where the control is - the read
   * below is the assertion, and it must be a check that can fail rather than a wait that throws.
   */
  await api.userforms("close", "Quarter Entry");
  await waitFor("the form to unload and give its designer back", async () =>
    await api.designer(form, project).then(() => true, () => false), { budgetMs: 20000 });

  const ranWith = (await api.designer(form, project)).controls.find((c) => c.name === "RegionPick");
  check("and what it launched was the DOCUMENT - applied on the way to the form",
    Math.abs(Number(ranWith?.left)) < 0.51 && Math.abs(Number(ranWith?.top)) < 0.51,
    `the form holds ${ranWith?.left},${ranWith?.top} and the document said 0,0`);

  check("so the tab is clean afterwards, the way an apply leaves it",
    (await api.act("designerCanvas", { module: form })).data?.dirty === false,
    `the canvas reports dirty ${JSON.stringify((await api.act("designerCanvas", { module: form })).data?.dirty)}`);

  // The 2026-08-27 half: the run wrote the FORM and nothing further. A red here on an old
  // build reads "the baselines differ", which is the save this pin exists to keep out.
  const fileAfter = JSON.stringify(await api.designerBaseline(form, project));
  check("and the SAVED baseline never felt the run - F5 does not write the file",
    fileAfter === fileBefore,
    fileAfter === fileBefore ? "byte-identical across the run"
      : `before ${fileBefore.length} chars, after ${fileAfter.length} chars, first diff at `
        + [...fileBefore].findIndex((ch, at) => fileAfter[at] !== ch));

  // And the form goes back where the rows below expect it. F5 moved the FORM (the apply is
  // the point of the first pin) while the file kept the old position, so this section's
  // business is to put both back: the document set to what the rows found, and one real save
  // to re-anchor the file - the reset the resize rows open with sets the document, and a
  // document set against a form that moved under it is dirty for ever.
  await api.act("designerSetMarkup", { module: form, markup: String(tabMarkup.data) });
  await api.command("save");
  await waitFor("the form back where the F5 rows found it", async () =>
    Math.abs(((await api.designer(form, project)).controls
      .find((c) => c.name === "RegionPick")?.left ?? 0) - 84) < 0.51, { budgetMs: 20000 });

  // ---- and the handles resize: the same commit, the same undo, the size clause too ----

  // From canonical, so every number below is exact rather than inherited from the rows above.
  await api.act("designerSetMarkup", { module: form, markup: await canonicalNow() });
  await waitFor("the document back at canonical before the resize rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });

  const tabText = async () => String((await api.act("designerMarkup", { module: form })).data);


  const grew = await api.act("designerResize",
    { module: form, control: "RegionPick", edge: "se", dx: 24, dy: 12 });
  check("a pull on the south-east handle resizes a control", grew.did === true, grew.detail);
  await waitFor("the size clause to follow the handle", async () =>
    sizedAt("RegionPick", 84, 38, 144, 32), { budgetMs: 15000 });
  check("120x20 becomes 144x32, and the position is untouched", true);

  await api.act("designerResize", { module: form, control: "RegionPick", edge: "nw", dx: 12, dy: 12 });
  await waitFor("the north-west pull", async () =>
    sizedAt("RegionPick", 96, 50, 132, 20), { budgetMs: 15000 });
  check("a north-west pull moves the origin and the extent together", true);

  const drawn = (await api.act("designerCanvas", { module: form })).data.controls
    .find((c) => c.name === "RegionPick");
  check("and the canvas draws the box it just wrote",
    Math.abs(Number(drawn?.width) - 132) < 0.51 && Math.abs(Number(drawn?.height) - 20) < 0.51,
    JSON.stringify(drawn));

  await api.act("designerResize", { module: form, control: "RegionPick", edge: "se", dx: -500, dy: -500 });
  await waitFor("the floored box", async () =>
    sizedAt("RegionPick", 96, 50, 4, 4), { budgetMs: 15000 });
  check("a pull past the far edge stops at the floor size rather than inverting the box", true);

  // Shift+arrow is the keyboard's resize, the native designer's own pairing with a bare arrow.
  await press("ArrowRight", "shiftKey: true,");
  await waitFor("the keyboard resize", async () =>
    sizedAt("RegionPick", 96, 50, 5, 4), { budgetMs: 15000 });
  check("Shift+arrow resizes by a point where a bare arrow moves", true);

  // The FORM's own frame resizes too - and its line takes a size and never a position.
  const formGrew = await api.act("designerResize", { module: form, edge: "se", dx: 20, dy: 12 });
  check("the form's own frame resizes by its handles", formGrew.did === true, formGrew.detail);
  await waitFor("the form's line to carry the new size", async () =>
    new RegExp(`<Form Name="${form}"[^>]*Width="380" Height="332"`).test(await tabText()), { budgetMs: 15000 });
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
  // SCROLLED TO FIRST, because the question is whether a press on the handle reaches the handle,
  // not whether the canvas happens to be scrolled where the last gesture left it. The row before
  // this one pulls the form's south-east corner, which leaves the canvas at the bottom of a tall
  // form; on a short tab NameBox is then off the top and elementFromPoint honestly answers
  // nothing. A hand scrolls to what it is reaching for.
  const handleCursor = await api.ask(`(() => { const h = ${inView(".dc-handle-se")}; if (!h) return "no handle"; h.scrollIntoView({ block: "center", inline: "center" }); const r = h.getBoundingClientRect(); const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2); return getComputedStyle(h).cursor + "|" + (hit === h ? "reachable" : "covered by " + (hit?.className || "nothing")); })()`);
  check("a handle wears its own resize cursor and the pointer can reach it",
    handleCursor === "nwse-resize|reachable", String(handleCursor));

  // ...and the resize that handle promises actually runs, over the neighbour and all.
  const overNeighbour = await api.act("designerResize",
    { module: form, control: "NameBox", edge: "se", dx: 6, dy: 0 });
  check("and the pull it promises runs, even where a neighbour overlaps the handle",
    overNeighbour.did === true, overNeighbour.detail);

  // ---- and Delete takes a control out, children and all ----

  await api.act("designerSetMarkup", { module: form, markup: await canonicalNow() });
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
    !/\bName="Options"/.test(afterDelete) && !/\bName="PickGround"/.test(afterDelete)
    && /<ToggleButton\b[^>]*\bName="HoldToggle"/.test(afterDelete),
    afterDelete.split(/\r?\n/).length + " line(s) left");
  check("the FORM still holds it until the save",
    (await api.designer(form, project)).controls.some((c) => c.name === "PickGround"));
  check("and the selection lands back on the form, where the panel follows it",
    ((await api.act("designerCanvas", { module: form })).data?.selected) === "");

  await press("z", "ctrlKey: true,");
  const blockBack = await waitFor("the block to come back", async () => {
    const text = await tabText();
    return /<Frame\b[^>]*\bName="Options"/.test(text) && /\bName="PickGround"/.test(text)
      ? "restored"
      : null;
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

  await api.act("designerSetMarkup", { module: form, markup: await canonicalNow() });
  await waitFor("the document back at canonical before the toolbox rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });

  // ---- a form's design rides SYNC as text ----
  //
  // A UserForm exported as code alone cannot be put back: Excel writes its controls into a binary
  // .frx and the text names it. The markup is the design as TEXT, so a form finally round-trips
  // through a folder the way a module does - and the file is the same projection the designer tab
  // edits, which is what keeps the two from drifting.
  const syncFolder = join(tmpdir(), `xlide-design-${process.pid}`);
  rmSync(syncFolder, { recursive: true, force: true });
  mkdirSync(syncFolder, { recursive: true });

  try {
    // ---- EXPORT SHIPS WHAT IS ON SCREEN (the owner, 2026-08-19) ----
    //
    // The document is the transaction log and the form only catches up on a save, so an export
    // that skipped the save shipped the LAST one - the same bug Run had (2026-08-16), now kept
    // out by the same rule: export applies-and-saves the dirty designers first, page-side for
    // the dialog and in the api route's pre-step for this door.
    // The rest of this section drives the shim alone, so no designer tab is open here; the
    // flush proof needs one to dirty.
    // Opened the way every other section opens it - no project spelling, so this is the SAME
    // view under the same document key, not an empty twin beside it.
    await api.pane("open", { module: form, face: "design" });
    // Until it HOLDS the form: a fresh view answers before its document arrives, and a drag
    // against an empty canvas refuses by name.
    await waitFor("the designer document to arrive on the canvas", async () =>
      ((await api.act("designerCanvas", { module: form })).data?.controls?.length ?? 0) > 0);

    const designLine = async (name) =>
      (await tabText()).split(/\r?\n/).find((line) => line.includes(`Name="${name}"`)) ?? "";
    const leftIn = (line) => Number(/\bLeft="(-?[\d.]+)"/.exec(line)?.[1] ?? NaN);

    // The probe control comes off the LIVE canvas, not a name written here: earlier sections
    // legitimately remove controls for good (PickAir went in the nested-remove rows), and a
    // hardcoded name is archaeology waiting to fail. Any drawn control whose line spells Left
    // serves - the row is about the FLUSH, not about a particular box.
    const drawn = ((await api.act("designerCanvas", { module: form })).data?.controls ?? [])
      .map((one) => one.name);
    let probe = null;
    for (const name of drawn) {
      if (Number.isFinite(leftIn(await designLine(name)))) {
        probe = name;
        break;
      }
    }
    if (probe === null) {
      throw new Error(`no drawn control spells Left in the document; canvas holds: ${drawn.join(", ")}`);
    }

    const probeLeft = leftIn(await designLine(probe));
    // alt held: a one-point drag under grid snap rounds back to where it started, and this row
    // needs a real unsaved delta, not a snapped-away one.
    const nudged = await api.act("designerDrag", { module: form, control: probe, dx: 1, dy: 0, alt: true });
    if (nudged.did !== true) {
      throw new Error(`the flush proof's nudge was refused: ${nudged.detail}`);
    }

    await waitFor("the unsaved nudge to land in the document", async () =>
      leftIn(await designLine(probe)) === probeLeft + 1);

    const flushedExport = await api.syncApply("export", { folder: syncFolder, select: "all" });
    const exportedProbe = readFileSync(join(syncFolder, `${form}.form`), "utf8")
      .split(/\r?\n/).find((line) => line.includes(`Name="${probe}"`)) ?? "";
    check("an export ships what is ON SCREEN: the unsaved nudge is in the files",
      (flushedExport.failed ?? []).length === 0 && leftIn(exportedProbe) === probeLeft + 1,
      `${probe}: document holds ${probeLeft + 1}; the exported .form holds ${leftIn(exportedProbe)}`);

    const idle = await api.act("designerSaveDirty");
    check("...because the export saved it first: nothing is left dirty",
      idle.did === true && /nothing was dirty/.test(idle.detail ?? ""), idle.detail);

    // The point goes back and is saved, so the section's own export below starts clean.
    await api.act("designerDrag", { module: form, control: probe, dx: -1, dy: 0, alt: true });
    await api.act("designerSaveDirty");
    await waitFor("the nudge to be undone and saved", async () =>
      leftIn(await designLine(probe)) === probeLeft);

    const exported = await api.syncApply("export", { folder: syncFolder, select: "all" });
    const written = readdirSync(syncFolder);
    check("an export writes the form's design beside its code",
      written.includes(`${form}.form`), written.join(", "));
    check("and nothing failed on the way", (exported.failed ?? []).length === 0,
      (exported.failed ?? []).join(", "));

    // The VBE's own exporter writes the binary sidecar, and the text names it: a form written by
    // splicing named a TEMPORARY path instead, so every export of a form produced a file whose
    // OleObjectBlob line changed and pointed at nothing (found 2026-08-16).
    //
    // THIS form's .frm, by exact name. The old find took anything starting `${form}.` - which
    // is the .form DESIGN file first in every directory listing - and on FormFixture it then
    // fell back onto the OTHER form's .frm and passed on that file's blob line. The gate's
    // DebugFixture run, with one form and no accident available, read the design file, found
    // no blob line, and died at the arrival rows below (2026-08-19, the v0.7.0 gate).
    const codeFile = written.find((one) => one.toLowerCase() === `${form.toLowerCase()}.frm`);
    const codeText = codeFile ? readFileSync(join(syncFolder, codeFile), "latin1") : "";
    const blob = /OleObjectBlob\s*=\s*"([^"]+)"/.exec(codeText)?.[1];
    check("the exported form names a sidecar that is actually there",
      blob !== undefined && written.includes(blob), `${blob ?? "no blob line"} of ${written.join(", ")}`);
    if (blob === undefined) {
      throw new Error(`${form}.frm carries no OleObjectBlob line, and the arrival rows below `
        + "clone that sidecar - there is nothing meaningful to run past this point.");
    }

    const design = readFileSync(join(syncFolder, `${form}.form`), "utf8");
    check("the design file holds the same projection the tab holds",
      design.startsWith(`<Form Name="${form}"`) && /<CommandButton Name="OkButton"/.test(design),
      design.split(String.fromCharCode(10))[0]);

    // Edited on disk the way a developer would in a pull request, then imported: the row applies
    // through the markup's own diff, so the FORM itself moves.
    writeFileSync(join(syncFolder, `${form}.form`),
      design.replace(
        /(<CommandButton Name="OkButton"[^>]*?) Left="\d+" Top="\d+"/,
        '$1 Left="40" Top="40"'), "utf8");

    const incoming = await api.syncPlan("import", { folder: syncFolder });
    const designRow = (incoming.items ?? []).find((row) => row.file === `${form}.form`);
    check("the import plan offers the design as its own row, ticked",
      designRow?.status === "will-update" && designRow?.checked === true, JSON.stringify(designRow ?? null));

    const applied = await api.syncApply("import", { folder: syncFolder, ids: [designRow.id] });
    check("applying it lands on the form itself",
      (applied.failed ?? []).length === 0
      && (applied.changed ?? []).some((one) => one.includes(form)),
      JSON.stringify({ changed: applied.changed, failed: applied.failed }));

    const moved = (await api.designer(form, project)).controls.find((one) => one.name === "OkButton");
    check("40,40 - the edit made in a text file reached the control",
      moved?.left === 40 && moved?.top === 40, JSON.stringify({ left: moved?.left, top: moved?.top }));

    // ---- LIVENESS BEYOND THE FUNNEL ----
    //
    // Every designer mutation this product makes re-projects its tab. A SYNC IMPORT does not: it
    // applies the markup straight at the form, which is a real edit from outside the funnel and
    // the same shape as one made in the native designer underneath. So the tab is stale, and the
    // thing that catches it is a cheap fingerprint - each control's name and bounds - asked on
    // the window events that already fire.
    const tabSaysNow = async () => String((await api.act("designerMarkup", { module: form })).data);
    const staleTab = /<CommandButton Name="OkButton"[^>]*Left="262" Top="250"/.test(await tabSaysNow());

    if (staleTab) {
      const caught = await api.designerEdit("liveness", { module: form });
      check("an edit from OUTSIDE the funnel is caught by the fingerprint",
        /re-projected/.test(caught.detail ?? "") && caught.detail.includes(form), caught.detail);
      await waitFor("the tab to show what the form actually holds", async () =>
        /<CommandButton Name="OkButton"[^>]*Left="40" Top="40"/.test(await tabSaysNow()), { budgetMs: 15000 });
      check("and the tab shows what the form holds rather than what it last drew", true);
    } else {
      check("an edit from OUTSIDE the funnel is caught by the fingerprint", true,
        "not exercisable: something re-projected the tab before the check ran");
    }

    // AND IT DOES NOT CRY WOLF. A second check finds the same form and says so, which is what
    // keeps this off the cost of a re-projection on every window that appears in the process.
    const quiet = await api.designerEdit("liveness", { module: form });
    check("a second check finds nothing, because nothing changed",
      /already showed/.test(quiet.detail ?? ""), quiet.detail);

    // ---- THE SAVED BASELINE, AND THE TWO IMPLEMENTATIONS OF IT ----
    //
    // The workbook's own storage records which properties are not the file format default, and
    // that walk is written TWICE on purpose: in Core, and in saved-design.mjs. Both were written
    // against [MS-OFORMS] rather than against each other, so the check that matters is pointing
    // them at one file - a misreading its author would make identically in both is exactly what
    // a single implementation cannot catch.
    //
    // It asks about THIS suite's own form, which is on disk by now: the suite calls the host's
    // save twice on its way here, and a save writes the VBA project. That matters because the
    // gate runs this suite against DebugFixture, where the FormFixture's own form is not present
    // at all - rows guarded on that would report not-exercisable for ever and check nothing.
    const baseline = await api.designerBaseline(form, project);
    check("the saved workbook answers a baseline for a form it has been saved with",
      baseline.saved === true && baseline.controls.length > 0,
      JSON.stringify({ saved: baseline.saved, workbook: baseline.workbook, rows: baseline.controls.length }));

    if (baseline.saved && baseline.workbook) {
      // Excel is holding the workbook open. The host reads it shared; this side uses plain
      // readFileSync, so a lock is a reason the comparison cannot run rather than a failure of
      // the thing being compared.
      const { readSavedDesign } = await import("./saved-design.mjs");
      let mine = null;
      let locked = null;
      try {
        mine = readSavedDesign(baseline.workbook).get(form) ?? new Map();
      } catch (why) {
        locked = why.message;
      }

      const theirs = new Map(baseline.controls.map((one) => [one.path, one.changed.join(",")]));
      const differ = mine === null ? [] : [...new Set([...mine.keys(), ...theirs.keys()])]
        .filter((path) => (mine.get(path) ?? []).join(",") !== (theirs.get(path) ?? ""));
      check("the host's reader and the harness's agree on every control",
        differ.length === 0,
        locked ? `not exercisable: Excel is holding the workbook (${locked})`
          : differ.slice(0, 4).map((path) =>
            `${path}: host [${theirs.get(path) ?? ""}] vs harness [${(mine.get(path) ?? []).join(",")}]`).join("; "));

      // Nothing structural: Size is MUST-be-1 in every mask and TabIndex is set on all but the
      // first control, so either one surviving means the filter stopped working - and a list that
      // always says the same thing says nothing.
      const structural = baseline.controls.filter((one) =>
        one.changed.includes("Size") || one.changed.includes("TabIndex"));
      check("nothing structural reaches the answer",
        structural.length === 0, JSON.stringify(structural.slice(0, 3)));

      // A SET BIT IS A CANDIDATE RATHER THAN A VERDICT, which is the one thing about this answer
      // a later reader must not forget: the mask measures against the FILE FORMAT default, so a
      // control nobody touched still lists what its KIND is born differing in. The suite sets
      // nothing but a Caption on its CheckBox, and the file still records its colours.
      const box = baseline.controls.find((one) => /Check/i.test(one.path));
      check("a set bit is a candidate rather than a verdict - an untouched CheckBox still lists its colours",
        box === undefined || box.changed.includes("BackColor"),
        JSON.stringify(box ?? "no CheckBox in this form"));
    }

    // THE NEVER-SAVED CASE IS NOT CHECKABLE HERE, and saying so is better than a row that looks
    // like it checks it. `saved: false` is what a form the file has never seen must answer -
    // an empty list would tell the projection that every control is untouched, which for a form
    // built five seconds ago is the opposite of the truth. But everything in this session has
    // been saved by now. The fallback is pinned in Core instead, where a workbook can simply not
    // exist.

    const noSuchBaseline = await api.designerBaseline("NoSuchForm", project)
      .catch((why) => ({ refusal: why.message }));
    check("a form that is not there is refused rather than answered empty",
      /not a UserForm/.test(noSuchBaseline.refusal ?? ""),
      JSON.stringify(noSuchBaseline).slice(0, 160));

    // ---- A THIRD-PARTY CONTROL, and what this machine says about one ----
    //
    // `Controls.Add` takes any ProgID and the answer depends on the MACHINE rather than on this
    // product, so the row asserts whichever truth it finds and both are worth having. A control
    // that arrives must walk under its own type rather than a guess; a control Office refuses
    // must be refused in words that say WHERE the refusal came from, because otherwise the
    // developer's next hour goes into the wrong place. On this machine, 2026-08-16, it is the
    // second: the Trust Center's ActiveX setting turns every non-MSForms control off, and
    // MSComctlLib.TreeCtrl.2, Shell.Explorer.2 and RefEdit.Ctrl are all registered and all
    // refused with TRUST_E_SUBJECT_NOT_TRUSTED.
    const foreign = await api.designerEdit("add", {
      module: form, project, type: "MSComctlLib.TreeCtrl.2", name: "Foreign",
      left: 8, top: 8, width: 60, height: 40,
    }).catch((why) => ({ error: why.message }));

    if (foreign.error) {
      check("a third-party control this machine will not create is refused in words that place the blame",
        /not registered on this machine|Trust Center/.test(foreign.error),
        String(foreign.error).slice(-150));
    } else {
      const row = (await api.designer(form, project)).controls.find((one) => one.name === foreign.name);
      check("a third-party control walks under its own type rather than a guess",
        row !== undefined && row.type !== "Control", JSON.stringify(row ?? null));
      await api.designerEdit("remove", { module: form, project, name: foreign.name }).catch(() => {});
    }

    // And a design whose form is not there is refused by name rather than guessed at.
    writeFileSync(join(syncFolder, "NoSuchForm.form"),
      `<Form Name="NoSuchForm" Width="100" Height="80" />\r\n`, "utf8");
    const orphan = (await api.syncPlan("import", { folder: syncFolder })).items
      .find((row) => row.file === "NoSuchForm.form");
    check("a design whose form is missing is skipped, saying what to do first",
      orphan?.status === "skipping-import" && /Add the form first/.test(orphan?.warning ?? ""),
      JSON.stringify(orphan ?? null));

    // ---- IMPORT CREATES A FORM, 2026-08-16 ----
    //
    // The last refusal standing between this product and a form that round-trips through source
    // control. The pair the VBE's own exporter writes is what its importer reads, so the whole
    // form comes back - controls, fonts, pictures - and the markup applies on top of it.
    //
    // Under a name the session has never seen, which is what a form arriving from someone else's
    // branch is. Renaming a form in source control means THREE edits, and the third one is the
    // one that is easy to miss: the `Begin` line names the form as well as `VB_Name` does, and
    // the importer refuses a pair where the two disagree ("Errors during load", measured
    // 2026-08-16 - which is also why a refusal now carries the editor's own log).
    const arrival = `Arrived${process.pid % 1000}`;
    const codeExt = codeFile.slice(codeFile.lastIndexOf("."));
    writeFileSync(join(syncFolder, `${arrival}${codeExt}`),
      codeText
        .replace(/^(Begin \{[0-9A-Fa-f-]+\}) \S+/m, `$1 ${arrival}`)
        .replace(/Attribute VB_Name = "[^"]+"/, `Attribute VB_Name = "${arrival}"`)
        .replace(/OleObjectBlob\s*=\s*"[^"]+"/, `OleObjectBlob   =   "${arrival}.frx"`),
      "latin1");
    writeFileSync(join(syncFolder, `${arrival}.frx`), readFileSync(join(syncFolder, blob)));
    writeFileSync(join(syncFolder, `${arrival}.form`),
      design.replace(new RegExp(`<Form Name="${form}"`), `<Form Name="${arrival}"`), "utf8");

    const arriving = await api.syncPlan("import", { folder: syncFolder });
    const createRow = (arriving.items ?? []).find((row) => row.file === `${arrival}${codeExt}`);
    check("a form's PAIR in the folder is a create rather than a refusal",
      createRow?.status === "will-create" && createRow?.checked === true, JSON.stringify(createRow ?? null));

    // THE FOOT-GUN GUARD. The sidecar carries every control and the markup's list is TOTAL, so
    // applying the text over a freshly imported binary can only take things away when the two
    // disagree. The row is offered and left for the developer to tick.
    const alongside = (arriving.items ?? []).find((row) => row.file === `${arrival}.form`);
    check("its design row rides along, offered but UNTICKED - the binary is authoritative",
      alongside?.status === "will-update" && alongside?.checked === false
      && /would be removed/.test(alongside?.warning ?? ""), JSON.stringify(alongside ?? null));

    const born = await api.syncApply("import", { folder: syncFolder, ids: [createRow.id] });
    check("and the apply makes the form", (born.failed ?? []).length === 0
      && (born.changed ?? []).includes(arrival), JSON.stringify(born));

    // WHOLE means every control the exported design named - compared against the file rather
    // than against the suite's own form, which the rows above have been adding controls to.
    // A TAB is a line of the design and is NOT a control - it holds nothing and has no geometry -
    // so the walk this compares against does not carry one, and neither does the expectation.
    // Read off the ELEMENTS the file declares: `<Kind Name="...">`. This matched
    // `Kind Name` at the head of a line under the old dialect and matches nothing at all under
    // the tagged one, so the expectation was an empty list and the row compared 19 against 0.
    const namedByFile = [...design.matchAll(/<([A-Za-z_]\w*)\b[^>]*\bName="([^"]+)"/g)]
      .filter((one) => one[1] !== "Tab" && one[1] !== "Form").map((one) => one[2]).sort();
    const wholeForm = await api.designer(arrival, project).catch((why) => ({ error: why.message }));
    const arrivedNames = (wholeForm.controls ?? []).map((one) => one.name).sort();
    check("it arrives WHOLE - every control the binary held, not an empty frame",
      namedByFile.length > 0 && namedByFile.every((name) => arrivedNames.includes(name)),
      `${arrivedNames.length} arrived for ${namedByFile.length} named`);
    check("...its pictures among them, which no text file could have carried",
      (wholeForm.controls ?? []).filter((one) => one.picture).length === 2,
      (wholeForm.controls ?? []).filter((one) => one.picture).map((one) => one.name).join(", "));

    const arrivedCode = await api.readModule(arrival, project).catch((why) => ({ error: why.message }));
    check("...and its code, which came in the same file",
      /Views_Change/.test(String(arrivedCode.text ?? arrivedCode.source ?? "")),
      JSON.stringify(arrivedCode).slice(0, 120));

    await api.component("remove", { name: arrival, project }).catch(() => {});

    // A form whose sidecar is not beside it cannot be created, and the refusal NAMES the file it
    // wanted. Its own name, never created, because a name that HAS been created reads as an
    // update. The built-in planner is the one that can say so before the apply; the shared
    // planner has never heard of a form's sidecar and fails at the apply instead.
    const lonelyName = `${arrival}Alone`;
    writeFileSync(join(syncFolder, `${lonelyName}${codeExt}`),
      codeText
        .replace(/^(Begin \{[0-9A-Fa-f-]+\}) \S+/m, `$1 ${lonelyName}`)
        .replace(/Attribute VB_Name = "[^"]+"/, `Attribute VB_Name = "${lonelyName}"`)
        .replace(/OleObjectBlob\s*=\s*"[^"]+"/, `OleObjectBlob   =   "${lonelyName}.frx"`),
      "latin1");

    const engineWas = (await api.ui()).settings?.syncEngine ?? "shared";
    await api.settings({ syncEngine: "builtIn" });
    const lonely = (await api.syncPlan("import", { folder: syncFolder })).items
      .find((row) => row.file === `${lonelyName}${codeExt}`);
    check("a form without its sidecar is refused by the name of the file it wants",
      lonely?.status === "skipping-import" && lonely?.checked === false
      && String(lonely.warning).includes(`${lonelyName}.frx`), JSON.stringify(lonely ?? null));
    await api.settings({ syncEngine: engineWas });
  } finally {
    rmSync(syncFolder, { recursive: true, force: true });
  }

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
    /<CommandButton Name="CommandButton1"[^>]*Left="230" Top="120" Width="72" Height="24"/.test(withDrop),
    withDrop.split(/\r?\n/).find((line) => /CommandButton1/.test(line)) ?? "no line");
  check("and the new control is SELECTED, the way one dropped from the native palette is",
    ((await api.act("designerCanvas", { module: form })).data?.selected) === "CommandButton1");
  check("the FORM does not have it until the save",
    !(await api.designer(form, project)).controls.some((c) => c.name === "CommandButton1"));

  /*
   * ---- WHERE A CONTAINER'S CLIENT BEGINS, which is not where its edge is drawn ----
   *
   * Two numbers per container and they are NOT the same number, which is the whole of the defect
   * this pins (2026-08-18). A Frame's rule sits 4.69pt below the control's top and its client
   * begins 6.05pt below it; a MultiPage's body edge sits 15.3pt below and its children begin
   * 17.22pt below. The canvas derived each pair from ONE constant - half the band for the rule,
   * the whole of it for the client - which ties them at a ratio the runtime does not keep. Every
   * control inside a Frame therefore painted 2.9pt low and every control on a Page 2.2pt high.
   *
   * designer-parity.mjs is what measured it, against the RUNNING form, and it went from seven
   * controls a point or more out to none. This row is the cheap guard that the constants behind
   * that do not drift back: it reads what the canvas actually lays out, in points, so a change
   * to either number fails here without needing a form to be launched and photographed.
   */
  const clientOrigins = await api.ask(`(() => {
    const view = document.querySelector('.designer-view[data-module=${JSON.stringify(form)}]');
    const PT = 4 / 3;
    const at = (element) => element.getBoundingClientRect();
    const frame = view.querySelector('.dc[data-control="Options"]');
    const rule = frame.querySelector('.dc-frame-rule');
    const client = frame.querySelector('.dc-frame-client');
    const multi = view.querySelector('.dc[data-control="Wizard"]');
    const body = view.querySelector('.dc-page-body');
    // A child is placed from its container's PADDING box, which starts below the border - so the
    // border is where a Page's children actually begin, and padding would have moved nothing.
    const inset = parseFloat(getComputedStyle(body).borderTopWidth) || 0;
    return JSON.stringify({
      frameRule: +((at(rule).top - at(frame).top) / PT).toFixed(2),
      frameClient: +((at(client).top - at(frame).top) / PT).toFixed(2),
      pageEdge: +((at(body).top - at(multi).top) / PT).toFixed(2),
      pageClient: +((at(body).top + inset - at(multi).top) / PT).toFixed(2),
    });
  })()`);
  {
    const seen = typeof clientOrigins === "string" ? JSON.parse(clientOrigins) : clientOrigins;
    const near = (a, b) => Math.abs(a - b) < 0.6;
    check("a Frame's rule and its CLIENT are two different distances, both as measured",
      near(seen.frameRule, 4.69) && near(seen.frameClient, 6.05), JSON.stringify(seen));
    check("and a MultiPage keeps the same pair apart - its body edge, then where children begin",
      near(seen.pageEdge, 15) && near(seen.pageClient, 17.22), JSON.stringify(seen));
    check("...and the client is BELOW the edge in both, which is the relation that was collapsed",
      seen.frameClient > seen.frameRule && seen.pageClient > seen.pageEdge, JSON.stringify(seen));
  }

  // Dropped INSIDE a Frame, the container under the pointer wins: the line nests under it and
  // its position is the frame's own, not the form's.
  const nestedDrop = await api.act("designerToolbox",
    { module: form, kind: "Label", left: 40, top: 150 });
  check("a drop over a Frame goes INTO the frame", nestedDrop.did === true, nestedDrop.detail);
  const withNested = await tabText();
  // Eight spaces of indent is the nesting - the Frame's own children sit one level in - and the
  // coordinates are the frame's, not the form's, which is what a drop INTO a container means.
  check("and its element is nested under it, placed in the frame's own coordinates",
    /^ {8}<Label\b[^>]*\bName="Label1"[^>]*\/>$/m.test(withNested)
    && (await placed("Label1")).width === 66 && (await placed("Label1")).height === 16,
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
    /<Form /.test(floor) && /<CommandButton Name="OkButton"/.test(floor), `${floor.length} char(s)`);

  await api.act("designerSetMarkup", { module: form, markup: await canonicalNow() });
  const restored = await waitFor("the document back at canonical", async () => {
    const read = (await api.act("designerCanvas", { module: form })).data;
    return read?.dirty === false ? read : null;
  }, { budgetMs: 15000 });
  check("and the unapplied moves let go, leaving the form as the plan draws it",
    restored.draft === false);
  await api.act("designerSelect", { module: form });

  // The stub gesture pulled the CODE tab active; the rename rows below need the face back.
  await api.act("activate", { module: form, face: "design" });

  // ---- COPY, PASTE, CUT AND DUPLICATE ----
  //
  // There was no clipboard on this canvas at all until 2026-08-18, so laying out six alike
  // controls was six trips to the toolbox. All four write the DOCUMENT and none of them touch the
  // form until Ctrl+S, which is what makes a mistaken paste one Ctrl+Z away.
  // The document as THIS section found it, so the restore below puts back what this section
  // changed rather than what the suite looked like several sections ago - by here the form has
  // legitimately grown and lost controls, and a stale snapshot never goes clean again.
  const beforeClipboard = await tabText();

  await api.act("designerSelect", { module: form, control: "OkButton" });
  const copied = await api.act("designerClipboard", { module: form, how: "copy" });
  check("a control can be copied", copied.did === true && /OkButton/.test(copied.detail ?? ""),
    copied.detail);

  const pasted = await api.act("designerClipboard", { module: form, how: "paste" });
  const pastedName = /pasted (\w+)/.exec(pasted.detail ?? "")?.[1] ?? "";
  check("and pasted under a name the document did not already hold",
    pasted.did === true && pastedName !== "" && pastedName !== "OkButton", pasted.detail);
  // Offset, so a copy is visibly a copy rather than hiding exactly under its original.
  const original = await placed("OkButton");
  const copy = await placed(pastedName);
  check("...offset from the original, and carrying its properties",
    copy.left === original.left + 6 && copy.top === original.top + 6
    && /PicturePosition="1"/.test(copy.line),
    `${await placedAt("OkButton")} -> ${await placedAt(pastedName)}`);
  check("and the FORM does not have it until the save",
    !(await api.designer(form, project)).controls.some((one) => one.name === pastedName));

  // A CONTAINER carries its children, and every one of them needs a free name too - two controls
  // of one name is a document the apply refuses.
  await api.act("designerSelect", { module: form, control: "Options" });
  await api.act("designerClipboard", { module: form, how: "copy" });
  const frame = await api.act("designerClipboard", { module: form, how: "paste" });
  const frameName = /pasted (\w+)/.exec(frame.detail ?? "")?.[1] ?? "";
  const withFrame = await tabText();
  const names = [...withFrame.matchAll(/\bName="([^"]*)"/g)].map((one) => one[1]);
  check("a container is copied with its children, all of them renamed",
    frame.did === true && names.length === new Set(names).size
    && new RegExp(`<Frame\\b[^>]*\\bName="${frameName}"[^>]*>`).test(withFrame),
    `${frame.detail}; ${names.length} names, ${new Set(names).size} distinct`);

  // COPYING A CONTAINER LEAVES IT SELECTED, and "paste into the selected container" then put the
  // copy inside the original - a Frame growing a Frame. A paste of the thing that IS selected
  // lands beside it; the container rule is for pasting something else IN.
  const pastedFrameLine = withFrame.split(/\r?\n/)
    .find((line) => new RegExp(`\\bName="${frameName}"`).test(line)) ?? "";
  check("...and a container pasted from its own selection lands BESIDE it, not inside it",
    pastedFrameLine.length - pastedFrameLine.trimStart().length === 4,
    JSON.stringify(pastedFrameLine));

  const cut = await api.act("designerClipboard", { module: form, how: "cut" });
  check("cut takes it back out", cut.did === true && !(await tabText()).includes(frameName), cut.detail);

  // A Page belongs to a MultiPage and nowhere else. The parser would refuse the document and the
  // canvas would stop previewing, so the refusal is said where it can name the reason.
  await api.act("designerSelect", { module: form, control: "Page1" });
  await api.act("designerClipboard", { module: form, how: "copy" });
  await api.act("designerSelect", { module: form });
  const pageRefused = await api.act("designerClipboard", { module: form, how: "paste" });
  check("a Page pasted onto the form is refused, saying what it needs",
    pageRefused.did === false && /only be pasted into a MultiPage/.test(pageRefused.detail ?? ""),
    pageRefused.detail);

  // A HAND-WRITTEN CAPTION CANNOT BE SHREDDED BY THE RENAME. The document is editable text and
  // attributes may be typed in any order, so `<Label Caption="Call me Name=" Name="ViewNote">`
  // is legal - and the paste's rename used to take the leftmost `Name="..."` ANYWHERE in the
  // line, which put the new name inside that caption and mangled the text (the 2026-08-19
  // hunt, round three). Anchored to the header's own position, the caption rides untouched;
  // a line reordered away from the printer's shape keeps its name, and the duplicate that
  // makes is a refusal that NAMES the problem rather than a shredded document.
  const trapped = beforeClipboard.replace(/\bName="ViewNote"/, 'Caption="Call me Name=" Name="ViewNote"');
  await api.act("designerSetMarkup", { module: form, markup: trapped });
  await api.act("designerSelect", { module: form, control: "ViewNote" });
  await api.act("designerClipboard", { module: form, how: "copy" });
  const tricky = await api.act("designerClipboard", { module: form, how: "paste" });
  const captionsAfter = ((await tabText()).match(/Caption="Call me Name="/g) ?? []).length;
  check("a caption that spells Name= is never rewritten by the paste's rename",
    tricky.did === true && captionsAfter === 2,
    `${tricky.detail}; intact captions: ${captionsAfter} (2 means both copies kept their text)`);

  await api.act("designerSetMarkup", { module: form, markup: beforeClipboard });
  await waitFor("the document back where the clipboard rows found it", async () =>
    (await tabText()) === beforeClipboard, { budgetMs: 15000 });

  // ---- a drag carries a control OUT of its container: reparenting ----
  //
  // Until this, a drag clamped at the container's edge and moving a control from a Frame to the
  // form could only be done by editing the markup by hand. The gesture needs only a short travel
  // - what matters is which container the POINTER ends over - so the row drops a control near the
  // frame's bottom edge and pushes it across, which fits whatever canvas the window leaves.

  // Dropped as low in the frame as it will go - the drop clamps inside the container, so this is
  // the deepest a 16-point label fits. That leaves its centre a few points above the boundary and
  // a ten-point drag crosses it: what a reparent needs is the crossing, and a short travel is
  // what fits the canvas the gate's window leaves (about thirty points).
  //
  const nearEdge = await api.act("designerToolbox",
    { module: form, kind: "Label", left: 40, top: 166 });
  // Eight spaces of indent is inside the Frame, four is the form's own level - the nesting is what
  // says which container holds it, and in the tagged dialect that is the element's indent.
  check("a label dropped near the frame's bottom edge lands inside the frame",
    nearEdge.did === true && /^ {8}<Label\b[^>]*\bName="Label\d+"/m.test(await tabText()),
    nearEdge.detail);
  const carried = /^ {8}<Label\b[^>]*\bName="(Label\d+)"/m.exec(await tabText())?.[1] ?? "Label1";

  const outward = await api.act("designerDrag", { module: form, control: carried, dx: 0, dy: 10 });
  check("dragging it past that edge is accepted", outward.did === true, outward.detail);
  const reparented = await waitFor("the control's element to leave the frame's block", async () =>
    new RegExp(`^ {4}<Label\\b[^>]*\\bName="${carried}"`, "m").test(await tabText())
      ? await tabText()
      : null,
  { budgetMs: 15000 });
  check("the control's whole element moves to the form's own level, in the form's coordinates",
    new RegExp(`^ {4}<Label\\b[^>]*\\bName="${carried}"`, "m").test(reparented)
    && (await placed(carried)).top >= 160 && (await placed(carried)).top < 200,
    reparented.split(/\r?\n/).find((line) => line.includes(carried)) ?? "no line");

  // And back in, which is the same gesture the other way: the pointer decides, not the distance.
  //
  // TWENTY POINTS, not ten. The outward drag leaves the label just below the frame's lower edge
  // but still within its span, so a ten-point return put the pointer back inside the frame's
  // RECTANGLE without reaching its client - and the control stayed on the form. Measured by hand
  // 2026-08-17: -10 does not reparent, -20 does, landing it at the frame's own top. A row that
  // depends on the smallest crossing that works is a row that breaks whenever the drop clamps
  // half a point differently.
  const inward = await api.act("designerDrag", { module: form, control: carried, dx: 0, dy: -20 });
  check("and a drag back over the frame puts it inside again", inward.did === true, inward.detail);
  await waitFor("the control's element to nest under the frame again", async () =>
    new RegExp(`^ {8}<Label\\b[^>]*\\bName="${carried}"`, "m").test(await tabText()),
  { budgetMs: 15000 });
  check("nested again, indented under the Frame that now holds it", true);

  // The FORM is untouched by all of it until the save, like every other canvas gesture.
  check("the form has never heard of the control the canvas has been carrying",
    !(await api.designer(form, project)).controls.some((one) => one.name === carried));

  await api.act("designerDelete", { module: form, control: carried });
  await waitFor("the probe control to leave the document", async () =>
    !(await tabText()).includes(carried), { budgetMs: 15000 });
  check("and Delete takes the probe control back out of the document", true);

  // ---- a GROUP: gathered by ctrl+click or a marquee, moved and onTheLeft up as one ----

  // The four by NAME off each element, the same way `placed` reads one: the tagged dialect writes
  // them as attributes in whatever order the printer chose, so `at x,y size WxH` finds nothing.
  const placesOf = async (...names) => {
    const text = await tabText();
    return names.map((name) => {
      const line = text.split(/\r?\n/).find((one) => new RegExp(`\\bName="${name}"`).test(one)) ?? "";
      const number = (attribute) => {
        const found = new RegExp(`\\b${attribute}="(-?[\\d.]+)"`).exec(line);
        return found === null ? -1 : Number(found[1]);
      };

      return {
        name,
        left: number("Left"), top: number("Top"),
        width: number("Width"), height: number("Height"),
      };
    });
  };

  const group = ["NameLabel", "Taxable", "HoldToggle"];
  const gather = async () => {
    await api.act("designerSelect", { module: form, control: group[0] });
    for (const also of group.slice(1)) {
      await api.act("designerSelect", { module: form, control: also, extend: 1 });
    }
    return (await api.act("designerCanvas", { module: form })).data?.group ?? [];
  };

  check("ctrl+click gathers a group, the first one picked staying the anchor",
    JSON.stringify(await gather()) === JSON.stringify(group), JSON.stringify(await gather()));

  // A group is one container's business: controls measured from different origins cannot be
  // moved together or onTheLeft up with each other in any way that means something on screen.
  await api.act("designerSelect", { module: form, control: "PickGround", extend: 1 });
  check("a control in another container starts a new selection rather than a nonsense group",
    JSON.stringify((await api.act("designerCanvas", { module: form })).data?.group) === '["PickGround"]',
    JSON.stringify((await api.act("designerCanvas", { module: form })).data?.group));

  await gather();
  const wasAt = await placesOf(...group);
  const groupDrag = await api.act("designerDrag", { module: form, control: group[0], dx: 12, dy: 0 });
  check("a drag on any of them moves the whole group", groupDrag.did === true, groupDrag.detail);
  const nowAt = await waitFor("the group to move in the document", async () => {
    const now = await placesOf(...group);
    return now[0].left !== wasAt[0].left ? now : null;
  }, { budgetMs: 15000 });
  const deltas = nowAt.map((one, at) => `${one.left - wasAt[at].left},${one.top - wasAt[at].top}`);
  check("all of them by the SAME delta, whatever the snap made of it",
    new Set(deltas).size === 1, deltas.join(" | "));

  await press("z", "ctrlKey: true,");
  await waitFor("one undo to give the whole group move back", async () =>
    (await placesOf(group[0]))[0].left === wasAt[0].left, { budgetMs: 15000 });
  check("and one Ctrl+Z gives the whole move back - a group gesture is one undo step", true);

  // Align: the ANCHOR is the reference and never moves, which is the native rule.
  await gather();
  const aligned = await api.act("designerArrange", { module: form, how: "left" });
  check("Align Left lines the group up on the anchor", aligned.did === true, aligned.detail);
  const onTheLeft = await placesOf(...group);
  check("every left is the anchor's, and the anchor itself has not moved",
    onTheLeft.every((one) => one.left === onTheLeft[0].left) && onTheLeft[0].left === wasAt[0].left,
    onTheLeft.map((one) => `${one.name} ${one.left}`).join(" | "));

  const samed = await api.act("designerArrange", { module: form, how: "width" });
  check("Same Width sizes them to the anchor", samed.did === true, samed.detail);
  const widths = await placesOf(...group);
  check("every width is the anchor's",
    widths.every((one) => one.width === widths[0].width),
    widths.map((one) => `${one.name} ${one.width}`).join(" | "));

  const spacing = await api.act("designerArrange", { module: form, how: "down" });
  check("Space Down spreads them evenly", spacing.did === true, spacing.detail);
  const spaced = (await placesOf(...group)).sort((a, b) => a.top - b.top);
  const gaps = spaced.slice(1).map((one, at) => one.top - (spaced[at].top + spaced[at].height));
  check("the GAPS between them are equal, and the two on the ends have not moved",
    Math.max(...gaps) - Math.min(...gaps) <= 1, gaps.join(" | "));

  // The menu these live on: the canvas's own, since the product has no menu bar and the editor's
  // Format menu would act on the native designer's selection rather than on ours.
  const arrangeMenu = await api.ask(`(() => {
    const el = ${inViewAll(".dc")}.find((one) => one.dataset.control === "${group[0]}");
    el.scrollIntoView({ block: "nearest", inline: "nearest" });
    const box = el.getBoundingClientRect();
    el.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true,
      clientX: Math.round(box.left + 4), clientY: Math.round(box.top + 4) }));
    return [...document.querySelectorAll(".menu-dropdown .menu-item")].map((one) => (one.textContent ?? "").trim()).join(" | ");
  })()`);
  check("a right-click on a group offers the arrange vocabulary",
    /Align Left/.test(String(arrangeMenu)) && /Space Down/.test(String(arrangeMenu)),
    String(arrangeMenu));
  await api.act("key", { code: "Escape" });

  // The marquee: what it TOUCHES, in that container, which is MSForms' own rule.
  const banded = await api.act("designerMarquee",
    { module: form, left: 4, top: 4, right: 210, bottom: 30 });
  check("a rubber band over the form's ground catches what it touches",
    banded.did === true && /NameLabel/.test(banded.detail ?? ""), banded.detail);
  const inTheBand = (await api.act("designerCanvas", { module: form })).data?.group ?? [];
  check("and the group it made holds only controls the band actually crossed",
    inTheBand.length >= 2 && inTheBand.includes("NameLabel") && !inTheBand.includes("HoldToggle"),
    inTheBand.join(", "));

  // Delete takes the whole group, and one undo brings all of it back.
  await gather();
  await press("Delete");
  const groupGone = await waitFor("the group to leave the document", async () => {
    const text = await tabText();
    return group.every((one) => !text.includes(one)) ? text : null;
  }, { budgetMs: 15000 });
  check("Delete takes every selected control out at once", groupGone.length > 0);
  await press("z", "ctrlKey: true,");
  await waitFor("one undo to bring the whole group back", async () => {
    const text = await tabText();
    return group.every((one) => text.includes(one));
  }, { budgetMs: 15000 });
  check("and one Ctrl+Z brings all of them back", true);

  await api.act("designerSetMarkup", { module: form, markup: await canonicalNow() });
  await waitFor("the document back at canonical after the group rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });
  await api.act("designerSelect", { module: form });

  // ---- depth: Bring to Front and Send to Back, proved on the RUNNING form ----
  //
  // The one canvas gesture that writes the model rather than the document, because MSForms'
  // Controls collection is not in z-order and does not move when ZOrder is called - measured, and
  // the reason this row photographs a running form instead of reading the walk back. Two
  // overlapping opaque labels, and the count of red against blue pixels says which is on top.
  const paint = async () => {
    // The same shape the Run row above uses: post the command, wait for the WINDOW to stand, and
    // only then photograph it. Capturing on a timer photographs whatever is there, which for a
    // form that has not opened yet is a picture with no form in it.
    await api.command("run");
    await waitFor("the form to stand for its picture", async () =>
      ((await api.userforms()).forms ?? []).some((title) => title.includes("Quarter Entry")),
    { budgetMs: 20000 });
    // Waited for the PICTURE, not just the window: a form that is up has not necessarily painted,
    // and a capture taken in that gap answers a rectangle with none of the form's own colours in
    // it (measured - the row read 0 red and 0 blue while the labels were plainly there).
    const shot = await waitFor("the form to have painted its colours", async () => {
      const image = Buffer.from(await api.capture("form", "Quarter Entry"));
      const from = image.readUInt32LE(10);
      for (let px = from; px + 3 < image.length; px += 4) {
        const [b, g, r] = [image[px], image[px + 1], image[px + 2]];
        if ((r > 200 && g < 60 && b < 60) || (b > 200 && r < 60 && g < 60)) {
          return image;
        }
      }

      return null;
    }, { budgetMs: 15000 });
    await api.userforms("close", "Quarter Entry");
    await waitFor("the form to close again", async () =>
      ((await api.userforms()).forms ?? []).length === 0, { budgetMs: 20000 });
    await waitFor("design mode to return", async () =>
      (await api.state()).debugMode === "design", { budgetMs: 20000 });

    // Bottom-up 32-bit BGRA, the same shape designer-parity.mjs decodes.
    const at = shot.readUInt32LE(10);
    let red = 0;
    let blue = 0;
    for (let px = at; px + 3 < shot.length; px += 4) {
      const [b, g, r] = [shot[px], shot[px + 1], shot[px + 2]];
      if (r > 200 && g < 60 && b < 60) { red += 1; } else if (b > 200 && r < 60 && g < 60) { blue += 1; }
    }

    return { red, blue };
  };

  for (const [name, left, top, colour] of [["ZRed", 210, 60, "255"], ["ZBlue", 230, 75, "16711680"]]) {
    await api.designerEdit("add", { module: form, project, type: "label", name, left, top, width: 90, height: 50 });
    await api.designerEdit("set", { module: form, project, name, property: "BackColor", value: colour });
    await api.designerEdit("set", { module: form, project, name, property: "BackStyle", value: "1" });
  }

  const overlapped = await paint();
  check("two overlapping labels, and the later one is on top to start with",
    overlapped.blue > overlapped.red && overlapped.red > 0, JSON.stringify(overlapped));

  /*
   * A CONTEXT MENU GOES WITH THE VIEW IT IS ABOUT.
   *
   * The owner photographed this canvas menu - Bring to Front, Send to Back, Center in Container,
   * Size to Grid, Tab Order - standing over a standard CODE module, where none of it means
   * anything. It cannot be opened there; it survived the tab switching beneath it, and in that
   * photograph it had survived the form being deleted as well. Every item acts on the designer's
   * remembered selection, so Delete on a stale menu deletes a control from a form nobody is
   * looking at.
   *
   * It was also corrupting the row below: the ZOrder check photographs a running form, and a menu
   * standing over the canvas made both photographs identical, so the row failed deterministically
   * and nothing said why.
   */
  const menuNow = () => api.ask(`(() => {
    const menu = document.querySelector(".menu-dropdown");
    return menu
      ? { open: true, items: [...menu.querySelectorAll(".menu-item")].map((o) => (o.textContent || "").trim()) }
      : { open: false, items: [] };
  })()`);

  const rightClicked = await api.ask(`(() => {
    const control = document.querySelector(".dc[data-control]");
    if (!control) { return "no control on the canvas"; }
    const box = control.getBoundingClientRect();
    const x = Math.round(box.left + box.width / 2);
    const y = Math.round(box.top + box.height / 2);
    control.dispatchEvent(new PointerEvent("pointerdown", {
      bubbles: true, cancelable: true, pointerId: 1, isPrimary: true, button: 2, buttons: 2,
      clientX: x, clientY: y,
    }));
    control.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: x, clientY: y }));
    return "right-clicked " + control.dataset.control;
  })()`);
  const onCanvas = await menuNow();
  check("the canvas offers its own menu on a right-click", onCanvas?.open === true,
    `${rightClicked}: ${JSON.stringify(onCanvas?.items ?? []).slice(0, 90)}`);

  await api.pane("open", { module: "ThisWorkbook", project });
  await new Promise((settle) => setTimeout(settle, 1200));
  const afterSwitch = await menuNow();
  check("and it goes when the view under it does - a stale one acts on a form nobody is showing",
    afterSwitch?.open === false, JSON.stringify(afterSwitch?.items ?? []).slice(0, 90));

  await api.pane("open", { module: form, project, face: "design" });
  await new Promise((settle) => setTimeout(settle, 1200));

  const raised = await api.act("designerZOrder", { module: form, control: "ZRed", front: 1 });
  check("Bring to Front is accepted", raised.did === true, raised.detail);
  check("and it leaves the DOCUMENT alone - depth is the one gesture the dialect cannot say",
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false);

  // The counts SWAP, within a few percent of each other rather than within a fixed number of
  // pixels. A photograph of a running form is not pixel-identical twice - the labels antialias
  // against whatever is behind them - and the absolute tolerance this used to carry was tuned
  // when the labels were half the size they are now: 272 pixels out of 18,000 failed a rule
  // written for 200 out of 8,000 (2026-08-16). What the row is about is which one is on top.
  const swapped = await paint();
  const withinAFewPercent = (a, b) => Math.abs(a - b) <= Math.max(a, b) * 0.05;
  check("the running form draws them the other way round - ZOrder reached the model",
    swapped.red > swapped.blue
    && withinAFewPercent(swapped.red, overlapped.blue)
    && withinAFewPercent(swapped.blue, overlapped.red),
    `${JSON.stringify(overlapped)} became ${JSON.stringify(swapped)}`);

  const depthMenu = await api.ask(`(() => {
    const el = ${inViewAll(".dc")}.find((one) => one.dataset.control === "ZRed");
    if (!el) { return "no ZRed on the canvas"; }
    el.scrollIntoView({ block: "nearest", inline: "nearest" });
    const box = el.getBoundingClientRect();
    el.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true,
      clientX: Math.round(box.left + 4), clientY: Math.round(box.top + 4) }));
    return [...document.querySelectorAll(".menu-dropdown .menu-item")].map((one) => (one.textContent ?? "").trim()).join(" | ");
  })()`);
  // The clipboard four joined this menu on 2026-08-18 and belong on it for ONE control as much
  // as for a group - it is a copy of this that a developer wants, and Delete had been on the
  // keyboard alone. What still needs a group is align, distribute and same-size, and their
  // absence here is what this row is really pinning.
  // The clipboard four and the two centring gestures joined this menu on 2026-08-18 and belong
  // on it for ONE control as much as for a group: it is a copy of THIS that a developer wants,
  // Delete had been on the keyboard alone, and centring a single control in its container is the
  // case those gestures are for. What still needs a group is align, distribute and same-size, and
  // their absence here is what this row is really pinning.
  check("one control's menu offers depth, the clipboard and shaping, and nothing that needs a group",
    String(depthMenu) === "Bring to Front | Send to Back | Cut | Copy | Paste | Duplicate"
      + " | Delete | Center in Container (across) | Center in Container (down) | Size to Fit"
      + " | Size to Grid | Tab Order...", String(depthMenu));
  await api.act("key", { code: "Escape" });

  await api.designerEdit("remove", { module: form, project, name: "ZRed" });
  await api.designerEdit("remove", { module: form, project, name: "ZBlue" });
  await waitFor("the two probe labels to leave the form", async () =>
    !(await api.designer(form, project)).controls.some((one) => one.name.startsWith("Z")),
  { budgetMs: 15000 });
  check("and the probe labels are off the form again", true);

  // ---- tab order: the dialog the native View menu keeps, on our own surface ----

  await api.act("designerSelect", { module: form });
  const orderOpened = await api.act("designerTabOrder", { module: form, open: 1 });
  check("the tab-order dialog opens for the form's own controls",
    orderOpened.did === true, orderOpened.detail);

  const inOrder = (await api.act("designerTabOrder", { module: form })).detail ?? "";
  const walkRows = (await api.designer(form, project)).controls.filter((one) => one.parent === form);
  const stops = walkRows.filter((one) => typeof one.tabIndex === "number");
  check("it lists them in TAB order, which is not the order the walk reads them in",
    inOrder.split(", ").length === stops.length
    && inOrder.split(", ").join() === [...stops]
      .sort((a, b) => a.tabIndex - b.tabIndex).map((one) => one.name).join(),
    inOrder);
  check("and a control with no tab stop of its own is not in the list at all",
    walkRows.some((one) => one.tabIndex === null) && !inOrder.includes("Badge"),
    walkRows.filter((one) => one.tabIndex === null).map((one) => one.name).join(", ") || "every control has one");

  const wereAt = Object.fromEntries(stops.map((one) => [one.name, one.tabIndex]));
  const lastStop = [...stops].sort((a, b) => a.tabIndex - b.tabIndex).at(-1);
  const aboveIt = [...stops].sort((a, b) => a.tabIndex - b.tabIndex).at(-2);
  const movedUp = await api.act("designerTabOrder", { module: form, control: lastStop.name, move: "up" });
  check("a Move Up is accepted", movedUp.did === true, movedUp.detail);
  await waitFor("the form to carry the new tab order", async () =>
    ((await api.designer(form, project)).controls.find((one) => one.name === lastStop.name)?.tabIndex)
      === wereAt[aboveIt.name], { budgetMs: 15000 });
  const areAt = (await api.designer(form, project)).controls;
  check("the control took the index above it, and MSForms pushed the other one down",
    areAt.find((one) => one.name === lastStop.name)?.tabIndex === wereAt[aboveIt.name]
    && areAt.find((one) => one.name === aboveIt.name)?.tabIndex === wereAt[lastStop.name],
    `${lastStop.name} ${wereAt[lastStop.name]}->${areAt.find((one) => one.name === lastStop.name)?.tabIndex}, `
    + `${aboveIt.name} ${wereAt[aboveIt.name]}->${areAt.find((one) => one.name === aboveIt.name)?.tabIndex}`);

  await api.act("designerTabOrder", { module: form, control: lastStop.name, move: "down" });
  await waitFor("the order to go back the way it was", async () =>
    ((await api.designer(form, project)).controls.find((one) => one.name === lastStop.name)?.tabIndex)
      === wereAt[lastStop.name], { budgetMs: 15000 });
  check("and a Move Down puts it back - the write is one TabIndex either way", true);
  await api.act("designerTabOrder", { module: form, close: 1 });

  // ONE container at a time, because tab order is per container in MSForms: a Frame's children
  // have their own 0..n and Tab walks into the frame and out again.
  // The document back to the form's own text first: the rows above have carried controls between
  // containers, and this one is about which container a control belongs to. Waited for rather
  // than assumed after that, because the dialog reads the selection at the moment it opens and a
  // projection landing in between - or a hand clicking in the same live session, which is what
  // this suite runs in - moves it.
  // The FORM's own text as it stands now, not the projection captured at the top of the run: the
  // rows between have moved controls and saved them, so the old canonical is a different form.
  await api.act("designerSetMarkup", { module: form, markup: await api.designerMarkup(form, project) });
  await waitFor("the document back at canonical before the container's own order", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });
  // A control of the FRAME as the canvas has it now, rather than a name written here: the rows
  // above delete and restore controls, and a name that is gone selects nothing while the act
  // cheerfully reports otherwise (which is fixed too - designerSelect refuses a name the canvas
  // does not draw).
  const framed = (await api.designerMarkup(form, project)).split(/\r?\n/)
    .filter((line) => /^ {8}<OptionButton\b/.test(line))
    .map((line) => /\bName="([^"]+)"/.exec(line)?.[1] ?? "")
    .filter((name) => name.length > 0);
  await api.act("designerSelect", { module: form, control: framed[0] });
  await waitFor("the canvas to hold the selection the dialog will read", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.selected) === framed[0],
  { budgetMs: 15000 });
  const frameOrder = await api.act("designerTabOrder", { module: form, open: 1 });
  const inFrame = (await api.act("designerTabOrder", { module: form })).detail ?? "";
  check("a control inside a Frame opens the FRAME's own order, not the form's",
    inFrame === framed.join(", "), `${frameOrder.detail} -> ${inFrame}`);
  await api.act("designerTabOrder", { module: form, close: 1 });
  await api.act("designerSelect", { module: form });

  // ---- containers: a page opens on the canvas, and the strip's menu adds and removes one ----
  //
  // Until this, a MultiPage drew its first page and nothing reached the second: its controls
  // were in the document and invisible, and a drop always landed on page one whatever the tabs
  // said. Which page is OPEN is view state deliberately - see the field's own note - so the
  // rows below check that opening one changes the picture and NOT the document.

  const canvasNow = async () => (await api.act("designerCanvas", { module: form })).data;
  const containerNow = async (name) =>
    (await canvasNow()).containers.find((one) => one.name === name) ?? null;
  const isDrawn = async (name) => (await canvasNow()).controls.some((one) => one.name === name);

  const bothStrips = (await canvasNow()).containers;
  check("both tabbed containers draw their tabs, the first one open",
    bothStrips.length === 2
    && bothStrips.every((one) => one.tabs.length === 2 && one.open === 0),
    JSON.stringify(bothStrips));

  const openedTwo = await api.act("designerOpenTab", { module: form, container: "Wizard", tab: "Page2" });
  check("a click on a tab opens that page", openedTwo.did === true, openedTwo.detail);
  check("the canvas draws the open page and not the one behind it",
    (await containerNow("Wizard"))?.page === "Page2" && !(await isDrawn("Agree")),
    JSON.stringify(await containerNow("Wizard")));

  const onPage = await canvasNow();
  check("and the PAGE is selected, with the markup caret on its own line - the native gesture",
    onPage.selected === "Page2" && /^\s+<Page\b[^>]*\bName="Page2"/.test((await tabText()).split(/\r?\n/)[onPage.markupLine - 1] ?? ""),
    `${onPage.selected} at line ${onPage.markupLine}`);
  check("opening a page leaves the DOCUMENT alone - looking is not an edit",
    onPage.dirty === false);

  const pagePanel = await waitFor("the panel to follow the page", async () => {
    const shown = (await api.ui()).properties;
    return shown?.component === "Page2" ? shown : null;
  }, { budgetMs: 15000 });
  check("the Properties panel targets the page, as the native designer's click does",
    pagePanel.kind === "Page" && pagePanel.rows.some((row) => row.name === "Caption"),
    `${pagePanel.kind}: ${pagePanel.rows.map((row) => row.name).join(", ")}`);

  // The other direction: a caret in page one's block opens page one. Without this a click on a
  // line inside a hidden page selected nothing at all, because the canvas had no element for it.
  const agreeLine = (await tabText()).split(/\r?\n/).findIndex((line) => /<CheckBox\b[^>]*\bName="Agree"/.test(line)) + 1;
  await api.act("designerCaret", { module: form, line: agreeLine });
  const backToOne = await canvasNow();
  check("a caret on a control inside a hidden page opens that page and selects it",
    backToOne.selected === "Agree" && (await isDrawn("Agree"))
    && (await containerNow("Wizard"))?.page === "Page1",
    `${backToOne.selected}, showing ${(await containerNow("Wizard"))?.page}`);

  // AND SO DOES SELECTING IT BY NAME, which is the same link and was refused for a while: the
  // act asked the CANVAS whether the control was there, and a control on a page that is not open
  // is real and undrawn (found in the 2026-08-16 hunt). The refusal it was meant for - a name
  // the form does not hold at all - still stands, and now asks the projection instead.
  await api.act("designerOpenTab", { module: form, container: "Wizard", tab: 2 });
  const pickedByName = await api.act("designerSelect", { module: form, control: "Agree" });
  check("selecting a control on a page that is NOT OPEN opens that page too",
    pickedByName.did === true && (await canvasNow()).selected === "Agree"
    && (await containerNow("Wizard"))?.page === "Page1",
    `${pickedByName.detail}, showing ${(await containerNow("Wizard"))?.page}`);

  const nobody = await api.act("designerSelect", { module: form, control: "NoSuchControl" });
  check("...while a name the form does not hold is still refused",
    nobody.did === false && /not on this form/.test(nobody.detail ?? ""), nobody.detail);

  // A drop lands on the page the developer is LOOKING at, which is the whole point of opening it.
  await api.act("designerOpenTab", { module: form, container: "Wizard", tab: "Page2" });
  const ontoPage = await api.act("designerToolbox", { module: form, kind: "CheckBox", left: 60, top: 230 });
  check("a drop on the open page lands on it", ontoPage.did === true, ontoPage.detail);
  const withPageDrop = await tabText();
  check("and its line is indented under THAT page, not under the first",
    /<Page\b[^>]*\bName="Page2"[^>]*>\r?\n {12}<CheckBox\b[^>]*\bName="CheckBox1"/.test(withPageDrop),
    withPageDrop.split(/\r?\n/).slice(-6).join(" / "));

  // New Page and Delete Page, where the native designer keeps them: the strip's own menu.
  const stripMenu = await api.act("designerTabMenu", { module: form, container: "Wizard" });
  check("the tab strip's menu offers the native designer's own pair",
    stripMenu.detail === "New Page | Delete Page", stripMenu.detail);

  await api.act("chooseMenuItem", { label: "New Page" });
  // Waited for on the CANVAS, not the document. The document takes the new page immediately and
  // the canvas only draws it once the lint round trip brings the parsed draft back, so a wait
  // that watched the text and then read the strip in the same tick read the strip too early.
  // Waited for on the CANVAS, not the document. The document takes the new page immediately and
  // the canvas only draws it once the lint round trip brings the parsed draft back, so a wait
  // that watched the text and then read the strip in the same tick read the strip too early.
  // Reports what it last saw rather than throwing bare, because "the strip is short a page" and
  // "the page never reached the document" want different answers.
  const added = await waitFor("the new page to be drawn on the strip", async () => {
    if (!/^ {8}<Page\b[^>]*\bName="Page3"[^>]*\bCaption="Page3"/m.test(await tabText())) {
      return null;
    }

    const seen = await canvasNow();
    return seen.containers.find((one) => one.name === "Wizard")?.page === "Page3" ? seen : null;
  }, { budgetMs: 15000 }).catch(async () => {
    const seen = await canvasNow();
    check("New Page reaches the canvas", false, JSON.stringify({
      strip: seen.containers.find((one) => one.name === "Wizard"),
      draft: seen.draft, dirty: seen.dirty, selected: seen.selected,
      lint: (await api.act("designerLint", { module: form })).data,
    }));
    return seen;
  });
  check("New Page writes the line MSForms would name and opens it",
    added.selected === "Page3" && added.containers.find((one) => one.name === "Wizard")?.page === "Page3",
    JSON.stringify(added.containers.find((one) => one.name === "Wizard")));

  // Delete Page acts on the page that is OPEN, whether the right-click landed on a tab or on
  // the empty end of the strip: "this page" means the one on screen.
  await api.act("designerTabMenu", { module: form, container: "Wizard", tab: "Page2" });
  await api.act("chooseMenuItem", { label: "Delete Page" });
  const pageGone = await waitFor("the page to leave the document", async () =>
    /<Page\b[^>]*\bName="Page2"/.test(await tabText()) ? null : await tabText(), { budgetMs: 15000 });
  check("Delete Page takes the page AND everything on it",
    !/\bName="CheckBox1"/.test(pageGone) && /<Page\b[^>]*\bName="Page1"/.test(pageGone)
    && /<Page\b[^>]*\bName="Page3"/.test(pageGone),
    pageGone.split(/\r?\n/).filter((line) => /Page/.test(line)).join(" / "));

  await press("z", "ctrlKey: true,");
  const undonePage = await waitFor("one undo to give the page back", async () =>
    /<Page\b[^>]*\bName="Page2"/.test(await tabText()) ? await tabText() : null, { budgetMs: 15000 });
  check("and one Ctrl+Z gives the whole page back, child and all",
    /\bName="CheckBox1"/.test(undonePage));

  // A page's ground only became double-clickable when the body took the page's identity, and a
  // PAGE has an event of its own: it raises Click, the VBE's own object list carries pages beside
  // the controls, and the host writes `Page1_Click` for one - measured on a fresh form, after a
  // first reading that said otherwise and was a read taken too early.
  //
  // Pressed TWICE, the way a hand does it: the view pairs doubles by name and clock at the press,
  // because the browser's dblclick never fires when the first click's re-render replaces the
  // element - so a dispatched dblclick would prove a listener real hands cannot reach.
  const doubled = await api.ask(`(() => {
    const find = () => ${inViewAll(".dc")}.find((one) => one.dataset.kind === "Page");
    const first = find();
    if (!first) { return "no page body"; }
    first.scrollIntoView({ block: "nearest", inline: "nearest" });
    const name = first.dataset.control;
    // Re-found for EVERY event: the first press selects, the selection re-renders, and a
    // reference held across that is a detached node whose events bubble to nothing - the
    // exact blindness the by-name detector exists to survive. A hand presses coordinates.
    const press = (type, buttons) => {
      const body = find();
      if (!body) { return; }
      const box = body.getBoundingClientRect();
      body.dispatchEvent(new PointerEvent(type, {
        bubbles: true, cancelable: true, pointerId: 1, isPrimary: true, button: 0, buttons,
        clientX: Math.round(box.left + box.width / 2), clientY: Math.round(box.top + box.height / 2) }));
    };
    press("pointerdown", 1); press("pointerup", 0);
    press("pointerdown", 1); press("pointerup", 0);
    return name;
  })()`);
  const pageHandler = new RegExp(`Private Sub ${doubled}_Click\\(\\)`);
  await waitFor("the page's own handler to land in the code-behind", async () =>
    pageHandler.test((await api.readModule(form, project)).text ?? ""), { budgetMs: 15000 });
  check("a double-click on a page's ground writes that page's own Click handler",
    true, `double-clicked ${doubled}`);
  await api.act("activate", { module: form, face: "design" });

  // A MultiPage the document leaves empty is a document the dialect allows, and the menu says
  // what it can do about it rather than offering an item that would delete nothing.
  await api.act("designerDelete", { module: form, control: "Page1" });
  await api.act("designerDelete", { module: form, control: "Page2" });
  await api.act("designerDelete", { module: form, control: "Page3" });
  const emptied = await waitFor("the MultiPage to be drawn with no tabs at all", async () => {
    const one = await containerNow("Wizard");
    return one?.tabs.length === 0 ? one : null;
  }, { budgetMs: 15000 });
  check("a MultiPage with no pages draws as bare chrome", emptied.open === -1, JSON.stringify(emptied));
  const emptyMenu = await api.act("designerTabMenu", { module: form, container: "Wizard" });
  check("and its menu greys Delete Page rather than offering to delete nothing",
    emptyMenu.detail === "New Page | Delete Page (disabled)", emptyMenu.detail);
  await api.act("key", { code: "Escape" });

  // A TabStrip switches too, and honestly shows nothing new: its tabs are an index rather than
  // containers, and the runtime draws the same controls under every one of them.
  const viewsTab = await api.act("designerOpenTab", { module: form, container: "Views", tab: 2 });
  check("a TabStrip's tab is picked up too", viewsTab.did === true, viewsTab.detail);
  const views = await containerNow("Views");
  check("it marks the tab, names no page, and selects the control itself",
    views?.open === 1 && views?.page === "" && (await canvasNow()).selected === "Views",
    JSON.stringify(views));
  // A TABSTRIP'S TABS are in the dialect since 2026-08-16: printed under the strip, parsed back,
  // added and removed by the same diff, and drawn from the projection rather than from a list of
  // strings that rode beside it. Before that the strip's line stood with nothing indented under
  // it while the canvas drew two tabs (the owner: "in the markdown, i dont see anything indented
  // under the tab view").
  const stripTabs = (text) => text.split(/\r?\n/)
    .filter((line) => /^\s+<Tab\b/.test(line))
    .map((line) => /\bName="([^"]*)"/.exec(line)?.[1] ?? "?").join(" | ");
  check("the strip's tabs are lines of the document, indented under it",
    stripTabs(await tabText()) === "Tab1 | Tab2", stripTabs(await tabText()));

  const tabMenu = await api.act("designerTabMenu", { module: form, container: "Views" });
  check("and its menu offers the pair a strip can keep - tabs, not pages",
    tabMenu.detail === "New Tab | Delete Tab", tabMenu.detail);

  await api.act("chooseMenuItem", { label: "New Tab" });
  const grown = await waitFor("the new tab to reach the document", async () =>
    /<Tab\b[^>]*\bName="Tab3"[^>]*\bCaption="Tab3"/.test(await tabText()) ? await canvasNow() : null, { budgetMs: 15000 });
  check("New Tab writes a line and the canvas draws it, open",
    grown.containers.find((one) => one.name === "Views")?.tabs.length === 3
    && grown.containers.find((one) => one.name === "Views")?.open === 2,
    JSON.stringify(grown.containers.find((one) => one.name === "Views")));

  await api.command("save");
  await waitFor("the tab to reach the FORM", async () =>
    /<Tab\b[^>]*\bName="Tab3"/.test(await api.designerMarkup(form, project)), { budgetMs: 20000 });
  check("and Ctrl+S puts it on the form itself - MSForms' own Tabs.Add", true);

  await api.act("designerTabMenu", { module: form, container: "Views", tab: 3 });
  await api.act("chooseMenuItem", { label: "Delete Tab" });
  await api.command("save");
  await waitFor("the tab to leave the form again", async () =>
    !/<Tab\b[^>]*\bName="Tab3"/.test(await api.designerMarkup(form, project)), { budgetMs: 20000 });
  check("Delete Tab takes it back off, through the same diff", true);

  check("the FORM kept all of its pages throughout: none of this reached it",
    (await api.designer(form, project)).controls.filter((one) => one.type === "Page").length === 2,
    JSON.stringify((await api.designer(form, project)).controls
      .filter((one) => one.type === "Page").map((one) => one.name)));

  await api.act("designerSetMarkup", { module: form, markup: await api.designerMarkup(form, project) });
  await waitFor("the document back at canonical after the container rows", async () =>
    ((await canvasNow())?.dirty) === false, { budgetMs: 15000 });

  // ---- zoom: the picture scales, the points do not ----
  //
  // The whole design of this one is that only the SCREEN boundary knows about the scale, so the
  // row that matters is the drag: twelve points at 200% has to be twelve points in the document.

  const zoomAt = async () => (await api.act("designerZoom", { module: form })).data;
  const formBox = async () => api.ask(`(() => {
    const el = ${inView(".dc-form")};
    const box = el.getBoundingClientRect();
    return Math.round(box.width) + "x" + Math.round(box.height);
  })()`);

  check("the canvas opens at 100%", (await zoomAt()) === 100, String(await zoomAt()));
  const drawnAt100 = String(await formBox());

  const twiceSize = await api.act("designerZoom", { module: form, to: 200 });
  check("and takes a zoom", twiceSize.did === true, twiceSize.detail);
  const drawnAt200 = String(await formBox());
  check("the form is drawn twice the size, and the DOCUMENT has not moved",
    Math.abs(Number(drawnAt200.split("x")[0]) - Number(drawnAt100.split("x")[0]) * 2) < 4
    && ((await api.act("designerCanvas", { module: form })).data?.dirty) === false,
    `${drawnAt100} became ${drawnAt200}`);

  const zoomedPlace = async () => {
    const box = await placed("HoldToggle");
    return box.left === null || box.top === null ? [] : [box.left, box.top];
  };

  // The snap OFF while the zoom is measured, so the row reads the gesture rather than the grid -
  // and put back the way it was found, because the grid rows below open by measuring an unsnapped
  // drag and a section that changes a setting owes the next one what it borrowed.
  const snapWas = (await api.ui()).settings?.designerSnap ?? "grid";
  await api.settings({ designerSnap: "off" });
  await waitFor("the snap off for the zoom rows", async () =>
    (await api.ui()).settings?.designerSnap === "off", { budgetMs: 15000 });

  const zoomedFrom = await zoomedPlace();
  await api.act("designerDrag", { module: form, control: "HoldToggle", dx: 12, dy: 0 });
  await waitFor("the drag at 200% to land", async () =>
    (await zoomedPlace())[0] !== zoomedFrom[0], { budgetMs: 15000 });
  const zoomedTo = await zoomedPlace();
  check("a drag of twelve POINTS moves twelve points, whatever the zoom is drawing",
    zoomedTo[0] === zoomedFrom[0] + 12 && zoomedTo[1] === zoomedFrom[1],
    `${zoomedFrom.join(",")} + 12,0 at 200% is ${zoomedTo.join(",")}`);

  const toFit = await api.act("designerZoom", { module: form, to: "fit" });
  check("Fit picks the largest scale that shows the whole form",
    toFit.did === true && /zoom \d+%/.test(toFit.detail ?? ""), toFit.detail);
  const fitPercent = await zoomAt();
  check("and it is a scale that fits, not one that was asked for",
    fitPercent >= 25 && fitPercent <= 200, String(fitPercent));

  await api.act("designerZoom", { module: form, to: 100 });
  check("back to 100% for the rows that follow", (await zoomAt()) === 100, String(await zoomAt()));
  await api.settings({ designerSnap: snapWas });
  await waitFor("the snap back the way the zoom rows found it", async () =>
    (await api.ui()).settings?.designerSnap === snapWas, { budgetMs: 15000 });

  // ---- pictures: an IPictureDisp becomes pixels the canvas can draw, and back ----
  //
  // The read is GDI rather than COM - the handle off the picture, GetDIBits for a bitmap and
  // DrawIconEx onto a DIB section for an icon - so both halves are exercised: the fixture's
  // Badge wears a PNG (a bitmap once loaded) and its OK button an ICO.

  const canvasPictures = async () =>
    (await api.act("designerCanvas", { module: form })).data?.pictures ?? [];

  const drawnPictures = await canvasPictures();
  const badgeDrawn = drawnPictures.find((one) => one.name === "Badge");
  check("the canvas draws the Image control's picture instead of a crossed box",
    (badgeDrawn?.bytes ?? 0) > 1000, JSON.stringify(badgeDrawn ?? null));
  check("and places it the way PictureSizeMode says - zoomed, centred",
    badgeDrawn?.where === "contain center center", badgeDrawn?.where);

  const buttonDrawn = drawnPictures.find((one) => one.name === "OkButton");
  check("a picture BESIDE a caption is drawn beside it, where PicturePosition says",
    (buttonDrawn?.bytes ?? 0) > 100 && buttonDrawn?.where === "row center",
    JSON.stringify(buttonDrawn ?? null));

  // The PANEL: a picture row says what it holds rather than showing a path, because MSForms
  // keeps the pixels and forgets the file they came from.
  await api.act("designerSelect", { module: form, control: "Badge" });
  const badgePanel = await waitFor("the panel to show the Image control", async () => {
    const shown = (await api.ui()).properties;
    return shown?.component === "Badge" ? shown : null;
  }, { budgetMs: 15000 });

  const badgeRow = badgePanel.rows.find((row) => row.name === "Picture");
  check("the panel's Picture row says what it holds, with a thumbnail behind it",
    badgeRow?.picture === true && badgeRow?.value === "(Bitmap)" && (badgeRow?.previewBytes ?? 0) > 1000,
    JSON.stringify(badgeRow ?? null));
  check("and the placement rows are there, named the way every enum row is",
    badgePanel.rows.find((row) => row.name === "PictureSizeMode")?.value === "fmPictureSizeModeZoom",
    badgePanel.rows.find((row) => row.name === "PictureSizeMode")?.value);

  // Cleared through the panel's own Clear, which is the empty write - and the canvas follows.
  await api.act("editProperty", { name: "Picture", value: "" });
  await waitFor("the cleared picture to leave the canvas", async () =>
    (await canvasPictures()).every((one) => one.name !== "Badge"), { budgetMs: 15000 });
  check("Clear takes the picture off, and the crossed box comes back", true);

  const emptyRow = (await api.ui()).properties?.rows?.find((row) => row.name === "Picture");
  check("the row says (None) with nothing to preview",
    emptyRow?.value === "(None)" && emptyRow?.previewBytes === undefined, JSON.stringify(emptyRow ?? null));

  // Written back through the api, by PATH - the one place the two paths differ, and a
  // designated deviation: the panel's Browse raises the machine's file dialog to PRODUCE a
  // path, and then takes exactly this road.
  const rewritten = await api.designerEdit("set", {
    module: form, project, name: "Badge", property: "Picture", value: FORM_PICTURES.bitmap,
  });
  check("set Picture takes a file path and answers what the control now holds",
    rewritten.ok === true && /Picture is \(Bitmap\)/.test(rewritten.detail ?? ""), rewritten.detail);
  await waitFor("the canvas to draw the picture that was just written", async () =>
    (await canvasPictures()).some((one) => one.name === "Badge"), { budgetMs: 15000 });
  check("liveness carries a picture like any other edit", true);

  await api.designerEdit("set", {
    module: form, project, name: "Badge", property: "PictureSizeMode", value: "3",
  });

  // THE ICON ROAD, on an icon built to be one: OLE's loader reads classic icons, whose frames
  // are BMPs, and hands back a picture of kind Icon - which GetDIBits cannot read at all, so the
  // canvas has to draw it onto a DIB section instead. Nothing in the repository is a classic
  // icon (see form-plan), so the suite writes one: sixteen by sixteen, solid red, opaque.
  const iconPath = join(tmpdir(), "xlide-suite-classic.ico");
  writeFileSync(iconPath, classicIcon(16, [0x00, 0x00, 0xFF, 0xFF]));

  const asIcon = await api.designerEdit("set", {
    module: form, project, name: "Badge", property: "Picture", value: iconPath,
  });
  check("a classic icon loads as an ICON, not as a flattened bitmap",
    /Picture is \(Icon\)/.test(asIcon.detail ?? ""), asIcon.detail);
  await waitFor("the icon to reach the canvas, drawn rather than read", async () =>
    (await canvasPictures()).some((one) => one.name === "Badge" && one.bytes > 1000), { budgetMs: 15000 });
  check("and an icon becomes pixels the canvas can draw - the DrawIconEx road", true);

  rmSync(iconPath, { force: true });
  await api.designerEdit("set", {
    module: form, project, name: "Badge", property: "Picture", value: FORM_PICTURES.bitmap,
  });

  const refused = await api.designerEdit("set", {
    module: form, project, name: "Badge", property: "Picture",
    value: join(tmpdir(), "xlide-there-is-no-such-picture.png"),
  }).catch((why) => ({ refusal: why.message }));
  check("a picture that is not there is refused by name, and nothing is written",
    /there is no file at/i.test(refused.refusal ?? refused.detail ?? ""),
    refused.refusal ?? refused.detail);
  check("and the control keeps the picture it had",
    (await api.designer(form, project)).controls
      .find((one) => one.name === "Badge")?.picture?.kind === "(Bitmap)");

  // A ROW NOBODY CAN SET IS NOT A PROPERTY OF THE DESIGN (the owner, 2026-08-16). The form's
  // panel listed CanPaste, CanUndo and CanRedo - questions about the editing session, False on
  // a form nobody has touched, and unwritable by their nature.
  await api.act("designerSelect", { module: form });
  const settableOnly = await waitFor("the panel back on the form", async () => {
    const shown = (await api.ui()).properties;
    return shown?.component === form ? shown : null;
  }, { budgetMs: 15000 });
  check("the panel shows no row it cannot write",
    settableOnly.rows.every((row) => row.writable),
    settableOnly.rows.filter((row) => !row.writable).map((row) => row.name).join(",") || "none");
  check("so the session-state rows are gone",
    !settableOnly.rows.some((row) => ["CanPaste", "CanUndo", "CanRedo"].includes(row.name)));
  check("and a form's own Picture row is there to be set",
    settableOnly.rows.some((row) => row.name === "Picture" && row.picture === true));

  // ---- the grid: pointer gestures land on it, the keyboard never does ----

  const gridPlace = async () => {
    const box = await placed("HoldToggle");
    return box.left === null || box.top === null ? [] : [box.left, box.top];
  };

  await api.act("designerSetMarkup", { module: form, markup: await canonicalNow() });
  await waitFor("the document back at canonical before the grid rows", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === false, { budgetMs: 15000 });

  // The same seven-point drag with the grid off lands seven points away.
  const beforeFree = await gridPlace();
  await api.act("designerDrag", { module: form, control: "HoldToggle", dx: 7, dy: 7 });
  await waitFor("the unsnapped drop", async () =>
    (await gridPlace())[0] === beforeFree[0] + 7, { budgetMs: 15000 });
  const free = await gridPlace();
  check("with the grid off a drag lands exactly where the pointer left it",
    free[0] === beforeFree[0] + 7 && free[1] === beforeFree[1] + 7,
    `${beforeFree.join(",")} + 7,7 is ${free.join(",")}`);

  await api.settings({ designerSnap: "grid" });
  await waitFor("the page to hear that the grid is back", async () =>
    (await api.ui()).settings?.designerSnap === "grid", { budgetMs: 15000 });

  // THE GRID IS ACTUALLY PAINTED - checked HERE, with snapping back on, not merely declared. It was declared and invisible: a radial
  // gradient with both colour stops at 0.5px is a half-pixel disc with no ramp, and Chromium
  // rasterised it to nothing, so the toggle lit and the form's ground stayed bare (the owner,
  // 2026-08-17: "grid snap indicator is lit, but grid snap not showing on form"). Nothing checked
  // that the dots existed, which is why it shipped. A row cannot count pixels from here, so it
  // pins the thing that was wrong: a dot with a whole pixel of radius to land on.
  const painted = String(await api.ask(
    `getComputedStyle(${inView(".dc-form-client")}).backgroundImage`));
  // THE SOLID CENTRE, which is the second stop: the browser expands `0 1px` into `0px` and `1px`,
  // and the transparent stop trails behind them. Taking the first number reads the 0 the dot
  // starts at, and taking the largest reads the transparent edge - neither says how much of the
  // dot is actually painted. The broken shape was `0.5px, 0.5px`, whose second stop is 0.5.
  const dotStops = [...painted.matchAll(/([\d.]+)px/g)].map((one) => Number(one[1]))
    .sort((a, b) => a - b);
  const solid = dotStops.length >= 2 ? dotStops[dotStops.length - 2] : 0;
  check("the grid is painted with a dot big enough to land on a pixel",
    /radial-gradient/.test(painted) && solid >= 1,
    `solid to ${solid}px of ${JSON.stringify(dotStops)} in ${painted.slice(0, 70)}`);


  // And with it on, four points land on the nearest six - a delta chosen because it does NOT
  // reach a grid line by itself, so the row can tell snapping from arithmetic. Seven points
  // from 119 reaches 126, which is on the grid already and would have proved nothing.
  await api.act("designerDrag", { module: form, control: "HoldToggle", dx: 4, dy: 4 });
  await waitFor("the snapped drop", async () =>
    (await gridPlace())[0] !== free[0], { budgetMs: 15000 });
  const snapped = await gridPlace();
  check("with it on a drag lands on the grid instead of where the pointer stopped",
    snapped[0] % 6 === 0 && snapped[1] % 6 === 0
      && snapped[0] !== free[0] + 4 && snapped[1] !== free[1] + 4,
    `landed at ${snapped.join(",")} from ${free.join(",")}; four points is `
    + `${free[0] + 4},${free[1] + 4} and the six-point grid is the nearest multiple`);

  // ALT overrides it, held rather than toggled: the one control that will not sit on the grid
  // is a reason to escape it for a moment, not to turn the grid off and forget to turn it on.
  const heldDrag = await api.act("designerDrag", { module: form, control: "HoldToggle", dx: 7, dy: 7, alt: 1 });
  check("a drag with alt held is accepted", heldDrag.did === true, heldDrag.detail);
  await waitFor("the drop with alt held", async () =>
    (await gridPlace())[0] !== snapped[0], { budgetMs: 15000 });
  const overridden = await gridPlace();
  check("holding alt overrides the grid for that one gesture",
    overridden[0] === snapped[0] + 7 && overridden[1] === snapped[1] + 7,
    `${snapped.join(",")} + 7,7 with alt held is ${overridden.join(",")}, and the grid `
    + "would have said " + `${Math.round((snapped[0] + 7) / 6) * 6},${Math.round((snapped[1] + 7) / 6) * 6}`);

  // The keyboard is exact whatever the grid says: an arrow moves ONE point, off the grid.
  await api.act("designerSelect", { module: form, control: "HoldToggle" });
  await press("ArrowRight");
  await waitFor("the nudge", async () => (await gridPlace())[0] === overridden[0] + 1, { budgetMs: 15000 });
  const nudged = await gridPlace();
  check("an arrow still moves a single point, off the grid and on purpose",
    nudged[0] === overridden[0] + 1,
    `${overridden.join(",")} then one arrow is ${nudged.join(",")}`);

  check("and the setting is as the developer's file had it, which the rows above borrowed", true);

  // ---- or it lines up with the OTHER CONTROLS, which is the other mode ----
  //
  // One or the other, never both: two snapping systems that disagree give a control two right
  // answers and it takes whichever the code asked first. The rows below switch the mode and
  // prove the difference by putting one control somewhere OFF the grid - with alt, the only way
  // to get there - and then bringing another near it: landing on that odd number is alignment,
  // and landing on the nearest six would be the grid it is no longer using.

  await api.settings({ designerSnap: "objects" });
  await waitFor("the page to hear that snapping follows the controls", async () =>
    (await api.ui()).settings?.designerSnap === "objects", { budgetMs: 15000 });

  const placeOf = async (control) => {
    const box = await placed(control);
    return box.left === null || box.top === null ? [] : [box.left, box.top];
  };

  const boxWas = await placeOf("NameBox");
  await api.act("designerDrag", { module: form, control: "NameBox", dx: 1, dy: 0, alt: 1 });
  await waitFor("the off-grid neighbour", async () =>
    (await placeOf("NameBox"))[0] === boxWas[0] + 1, { budgetMs: 15000 });
  const oddEdge = (await placeOf("NameBox"))[0];
  check("a control can be put somewhere the grid would never allow", oddEdge % 6 !== 0,
    `${oddEdge} is not a multiple of six, which is what makes the next row mean something`);

  // Whatever the grid would have said is irrelevant in this mode - that is the point of the
  // either/or - but the odd number still tells an alignment from an accident.

  const holdWas = await placeOf("HoldToggle");
  const lining = await api.act("designerDrag",
    { module: form, control: "HoldToggle", dx: oddEdge - holdWas[0] + 2, dy: 0 });
  await waitFor("the aligned drop", async () =>
    (await placeOf("HoldToggle"))[0] !== holdWas[0], { budgetMs: 15000 });
  const lined = await placeOf("HoldToggle");
  check("a drag that comes near another control's edge lines up with it",
    lined[0] === oddEdge,
    `landed at ${lined[0]} where the neighbour's edge is ${oddEdge}`);

  check("and the act says what it lined up with, which is the guide the canvas drew",
    /lined up with x=/.test(lining.detail ?? ""), lining.detail);

  // Alt escapes the alignment too, not only the grid.
  const before = await placeOf("HoldToggle");
  const freed = await api.act("designerDrag",
    { module: form, control: "HoldToggle", dx: 2, dy: 0, alt: 1 });
  await waitFor("the drop with alt held", async () =>
    (await placeOf("HoldToggle"))[0] === before[0] + 2, { budgetMs: 15000 });
  check("and alt escapes the alignment as well as the grid",
    !/lined up with/.test(freed.detail ?? ""), freed.detail);

  await api.settings({ designerSnap: "grid" });
  await waitFor("snapping back the way the developer's file had it", async () =>
    (await api.ui()).settings?.designerSnap === "grid", { budgetMs: 15000 });

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
      return read.did === true && String(read.data).startsWith(`<Form Name="${renamedForm}"`) ? read : null;
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

  // ---- A NAVIGATION INTO A FORM'S CODE TAKES THE SLOT FROM ITS DESIGNER TAB ----
  //
  // "The shown module" carries a name and a workbook and no FACE, so a form whose designer tab
  // held the active slot reported that module as already showing: GoTo took its already-showing
  // branch and revealed a line in a document nobody was looking at. Clicking a Sub under a form
  // in the tree therefore did nothing visible at all (the owner, 2026-08-18).
  //
  // The consequence past the tab strip is why this is more than a cosmetic slot: `_activeDesignerTab`
  // is also what F5 and Ctrl+S read to decide whether they are applying a document or running a
  // module, so a navigation that left it standing left Run aimed at a form the developer had
  // navigated away from.
  await api.act("activate", { module: form, face: "design" });
  await waitFor("the designer tab holding the slot", async () =>
    (await api.ui()).workspace?.active?.face === "design", { budgetMs: 15000 });

  await api.caret(1, { module: form, column: 1 });
  const tookTheSlot = await waitFor("the code face to take the slot", async () => {
    const active = (await api.ui()).workspace?.active;
    return active?.module === form && !active.face ? active : null;
  }, { budgetMs: 15000 }).catch(() => null);
  check("a navigation into a form's code takes the active slot from its designer tab",
    tookTheSlot !== null, JSON.stringify((await api.ui()).workspace?.active));

  // ---- SIZE TO FIT COPIES THE RUNTIME, INCLUDING WHERE THE RUNTIME IS ABSURD ----
  //
  // The size cannot be worked out on the page: a check box's glyph, a button's chrome and a
  // picture drawn at natural size are none of them in the caption's ink, and a page-side read put
  // "Hold" at 18pt - narrower than its own caption. So the gesture asks the host, which sets
  // AutoSize, reads the box and puts BOTH the flag and the geometry back.
  //
  // The picture row is the one that pins the owner's decision (2026-08-18): a CommandButton
  // wearing a 256-square logo autosizes far past the 72x24 it started at, bigger than a caption
  // could ever justify, and that is the answer this writes. Not capped, not clamped to the
  // container, not quietly replaced by the caption's width.
  //
  // ASKED IN RATIOS, NOT IN POINTS. These pinned MSForms' exact output - 29x22 and 220.5x198.1 -
  // and MSForms derives both from the screen's DPI and the font's metrics, so the numbers are
  // facts about the MACHINE as much as about the product. On a differently scaled display the
  // same correct behaviour reads 26x19 and 153x133, and three checks fail for a reason that has
  // nothing to do with the code (2026-08-22, a release gate). What the rows are FOR survives the
  // rescaling: that the size came from the host rather than from the page's guess at the
  // caption's ink, and that a picture button is not clamped to its caption or its container.
  const beforeFit = await tabText();
  const toggleWas = await placed("HoldToggle");
  await api.act("designerSelect", { module: form, control: "HoldToggle" });
  const fitToggle = await api.act("designerFormat", { module: form, how: "fit" });
  const toggleBox = await placed("HoldToggle");

  // A page-side read of "Hold" put it at 18pt, NARROWER than its own caption, because the glyph
  // and the chrome are not in the ink. So the host's answer has to be wider than that, and the
  // box has to have actually moved.
  check("Size to Fit takes MSForms' own AutoSize, not the caption's ink",
    fitToggle.did === true
      && toggleBox.width > 18
      && toggleBox.height > 0
      && (toggleBox.width !== toggleWas.width || toggleBox.height !== toggleWas.height),
    `${fitToggle.detail}; ${await placedAt("HoldToggle")} from ${toggleWas.width}x${toggleWas.height}`);

  await api.act("designerSelect", { module: form, control: "OkButton" });
  await api.act("designerFormat", { module: form, how: "fit" });
  const fitted = await placed("OkButton");

  // Started 72x24. A caption fit would leave it about that wide and no taller; a PICTURE fit is
  // several times both. The multiples hold at any scaling - the ratio is the picture's, not the
  // display's.
  check("...and a button wearing an oversized picture fits to the PICTURE, uncapped",
    fitted.width > 72 * 1.5 && fitted.height > 24 * 4,
    `${await placedAt("OkButton")} from 72x24`);
  check("...which is an order larger than the 72x24 it started at, and is not capped",
    fitted.height > fitted.width * 0.5 && fitted.height > 24 * 4,
    `${fitted.width}x${fitted.height} from 72x24 - a caption fit would be wide and short`);

  await api.act("designerSetMarkup", { module: form, markup: beforeFit });
  await waitFor("the document back after the fit rows", async () =>
    (await tabText()) === beforeFit, { budgetMs: 15000 });

  // ---- CLOSING A DESIGNER TAB WITH UNAPPLIED EDITS ASKS ----
  //
  // It did not until 2026-08-16: the close was unconditional and a Ctrl+W after three moves lost
  // the three moves with nothing said (found in the hunt). The question is the PAGE's, because
  // the state is - unapplied markup lives in the view, and nothing host-side can know to hold
  // the close - but it is the same modal, the same three buttons, and `closeActive` reports the
  // outcome rather than the request, exactly as a module's does.
  const designerTabNow = async () => ((await api.ui()).workspace?.groups ?? [])
    .flatMap((group) => group.tabs).find((tab) => tab.module === form && tab.face === "design");

  await api.act("activate", { module: form, face: "design" });
  await api.act("designerDrag", { module: form, control: "HoldToggle", dx: 6, dy: 0 });
  await waitFor("the document to go dirty before the close", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === true, { budgetMs: 15000 });

  const held = await api.act("closeActive", {});
  check("closing a designer tab with unapplied edits is HELD, not done",
    held.did === false && /Unsaved changes/.test(held.detail ?? ""), held.detail);
  check("...and the question standing is the one a module's unsaved text raises",
    ((await api.ui()).dialogs ?? []).some((one) => one.id === "close-confirm-card"),
    JSON.stringify((await api.ui()).dialogs ?? []));

  await api.act("answerCloseConfirm", { answer: "cancel" });
  await waitFor("the question to go", async () =>
    ((await api.ui()).dialogs ?? []).length === 0, { budgetMs: 15000 });
  check("Cancel leaves the tab open with its edits",
    (await designerTabNow()) !== undefined
    && ((await api.act("designerCanvas", { module: form })).data?.dirty) === true);

  await api.act("closeActive", {});
  await api.act("answerCloseConfirm", { answer: "discard" });
  await waitFor("the designer tab to leave on Don't Save", async () =>
    (await designerTabNow()) === undefined, { budgetMs: 15000 });
  const afterDiscard = (await api.designer(form, project)).controls
    .find((one) => one.name === "HoldToggle");
  check("Don't Save closes it and the FORM never moved",
    afterDiscard?.left === 112, String(afterDiscard?.left));

  await api.pane("open", { module: form, face: "design" });
  await waitFor("the tab back for the Save half", async () =>
    (await designerTabNow()) !== undefined, { budgetMs: 15000 });
  await waitFor("its document to arrive", async () =>
    /HoldToggle/.test(String((await api.act("designerMarkup", { module: form })).data)), { budgetMs: 15000 });
  await api.act("designerDrag", { module: form, control: "HoldToggle", dx: 6, dy: 0 });
  await waitFor("the document dirty again", async () =>
    ((await api.act("designerCanvas", { module: form })).data?.dirty) === true, { budgetMs: 15000 });
  await api.act("closeActive", {});
  await api.act("answerCloseConfirm", { answer: "save" });
  await waitFor("Save to reach the form itself", async () =>
    ((await api.designer(form, project)).controls.find((one) => one.name === "HoldToggle")?.left ?? 0) > 112,
  { budgetMs: 20000 });
  check("Save means what Ctrl+S means here - the document applied, then the workbook", true);
  await waitFor("and the tab goes after the save", async () =>
    (await designerTabNow()) === undefined, { budgetMs: 15000 });
  check("and the tab goes with it", true);

  await api.designerEdit("set", { module: form, project, name: "HoldToggle", property: "Left", value: "112" });

  await api.pane("open", { module: form, face: "design" });
  await waitFor("the tab back for the rows below", async () =>
    (await designerTabNow()) !== undefined, { budgetMs: 15000 });
  await api.pane("close", { module: form, face: "design" });
  await waitFor("the designer tab to leave the strip", async () =>
    !((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === form && tab.face === "design"), { budgetMs: 15000 });
  check("and a CLEAN designer tab still closes without a question",
    ((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === form && !tab.face));

  // ---- A COLOUR ROW WITH NO DESIGNER TAB OPEN: THE HOST PATH ----
  //
  // The rows above write colours through the DOCUMENT, because a Properties row on an open
  // designer tab joins that tab's transaction (#68). With no designer tab there is no document
  // to join and the row goes straight to the component, which is the path #66 was filed against
  // and the only one of the two that had never been driven. It is reachable in ordinary use -
  // a form nobody has opened a designer for, selected in the tree - so it is pinned here, at
  // the one point in the suite where the designer tab is known to be down.
  //
  // #66's narrowing said the raw `#rrggbb` reached the COM put and MSForms threw on it. It does
  // not: the VBE's Property for a colour reads as VT_I4, so text that will not convert is
  // refused by the writer BEFORE the model is touched, and text that will convert is a number by
  // then. Both halves are pinned, because the value of this row is the second one - a refusal
  // that costs nothing is what keeps a bad value from reaching the designer at all.
  await waitFor("the panel on the form, with no designer tab to write through", async () =>
    (await api.ui()).properties?.component === form, { budgetMs: 15000 });
  check("the designer tab really is down for these rows",
    !((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === form && tab.face === "design"));

  const hostColour = await api.act("editProperty", { name: "BackColor", value: "#0080ff" });
  check("a #rrggbb typed with no designer tab open writes straight through to the component",
    hostColour.did === true, hostColour.detail);
  check("...and reaches the form as the NUMBER the model stores, not as the text",
    ((await api.designer(form, project)).form?.backColor ?? 0) === 16744448,
    String((await api.designer(form, project)).form?.backColor));

  const hostRefused = await api.act("editProperty", { name: "BackColor", value: "not a colour" });
  check("text that is no colour at all is refused before the model is touched",
    hostRefused.did === false && /not a whole number/.test(hostRefused.detail ?? ""),
    hostRefused.detail);
  check("...and the refusal costs nothing - the form still reads, holding what it held",
    ((await api.designer(form, project)).form?.backColor ?? 0) === 16744448,
    String((await api.designer(form, project)).form?.backColor));

  await api.act("editProperty", { name: "BackColor", value: "Button Face" });

  // ---- A REMOVED FORM TAKES ITS DESIGNER TAB WITH IT ----
  //
  // It did not, and what stood in its place was an overlay reading "<form> has no designer" -
  // which is what a form mid-teardown answers, because the component is still in the collection
  // for a moment after a Remove (the owner, 2026-08-16: "sometimes i see an overlay that says
  // entryform2 has no designer"). A tab over a form that is gone is a corpse; the removal knows,
  // so it collects it rather than leaving the next request to find out.
  const ghost = `Ghost${process.pid % 1000}`;
  await api.component("add", { kind: 3, name: ghost, project });
  await api.pane("open", { module: ghost, face: "design" });
  await waitFor("the ghost's designer tab to stand", async () =>
    ((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === ghost && tab.face === "design"), { budgetMs: 15000 });

  await api.component("remove", { name: ghost, project });
  await waitFor("the tab to go with the form", async () =>
    !((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === ghost), { budgetMs: 15000 });
  check("a removed form takes its designer tab, leaving no overlay behind", true);

  const leftover = await api.ask('(() => { const el = document.querySelector(".designer-notice");'
    + ' return el && !el.hidden ? (el.textContent || "").trim() : ""; })()');
  check("...and no canvas is left saying a form has no designer",
    String(leftover) === "", String(leftover).slice(0, 90));

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
  // THE GRID GOES BACK ON HERE, not where the rows that borrowed it finish.
  //
  // It ships on and the gesture rows turn it off, so a run that dies between the two leaves the
  // DEVELOPER'S OWN settings file with snapping disabled - and the next run's first grid row
  // then fails for a reason that has nothing to do with the code. Which is exactly what
  // happened the first time this suite aborted mid-section (2026-08-16). A borrowed setting is
  // returned in the finally, like anything else borrowed.
  await api.settings({ designerSnap: "grid" }).catch(() => {});

  // THE RENAMED-AWAY NAME TOO. The rename section takes the form to `${form}R` and back, so an
  // abort between those two steps leaves a form nothing removes - and a form left in the fixture
  // is not a tidiness problem: module-sync's export counts one plan item per module and a form
  // writes TWO files, so the next run of a suite that has nothing to do with designers fails on
  // a fixture the last run poisoned. Observed exactly that way on 2026-09-03, where an aborted
  // deep run left EntryForm2R behind and the gate failed on module-sync the next time.
  await api.component("remove", { name: `${form}R`, project }).catch(() => {});
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
