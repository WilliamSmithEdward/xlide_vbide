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
import { open, reporter } from "./xlide-api.mjs";
import { FORM_CONTROLS, FORM_MODULE, FORM_PROPERTIES, buildForm } from "./form-plan.mjs";

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

try {
  // ---- built through the model, as the fixture generator builds it ----

  await api.component("remove", { name: FORM_MODULE, project }).catch(() => {});
  await buildForm(api, project);

  const design = await api.designer(FORM_MODULE, project);

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
      wanted.parent ? found.parent === wanted.parent : found.parent === FORM_MODULE,
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
    module: FORM_MODULE, project, type: "commandButton", name: "ProbeButton",
    left: 10, top: 280, width: 60, height: 20,
  });
  check("add answers the name and the kind it made", probe.ok === true && probe.name === "ProbeButton"
    && probe.type === "CommandButton", JSON.stringify(probe));

  const captioned = await api.designerEdit("set", {
    module: FORM_MODULE, project, name: "ProbeButton", property: "Caption", value: "Probe", as: "text",
  });
  check("set answers what the property reads back", /Caption is Probe/.test(captioned.detail ?? ""),
    captioned.detail);

  const sized = await api.designerEdit("set", {
    module: FORM_MODULE, project, name: "ProbeButton", property: "Font.Size", value: "12",
  });
  check("a dotted set reaches the font", /Font\.Size is 12/.test(sized.detail ?? ""), sized.detail);

  const afterAdd = await api.designer(FORM_MODULE, project);
  const probeRow = afterAdd.controls.find((control) => control.name === "ProbeButton");
  check("the added control reads back where it was put, with its caption and font",
    probeRow !== undefined && near(probeRow.left, 10) && probeRow.caption === "Probe"
    && near(probeRow.font?.size, 12),
    JSON.stringify(probeRow ?? null));

  const removed = await api.designerEdit("remove", { module: FORM_MODULE, project, name: "ProbeButton" });
  const afterRemove = await api.designer(FORM_MODULE, project);
  check("remove takes it back out",
    removed.ok === true && !afterRemove.controls.some((control) => control.name === "ProbeButton"));

  const nested = await api.designerEdit("remove", { module: FORM_MODULE, project, name: "PickAir" });
  const afterNested = await api.designer(FORM_MODULE, project);
  check("a control inside a Frame is removed through its own container",
    nested.ok === true && !afterNested.controls.some((control) => control.name === "PickAir")
    && afterNested.controls.some((control) => control.name === "PickGround"));

  // ---- the form as text ----

  const markup = await api.designerMarkup(FORM_MODULE, project);
  check("the markup opens with the form line",
    markup.startsWith(`Form ${FORM_MODULE} "`), markup.split("\n")[0]);
  check("a nested control prints inside its container",
    /\r?\n  Frame Options "Freight" at 12,112 size 92x66\r?\n    OptionButton PickGround/.test(markup),
    markup.slice(0, 400));
  check("a page prints under its MultiPage, a control under the page",
    /\r?\n  MultiPage Wizard[^\r\n]*\r?\n    Page Page1 "Page1"\r?\n      CheckBox Agree/.test(markup));

  const idempotent = await api.applyMarkup(FORM_MODULE, markup, project);
  check("applying the form's own markup adds and removes nothing",
    idempotent.ok === true && idempotent.added.length === 0 && idempotent.removed.length === 0,
    JSON.stringify({ added: idempotent.added, removed: idempotent.removed, set: idempotent.set }));

  const withButton = `${markup.trimEnd()}\r\n  CommandButton MarkupBtn "Go" at 8,282 size 60x20\r\n`;
  const applied = await api.applyMarkup(FORM_MODULE, withButton, project);
  check("a line added to the document adds the control",
    applied.ok === true && applied.added.includes("MarkupBtn"), JSON.stringify(applied));

  const afterMarkup = await api.designer(FORM_MODULE, project);
  const btn = afterMarkup.controls.find((control) => control.name === "MarkupBtn");
  check("and it reads back placed, captioned, the kind the line said",
    btn !== undefined && btn.type === "CommandButton" && btn.caption === "Go"
    && near(btn.left, 8) && near(btn.top, 282),
    JSON.stringify(btn ?? null));

  const backAgain = await api.applyMarkup(FORM_MODULE, markup, project);
  check("re-applying the original document removes it again",
    backAgain.ok === true && backAgain.removed.includes("MarkupBtn"), JSON.stringify(backAgain));

  const refusedMarkup = await api.applyMarkup(FORM_MODULE, "Form X\n  Label L at banana\n", project)
    .catch((why) => why.message);
  check("a document that does not parse applies nothing, naming the line",
    /did not parse/.test(String(refusedMarkup)) && /line 2/.test(String(refusedMarkup)),
    String(refusedMarkup));
  const untouched = await api.designer(FORM_MODULE, project);
  check("and the form is untouched by it",
    untouched.controls.length === afterMarkup.controls.length - 1,
    `${untouched.controls.length} controls vs ${afterMarkup.controls.length} with the probe button`);

  // ---- refusals are answers ----

  const missing = await api.designer("NoSuchForm", project).catch((why) => why.message);
  check("a form that is not there is refused by name", /no component named NoSuchForm/.test(String(missing)),
    String(missing));

  const notAForm = await api.designer(plainModule, project).catch((why) => why.message);
  check(`a module that is not a form (${plainModule}) is refused as one`,
    /not a UserForm/.test(String(notAForm)), String(notAForm));

  const badKind = await api.designerEdit("add", { module: FORM_MODULE, project, type: "gizmo" })
    .catch((why) => why.message);
  check("an unknown control kind is refused with the list", /not a control kind/.test(String(badKind)),
    String(badKind));
} finally {
  await api.component("remove", { name: FORM_MODULE, project }).catch(() => {});
}

const swept = await api.project();
check("the form is gone again, leaving the fixture as found",
  !swept.components.some((component) => component.name === FORM_MODULE));

done();
