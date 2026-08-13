/*
 * The one declaration of the fixture form: every standard toolbox control, a Frame with
 * children, a MultiPage with a control on its first page.
 *
 * Two consumers, deliberately the same file. build-form-fixture.mjs BUILDS this through the
 * designer route, and designer-features.mjs VERIFIES a read against it - so what the fixture
 * holds and what the suite expects cannot drift apart, which is the failure mode a fixture
 * checked against a copy of its own description invites.
 *
 * Geometry is points. Captions are set explicitly because the model, unlike the native
 * toolbox gesture, adds controls with EMPTY captions - a Label with no caption is what
 * Controls.Add honestly produces.
 */

export const FORM_MODULE = "EntryForm";

export const FORM_PROPERTIES = [
  { property: "Caption", value: "Quarter Entry", as: "text" },
  { property: "Width", value: "360" },
  { property: "Height", value: "320" },
];

export const FORM_CONTROLS = [
  { type: "Label", name: "NameLabel", left: 12, top: 14, width: 66, height: 16, set: { Caption: "Customer" } },
  { type: "TextBox", name: "NameBox", left: 84, top: 12, width: 120, height: 20 },
  { type: "ComboBox", name: "RegionPick", left: 84, top: 38, width: 120, height: 20 },
  { type: "ListBox", name: "HistoryList", left: 84, top: 64, width: 120, height: 42 },
  { type: "CheckBox", name: "Taxable", left: 12, top: 40, width: 66, height: 16, set: { Caption: "Taxable" } },
  { type: "Frame", name: "Options", left: 12, top: 112, width: 92, height: 66, set: { Caption: "Freight" } },
  { type: "OptionButton", name: "PickGround", parent: "Options", left: 8, top: 14, width: 76, height: 16, set: { Caption: "Ground" } },
  { type: "OptionButton", name: "PickAir", parent: "Options", left: 8, top: 34, width: 76, height: 16, set: { Caption: "Air" } },
  { type: "ToggleButton", name: "HoldToggle", left: 112, top: 112, width: 92, height: 22, set: { Caption: "Hold" } },
  { type: "MultiPage", name: "Wizard", left: 12, top: 188, width: 192, height: 86 },
  { type: "CheckBox", name: "Agree", parent: "Page1", left: 8, top: 8, width: 100, height: 16, set: { Caption: "Agreed" } },
  { type: "TabStrip", name: "Views", left: 212, top: 188, width: 122, height: 86 },
  { type: "ScrollBar", name: "Amount", left: 212, top: 12, width: 14, height: 96 },
  { type: "SpinButton", name: "Steps", left: 234, top: 12, width: 14, height: 42 },
  { type: "Image", name: "Badge", left: 258, top: 12, width: 76, height: 42 },
  { type: "CommandButton", name: "OkButton", left: 262, top: 250, width: 72, height: 24, set: { Caption: "Start" } },
];

/** The code-behind: handlers naming real controls, so the fixture COMPILES with the form in it. */
export const FORM_CODE = [
  "Option Explicit",
  "",
  "Private Sub UserForm_Initialize()",
  "    RegionPick.AddItem \"North\"",
  "    RegionPick.AddItem \"South\"",
  "    Taxable.Value = True",
  "End Sub",
  "",
  "Private Sub OkButton_Click()",
  "    If Len(NameBox.Text) = 0 Then",
  "        NameBox.SetFocus",
  "        Exit Sub",
  "    End If",
  "    Me.Hide",
  "End Sub",
  "",
].join("\r\n");

/**
 * Builds the whole form through the api against the session's `project`.
 *
 * The name is a parameter because a form name can be REFUSED FOR THE SESSION: observed
 * 2026-08-13, a session in which adding a form named EntryForm to one workbook failed once
 * kept refusing that exact name-kind-project combination quietly ever after, while a module
 * of that name, a form of any other name, and the same form name in another workbook all
 * worked. The identifier-registry lesson (write-rollback's) in a new face: removal does not
 * give a name back. Callers that rebuild repeatedly pass a fresh name each time.
 */
export async function buildForm(api, project, module = FORM_MODULE) {
  await api.component("add", { kind: 3, name: module, project });

  for (const { property, value, as } of FORM_PROPERTIES) {
    await api.designerEdit("set", { module, project, property, value, as });
  }

  for (const control of FORM_CONTROLS) {
    await api.designerEdit("add", {
      module, project, type: control.type.toLowerCase(), name: control.name,
      parent: control.parent, left: control.left, top: control.top,
      width: control.width, height: control.height,
    });
    for (const [property, value] of Object.entries(control.set ?? {})) {
      await api.designerEdit("set", { module, project, name: control.name, property, value, as: "text" });
    }
  }
}
