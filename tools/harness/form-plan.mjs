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

import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

export const FORM_MODULE = "EntryForm";

/**
 * The two pictures the fixture wears, as absolute paths - which is what the designer route
 * takes, because MSForms loads a picture from a FILE and remembers only the pixels.
 *
 * Both are the product's OWN artwork, already in the repository: nothing to generate, nothing to
 * clean up, and a fixture that looks like something rather than a coloured square.
 *
 * BOTH ARRIVE AS BITMAPS, measured 2026-08-16, and the ICO does too - which is worth writing
 * down because the name says otherwise. OLE's own picture loader reads classic icons, whose
 * frames are BMPs; every frame of this one is PNG-compressed, which OLE has never understood, so
 * it goes down the GDI+ road with the PNG and comes back as the flat bitmap GDI+ decodes. The
 * ICON road through DrawIconEx is real and is covered where it can be produced on purpose:
 * designer-features builds a classic icon in the temp folder for exactly that.
 */
const REPO = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
export const FORM_PICTURES = {
  bitmap: join(REPO, "assets", "images", "extension_logo.png"),
  icon: join(REPO, "assets", "xlide.ico"),
};

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
  // SHORTER than the MultiPage beside it, and deliberately: at 86 points it reached y=274 and the
  // OK button at 250 sat on its face, which reads as a button inside a tab that does not change
  // when the tab does (the owner, 2026-08-16: "there is a button drawn over it, but changing the
  // tabs doesn't visually do anything"). Nothing was wrong with the control; the fixture was
  // telling a lie about it.
  { type: "TabStrip", name: "Views", left: 212, top: 188, width: 122, height: 56 },
  // ON the TabStrip's face, and a FORM control all the same: a TabStrip is not a container in
  // MSForms - it draws tabs over ONE set of controls, which the code swaps as the tab changes
  // (Views_Change below). Without this the strip was an empty box, which is the other half of
  // the same misreading.
  // The caption SAYS what the control is for, because the design surface cannot show it: a
  // TabStrip's face holds one set of controls and only the code swaps what they say, so at
  // design time both tabs look the same and that is the truth rather than a fault (the owner,
  // 2026-08-16: "both tabs show 'summary view'"). Run the form and this reads "Tab1 view" or
  // "Tab2 view" as the tabs are clicked.
  { type: "Label", name: "ViewNote", left: 220, top: 210, width: 106, height: 16, set: { Caption: "Views_Change fills this" } },
  { type: "ScrollBar", name: "Amount", left: 212, top: 12, width: 14, height: 96 },
  { type: "SpinButton", name: "Steps", left: 234, top: 12, width: 14, height: 42 },
  // WEARING A PICTURE, since 2026-08-16: an Image control that holds nothing is an Image control
  // whose whole reason for existing is untested, and the canvas drew a crossed box for every one
  // of them. Zoomed (PictureSizeMode 3) so the logo fits the 76x42 box whatever its own size.
  {
    type: "Image", name: "Badge", left: 258, top: 12, width: 76, height: 42,
    picture: { file: "bitmap", place: { PictureSizeMode: 3 } },
  },
  // And the other half: a picture BESIDE a caption rather than as the whole face. Position 1 is
  // fmPicturePositionLeftCenter - the logo left of "Start".
  {
    type: "CommandButton", name: "OkButton", left: 262, top: 250, width: 72, height: 24,
    picture: { file: "icon", place: { PicturePosition: 1 } }, set: { Caption: "Start" },
  },
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
  "Private Sub Views_Change()",
  "    ViewNote.Caption = Views.SelectedItem.Caption & \" view\"",
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
    // The picture BEFORE its placement: PictureSizeMode on a control holding nothing is a
    // setting about nothing, and the panel only offers those rows once there is a picture.
    if (control.picture) {
      await api.designerEdit("set", {
        module, project, name: control.name,
        property: "Picture", value: FORM_PICTURES[control.picture.file],
      });
      for (const [property, value] of Object.entries(control.picture.place ?? {})) {
        await api.designerEdit("set", { module, project, name: control.name, property, value: String(value) });
      }
    }

    for (const [property, value] of Object.entries(control.set ?? {})) {
      await api.designerEdit("set", { module, project, name: control.name, property, value, as: "text" });
    }
  }
}
