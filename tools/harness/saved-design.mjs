/*
 * What the developer CHANGED, read out of the workbook Excel has already saved.
 *
 * This is the instrument for #53 (docs/userform-designer.md: every changed property rides the
 * document). MSForms persists only properties that differ from the file format default, and it
 * says which ones in a PropMask bitfield at a fixed offset in every control's block. So the mask
 * alone answers "what did the developer touch" - the VALUES still come from the live object
 * model, and none of them need decoding here.
 *
 * No export, no COM, no Excel. The road, all of it grounded in [MS-OFORMS] rather than inspection:
 *
 *   the .xlsm is a ZIP        ->  xl/vbaProject.bin
 *   vbaProject.bin is a CFB   ->  /<Form>/f   FormControl        (2.2.10.1)
 *                                 /<Form>/o   the controls' own blocks, in site order
 *                                 /<Form>/iNN a container's own storage, NN being its site ID
 *
 * A form's `f` stream is a FormControl: MinorVersion(1) MajorVersion(1) cbForm(2) then cbForm
 * bytes of PropMask + DataBlock + ExtraDataBlock, then StreamData, then FormSiteData. The site
 * array holds one OleSiteConcreteControl per control - Version(2) cbSite(2) PropMask(4) and two
 * blocks - which carries the control's Name, its ClsidCacheIndex (what KIND it is) and its
 * ObjectStreamSize (how much of `o` belongs to it). Walking `o` by those sizes lands on each
 * control's own PropMask, four bytes in.
 *
 * Verified against FormFixture.xlsm, 2026-08-17: every site record consumes exactly its cbSite,
 * the sum of ObjectStreamSize equals the `o` stream to the byte, all fifteen cache indexes match
 * the fixture's actual control kinds, and the masks decode to what the fixture was built with -
 * a Label to Caption+Size, the OK button to Caption+PicturePosition+Size+Picture, the Image to
 * PictureSizeMode+Size+Picture.
 *
 * WHAT A SET BIT ACTUALLY MEANS, which is the finding that decides how this gets used. The mask
 * says a property differs from the FILE FORMAT default - not that the developer chose it. Where a
 * control KIND is born with something other than the file's default, the bit is set on controls
 * nobody touched. Measured on the fixture: every control carries FontName, because the form is
 * Tahoma and the file's default is MS Sans Serif; and every CheckBox, OptionButton and
 * ToggleButton carries BackColor and ForeColor, though form-plan.mjs sets nothing but a Caption
 * on any of them.
 *
 * So this does not replace the walk's existing comparison against a bare coclass - it NARROWS it.
 * The mask is the short list of properties that could possibly be non-default, read once per form
 * and cached; the walk then asks only those and compares as it already does. Reading fifty
 * properties of every control on every projection is the cost that comparison was avoiding, and
 * this is how it stops having to.
 *
 * WHAT THIS CANNOT SAY. VariousPropertyBits is one mask bit over a packed field holding Enabled,
 * Locked, Visible, AutoSize, WordWrap and more. A set bit means one of them changed, not which,
 * and telling them apart means decoding the DataBlock rather than the mask. It is reported as
 * itself, and the caller decides.
 */
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { readCfb } from "./vba-storage.mjs";

/* ---- the mask tables, [MS-OFORMS] 2.2.x, one entry per bit that names a property -------------
 *
 * Bit numbering is the spec's own: the lettered fields run from bit 0 upwards, and a multi-bit
 * unused run shifts everything after it (FormPropMask's Unused2 is two bits wide, which is why
 * fBooleanProperties is bit 6 rather than bit 5).
 *
 * STRUCTURAL names are the ones a developer never sets: bookkeeping, the MUST-be-1 bits, and the
 * fields that carry a container's own children. They are read and then dropped, because a
 * property list that always says "Size" says nothing.
 */
const STRUCTURAL = new Set([
  "Size", "NewVersion", "Reserved", "DrawBuffer", "NextAvailableID", "ShapeCookie",
  "GroupCnt", "LogicalSize", "DisplayedSize", "ScrollPosition", "ObjectStreamSize",
  "ID", "ClsidCacheIndex", "BitFlags", "Position", "GroupID", "Name", "DisplayStyle",
  "cColumnInfo", "TabsAllocated", "TabData", "Items", "Names", "Tags", "TipStrings",
  "Accelerators", "ListIndex", "PrevEnabled", "NextEnabled", "RuntimeLicKey",
  // Set on every control but the first, so it distinguishes nothing - and the projection
  // already carries tabIndex on its own row for the tab-order dialog to read.
  "TabIndex",
]);

const FORM = [
  [1, "BackColor"], [2, "ForeColor"], [3, "NextAvailableID"], [6, "BooleanProperties"],
  [7, "BorderStyle"], [8, "MousePointer"], [9, "ScrollBars"], [10, "DisplayedSize"],
  [11, "LogicalSize"], [12, "ScrollPosition"], [13, "GroupCnt"], [15, "MouseIcon"],
  [16, "Cycle"], [17, "SpecialEffect"], [18, "BorderColor"], [19, "Caption"], [20, "Font"],
  [21, "Picture"], [22, "Zoom"], [23, "PictureAlignment"], [24, "PictureTiling"],
  [25, "PictureSizeMode"], [26, "ShapeCookie"], [27, "DrawBuffer"],
];

const SITE = [
  [0, "Name"], [1, "Tag"], [2, "ID"], [3, "HelpContextID"], [4, "BitFlags"],
  [5, "ObjectStreamSize"], [6, "TabIndex"], [7, "ClsidCacheIndex"], [8, "Position"],
  [9, "GroupID"], [11, "ControlTipText"], [12, "RuntimeLicKey"], [13, "ControlSource"],
  [14, "RowSource"],
];

const MORPH_DATA = [
  [0, "VariousPropertyBits"], [1, "BackColor"], [2, "ForeColor"], [3, "MaxLength"],
  [4, "BorderStyle"], [5, "ScrollBars"], [6, "DisplayStyle"], [7, "MousePointer"],
  [8, "Size"], [9, "PasswordChar"], [10, "ListWidth"], [11, "BoundColumn"],
  [12, "TextColumn"], [13, "ColumnCount"], [14, "ListRows"], [15, "cColumnInfo"],
  [16, "MatchEntry"], [17, "ListStyle"], [18, "ShowDropButtonWhen"], [20, "DropButtonStyle"],
  [21, "MultiSelect"], [22, "Value"], [23, "Caption"], [24, "PicturePosition"],
  [25, "BorderColor"], [26, "SpecialEffect"], [27, "MouseIcon"], [28, "Picture"],
  [29, "Accelerator"], [31, "Reserved"], [32, "GroupName"],
];

const LABEL = [
  [0, "ForeColor"], [1, "BackColor"], [2, "VariousPropertyBits"], [3, "Caption"],
  [4, "PicturePosition"], [5, "Size"], [6, "MousePointer"], [7, "BorderColor"],
  [8, "BorderStyle"], [9, "SpecialEffect"], [10, "Picture"], [11, "Accelerator"],
  [12, "MouseIcon"],
];

const COMMAND_BUTTON = [
  [0, "ForeColor"], [1, "BackColor"], [2, "VariousPropertyBits"], [3, "Caption"],
  [4, "PicturePosition"], [5, "Size"], [6, "MousePointer"], [7, "Picture"],
  [8, "Accelerator"], [9, "TakeFocusOnClick"], [10, "MouseIcon"],
];

const IMAGE = [
  [2, "AutoSize"], [3, "BorderColor"], [4, "BackColor"], [5, "BorderStyle"],
  [6, "MousePointer"], [7, "PictureSizeMode"], [8, "SpecialEffect"], [9, "Size"],
  [10, "Picture"], [11, "PictureAlignment"], [12, "PictureTiling"],
  [13, "VariousPropertyBits"], [14, "MouseIcon"],
];

const TAB_STRIP = [
  [0, "ListIndex"], [1, "BackColor"], [2, "ForeColor"], [4, "Size"], [5, "Items"],
  [6, "MousePointer"], [8, "TabOrientation"], [9, "TabStyle"], [10, "MultiRow"],
  [11, "TabFixedWidth"], [12, "TabFixedHeight"], [13, "Tooltips"], [15, "TipStrings"],
  [17, "Names"], [18, "VariousPropertyBits"], [19, "NewVersion"], [20, "TabsAllocated"],
  [21, "Tags"], [22, "TabData"], [23, "Accelerators"], [24, "MouseIcon"],
];

const SCROLL_BAR = [
  [0, "ForeColor"], [1, "BackColor"], [2, "VariousPropertyBits"], [3, "Size"],
  [4, "MousePointer"], [5, "Min"], [6, "Max"], [7, "Position"], [9, "PrevEnabled"],
  [10, "NextEnabled"], [11, "SmallChange"], [12, "LargeChange"], [13, "Orientation"],
  [14, "ProportionalThumb"], [15, "Delay"], [16, "MouseIcon"],
];

const SPIN_BUTTON = [
  [0, "ForeColor"], [1, "BackColor"], [2, "VariousPropertyBits"], [3, "Size"],
  [5, "Min"], [6, "Max"], [7, "Position"], [8, "PrevEnabled"], [9, "NextEnabled"],
  [10, "SmallChange"], [11, "Orientation"], [12, "Delay"], [13, "MouseIcon"],
  [14, "MousePointer"],
];

/* TextProps rides at the END of a control's block and carries the font. Its own mask is four
 * bytes after a two-byte version and a two-byte size, exactly like a control's. */
const TEXT_PROPS = [
  [0, "FontName"], [1, "FontEffects"], [2, "FontHeight"], [4, "FontCharSet"],
  [5, "FontPitchAndFamily"], [6, "ParagraphAlign"], [7, "FontWeight"],
];

/* [MS-OFORMS] 2.4.2 FormEmbeddedActiveXControlCached: what a ClsidCacheIndex means. */
const KINDS = new Map([
  [7, "Form"], [12, "Image"], [14, "Frame"], [15, "MorphData"], [16, "SpinButton"],
  [17, "CommandButton"], [18, "TabStrip"], [21, "Label"], [23, "TextBox"], [24, "ListBox"],
  [25, "ComboBox"], [26, "CheckBox"], [27, "OptionButton"], [28, "ToggleButton"],
  [47, "ScrollBar"], [57, "MultiPage"],
]);

/* Six kinds share one structure and therefore one mask, which is eight bytes wide rather than
 * four - the only kind that is. */
const MASKS = new Map([
  ["Image", { bits: IMAGE, wide: false, text: false }],
  ["SpinButton", { bits: SPIN_BUTTON, wide: false, text: false }],
  ["CommandButton", { bits: COMMAND_BUTTON, wide: false, text: true }],
  ["TabStrip", { bits: TAB_STRIP, wide: false, text: true }],
  ["Label", { bits: LABEL, wide: false, text: true }],
  ["ScrollBar", { bits: SCROLL_BAR, wide: false, text: false }],
  ["TextBox", { bits: MORPH_DATA, wide: true, text: true }],
  ["ListBox", { bits: MORPH_DATA, wide: true, text: true }],
  ["ComboBox", { bits: MORPH_DATA, wide: true, text: true }],
  ["CheckBox", { bits: MORPH_DATA, wide: true, text: true }],
  ["OptionButton", { bits: MORPH_DATA, wide: true, text: true }],
  ["ToggleButton", { bits: MORPH_DATA, wide: true, text: true }],
]);

const CONTAINERS = new Set(["Form", "Frame", "MultiPage"]);

const namesIn = (mask, bits) => bits
  .filter(([bit]) => (mask & (1n << BigInt(bit))) !== 0n)
  .map(([, name]) => name)
  .filter((name) => !STRUCTURAL.has(name));

/* ---- the workbook, unopened ---------------------------------------------------------------- */

/** xl/vbaProject.bin out of an .xlsm, by walking the ZIP's central directory. */
export function vbaProjectOf(workbookPath) {
  const zip = readFileSync(workbookPath);
  let end = zip.length - 22;
  while (end >= 0 && zip.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < 0) throw new Error(`not a zip: ${workbookPath}`);

  let at = zip.readUInt32LE(end + 16);
  const count = zip.readUInt16LE(end + 10);
  for (let i = 0; i < count; i++) {
    const nameLength = zip.readUInt16LE(at + 28);
    const name = zip.subarray(at + 46, at + 46 + nameLength).toString("latin1");
    if (name === "xl/vbaProject.bin") {
      const local = zip.readUInt32LE(at + 42);
      const method = zip.readUInt16LE(at + 10);
      const compressed = zip.readUInt32LE(at + 20);
      const body = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28);
      const bytes = zip.subarray(body, body + compressed);
      return method === 0 ? bytes : inflateRawSync(bytes);
    }
    at += 46 + nameLength + zip.readUInt16LE(at + 30) + zip.readUInt16LE(at + 32);
  }
  throw new Error(`no VBA project in ${workbookPath}`);
}

/* ---- one site record, [MS-OFORMS] 2.2.10 --------------------------------------------------- */

/**
 * Alignment inside a site is measured from the START of the record - which is what the spec's
 * "from the beginning of the version number" means - so every read goes through `align`.
 */
function readSite(buf, start) {
  const cbSite = buf.readUInt16LE(start + 2);
  const mask = buf.readUInt32LE(start + 4);
  let p = start + 8;
  const has = (bit) => (mask & (1 << bit)) !== 0;
  const align = (n) => { const over = (p - start) % n; if (over) p += n - over; };
  const u32 = () => { align(4); const v = buf.readUInt32LE(p); p += 4; return v; };
  const u16 = () => { align(2); const v = buf.readUInt16LE(p); p += 2; return v; };

  const nameData = has(0) ? u32() : 0;
  const tagData = has(1) ? u32() : 0;
  const id = has(2) ? u32() : 0;
  if (has(3)) u32();                                   // HelpContextID
  if (has(4)) u32();                                   // BitFlags
  const objectStreamSize = has(5) ? u32() : 0;
  if (has(6)) u16();                                   // TabIndex
  const clsidCacheIndex = has(7) ? u16() : null;
  if (has(9)) u16();                                   // GroupID
  const tipData = has(11) ? u32() : 0;
  const licData = has(12) ? u32() : 0;
  const sourceData = has(13) ? u32() : 0;
  const rowData = has(14) ? u32() : 0;

  { const over = (p - (start + 8)) % 4; if (over) p += 4 - over; }   // the DataBlock's own tail

  const str = (data) => {
    if (!data) return null;
    const cb = data & 0x7fffffff;
    const text = buf.subarray(p, p + cb).toString((data >>> 31) === 1 ? "latin1" : "utf16le");
    p += cb + ((4 - (cb % 4)) % 4);                    // strings pad to a multiple of four
    return text;
  };

  const name = str(nameData);
  str(tagData);
  if (has(8)) p += 8;                                  // SitePosition
  str(tipData);
  str(licData);
  str(sourceData);
  str(rowData);

  return {
    name, id, objectStreamSize, clsidCacheIndex,
    kind: KINDS.get(clsidCacheIndex ?? -1) ?? null,
    changed: namesIn(BigInt(mask), SITE),
    length: 4 + cbSite,
  };
}

/* ---- a form's own stream, and the sites under it -------------------------------------------- */

function readSites(f) {
  const cbForm = f.readUInt16LE(2);
  const formMask = BigInt(f.readUInt32LE(4));

  // StreamData holds the form's mouse icon, font and picture as GUID-prefixed blobs. Rather than
  // decode three OLE structures to skip them, find FormSiteData by its own arithmetic: the pair
  // (CountOfSites, CountOfBytes) whose byte count exactly covers the sites that follow.
  let at = -1;
  for (let probe = 4 + cbForm; probe + 8 < f.length; probe++) {
    const count = f.readUInt32LE(probe);
    const bytes = f.readUInt32LE(probe + 4);
    if (count > 0 && count < 1000 && probe + 8 + bytes <= f.length) { at = probe; break; }
  }
  if (at < 0) return { form: namesIn(formMask, FORM), sites: [] };

  const count = f.readUInt32LE(at);
  let p = at + 8;

  // SiteDepthsAndTypes: one FormObjectDepthTypeCount per run, two bytes, or three when the high
  // bit of the second says the byte after it is a type shared by that many consecutive sites.
  const start = p;
  for (let seen = 0; seen < count;) {
    const typeOrCount = f.readUInt8(p + 1);
    if ((typeOrCount & 0x80) !== 0) { seen += typeOrCount & 0x7f; p += 3; } else { seen += 1; p += 2; }
  }
  p += (4 - ((p - start) % 4)) % 4;                    // ArrayPadding

  const sites = [];
  for (let i = 0; i < count; i++) {
    const site = readSite(f, p);
    sites.push(site);
    p += site.length;
  }
  return { form: namesIn(formMask, FORM), sites };
}

/**
 * Every control under one storage: its own PropMask out of the `o` stream, plus the site-level
 * properties out of `f`, plus the font out of the TextProps that trails the block.
 */
function readStorage(cfb, prefix, into, path) {
  const f = cfb.paths.get(`${prefix}/f`);
  if (!f) return;
  const { form, sites } = readSites(cfb.readStream(f));

  const o = cfb.paths.get(`${prefix}/o`);
  const blocks = o ? cfb.readStream(o) : Buffer.alloc(0);

  // A container is described TWICE - once as a site in its parent, once by the FormPropMask in
  // its own storage - so the two are merged rather than the second replacing the first.
  const record = (where, names) => {
    if (!names.length) return;
    const all = new Set(into.get(where) ?? []);
    for (const name of names) all.add(name);
    into.set(where, [...all].sort());
  };

  if (path.length) record(path.join("."), form);

  let at = 0;
  for (const site of sites) {
    // A MultiPage keeps its own TabStrip as a nameless site. It is the container's chrome rather
    // than a control the developer put there, and no name means nothing can ask about it.
    if (!site.name) { at += site.objectStreamSize; continue; }

    const where = [...path, site.name].join(".");
    const changed = new Set(site.changed);

    if (site.objectStreamSize && MASKS.has(site.kind)) {
      const { bits, wide, text } = MASKS.get(site.kind);
      const cb = blocks.readUInt16LE(at + 2);
      const mask = wide ? blocks.readBigUInt64LE(at + 4) : BigInt(blocks.readUInt32LE(at + 4));
      for (const name of namesIn(mask, bits)) changed.add(name);

      // TextProps sits after PropMask + DataBlock + ExtraDataBlock + StreamData. Only reach for
      // it when nothing lies between: a picture's bytes are unmeasured here, and a font read at
      // the wrong offset is worse than an unread one.
      if (text && !changed.has("Picture") && !changed.has("MouseIcon")) {
        const props = at + 4 + cb;
        if (props + 8 <= at + site.objectStreamSize && blocks.readUInt8(props + 1) === 2) {
          for (const name of namesIn(BigInt(blocks.readUInt32LE(props + 4)), TEXT_PROPS)) {
            changed.add(name);
          }
        }
      }
    }

    record(where, [...changed]);
    at += site.objectStreamSize;

    if (CONTAINERS.has(site.kind ?? "")) {
      readStorage(cfb, `${prefix}/i${String(site.id).padStart(2, "0")}`, into,
        [...path, site.name]);
    }
  }
}

/**
 * Every form in a saved workbook, and for each one a map from control path to the properties
 * MSForms recorded as changed. A control absent from the map has nothing recorded, which is a
 * real answer rather than a missing one.
 */
export function readSavedDesign(workbookPath) {
  const cfb = readCfb(vbaProjectOf(workbookPath));
  const forms = new Map();
  for (const [path, entry] of cfb.paths) {
    if (entry.type !== 1 || path.lastIndexOf("/") !== 0 || path === "/VBA") continue;
    if (!cfb.paths.has(`${path}/f`)) continue;
    const changed = new Map();
    readStorage(cfb, path, changed, []);
    forms.set(path.slice(1), changed);
  }
  return forms;
}

if (process.argv[1] && process.argv[1].endsWith("saved-design.mjs")) {
  const target = process.argv[2];
  if (!target) {
    console.error("usage: node saved-design.mjs <workbook.xlsm>");
    process.exit(2);
  }
  for (const [form, changed] of readSavedDesign(target)) {
    console.log(`\n${form}`);
    for (const [where, names] of changed) {
      console.log(`  ${where.padEnd(26)} ${names.join(", ")}`);
    }
  }
}
