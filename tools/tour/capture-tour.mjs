/*
 * Stages the README's scenes in a live session and captures each as a PNG still.
 *
 * The pictures are committed; this is how they stay honest. Every scene is staged through the
 * same api the suites drive - real widgets on a real session, no mockups - so reshooting after
 * the surface changes is: build the fixture, run this, run make-tour-gif.mjs.
 *
 *     tools\New-ShowcaseFixture.ps1          # builds the workbook and leaves Excel open on it
 *     node tools\tour\capture-tour.mjs       # all scenes; or name some: editor completion
 *     node tools\tour\make-tour-gif.mjs      # stills -> assets\images\tour.gif
 *
 * Writes assets\images\tour-<scene>.png at exactly FRAME_WIDTH x FRAME_HEIGHT, because
 * PrintWindow captures at the window's own size and the GIF wants every frame equal.
 *
 * The diagnostics scene adds a Validation module full of deliberate findings THROUGH THE API
 * and removes it afterwards, never saving - so the workbook on disk keeps compiling, which the
 * debugger scene depends on. That is also why the debugger scene shoots before diagnostics.
 */

import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { deflateSync } from "node:zlib";
import { open, wait, waitFor } from "../harness/xlide-api.mjs";

const FRAME_WIDTH = 1280;
const FRAME_HEIGHT = 800;
const WORKBOOK = "QuarterlyReport.xlsm";

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const outDir = join(repoRoot, "assets", "images");

/** A paint beat: the api answers from state, the capture reads pixels, and the pixels lag the
 *  state by however long the next frame takes. Nothing here is a test, so a beat is honest. */
const beat = () => wait(400);

// ---------------------------------------------------------------------------- pixels

/** The capture route answers what PrintWindow gives: a bottom-up BGR(A) device bitmap. */
function decodeBmp(bytes) {
  if (bytes.length < 54 || bytes[0] !== 0x42 || bytes[1] !== 0x4d) {
    throw new Error(`not a BMP (${bytes.length} bytes; the route answers JSON on failure: ${bytes.slice(0, 120)})`);
  }
  const pixelStart = bytes.readUInt32LE(10);
  const width = bytes.readInt32LE(18);
  const rawHeight = bytes.readInt32LE(22);
  const bitCount = bytes.readUInt16LE(28);
  const compression = bytes.readUInt32LE(30);
  if ((bitCount !== 24 && bitCount !== 32) || (compression !== 0 && compression !== 3)) {
    throw new Error(`unhandled BMP shape: ${bitCount}bpp compression ${compression}`);
  }

  const height = Math.abs(rawHeight);
  const bottomUp = rawHeight > 0;
  const sourcePixel = bitCount / 8;
  const stride = (width * sourcePixel + 3) & ~3;
  const rgb = Buffer.allocUnsafe(width * height * 3);

  for (let y = 0; y < height; y++) {
    const source = pixelStart + (bottomUp ? height - 1 - y : y) * stride;
    let target = y * width * 3;
    for (let x = 0; x < width; x++) {
      const at = source + x * sourcePixel;
      rgb[target++] = bytes[at + 2];
      rgb[target++] = bytes[at + 1];
      rgb[target++] = bytes[at];
    }
  }

  return { width, height, rgb };
}

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) { c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; }
    table[n] = c;
  }
  return table;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (const byte of bytes) { c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8); }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const chunk = Buffer.allocUnsafe(12 + data.length);
  chunk.writeUInt32BE(data.length, 0);
  chunk.write(type, 4, "latin1");
  data.copy(chunk, 8);
  chunk.writeUInt32BE(crc32(chunk.subarray(4, 8 + data.length)), 8 + data.length);
  return chunk;
}

/** 8-bit RGB, every scanline Up-filtered: on a surface full of vertical runs the row deltas
 *  are mostly zero, which is the kind of input deflate was born for. */
function encodePng(width, height, rgb) {
  const rowBytes = width * 3;
  const raw = Buffer.allocUnsafe((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    const at = y * (rowBytes + 1);
    raw[at] = 2;
    for (let i = 0; i < rowBytes; i++) {
      const here = rgb[y * rowBytes + i];
      const above = y === 0 ? 0 : rgb[(y - 1) * rowBytes + i];
      raw[at + 1 + i] = (here - above) & 0xff;
    }
  }

  const header = Buffer.allocUnsafe(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8;   // bit depth
  header[9] = 2;   // colour type: RGB
  header[10] = 0; header[11] = 0; header[12] = 0;

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", header),
    pngChunk("IDAT", deflateSync(raw, { level: 9 })),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

// ---------------------------------------------------------------------------- staging

/** 1-based line and column of a needle inside module text, so the scenes name code rather
 *  than coordinates and survive edits to the fixture. Case-insensitive, because the host
 *  recases identifiers: the `checked` this file writes comes back `Checked`. */
function placeOf(text, needle) {
  const at = text.toLowerCase().indexOf(needle.toLowerCase());
  if (at < 0) { throw new Error(`the fixture no longer contains "${needle}"`); }
  const before = text.slice(0, at);
  const lastBreak = before.lastIndexOf("\n");
  return { line: before.split("\n").length, column: at - lastBreak };
}

async function snap(api, name) {
  const bmp = await api.capture("frame");
  const { width, height, rgb } = decodeBmp(bmp);
  const target = join(outDir, `tour-${name}.png`);
  writeFileSync(target, encodePng(width, height, rgb));
  console.log(`  ${name}: ${width}x${height} -> ${target}`);
  if (width !== FRAME_WIDTH || height !== FRAME_HEIGHT) {
    console.log(`    NOTE: expected ${FRAME_WIDTH}x${FRAME_HEIGHT}; the frame moved?`);
  }
}

/** The frame at the tour's exact size, restored from whatever it was left at. */
async function sizeFrame(api) {
  const state = await api.state();
  const handle = Number(state.frame);
  const script = join(dirname(fileURLToPath(import.meta.url)), "Set-WindowRect.ps1");
  const moved = spawnSync("powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script,
      "-Handle", String(handle), "-Width", String(FRAME_WIDTH), "-Height", String(FRAME_HEIGHT)],
    { encoding: "utf8" });
  if (moved.status !== 0) {
    throw new Error(`Set-WindowRect failed: ${moved.stderr || moved.stdout}`);
  }
  console.log(`  frame at ${String(moved.stdout).trim()}`);
  await api.placement();
  await beat();
}

/**
 * Puts a module on screen with the caret at a named piece of its code, and does not return
 * until the page AND the host agree that is where things stand.
 *
 * Two lessons are folded in. `act("activate")` needs the full tab identity - a bare module
 * name misses and answers did:false - so the project rides along everywhere here. And a busy
 * session has host echoes in flight (a pane open, a tree selection) that can land AFTER a
 * navigation and yank the active editor back, so the outcome is verified and retried rather
 * than trusted.
 */
async function show(api, module, needle) {
  const text = (await api.readModule(module, WORKBOOK)).text;
  const at = placeOf(text, needle);
  const wanted = `/${module.toLowerCase()}`;

  for (let attempt = 0; attempt < 3; attempt++) {
    await api.act("activate", { module, project: WORKBOOK });
    await api.caret(at.line, { module, column: at.column, project: WORKBOOK });
    try {
      await waitFor(`${module} focused at line ${at.line}`, async () => {
        const focus = (await api.ui()).focus;
        return (focus.model ?? "").toLowerCase().endsWith(wanted)
          && (focus.host?.model ?? "").toLowerCase().endsWith(wanted)
          && focus.line === at.line;
      }, { budgetMs: 4000 });
      return at;
    } catch {
      // An echo moved it; place it again.
    }
  }
  throw new Error(`${module} would not stay focused`);
}

const VALIDATION_CODE = [
  "Option Explicit",
  "",
  "' Sanity checks over the raw export, before the summary is built.",
  "",
  "Public Function RowLooksComplete(ByVal source As Worksheet, ByVal row As Long) As Boolean",
  "    Dim customer As String",
  "    customer = source.Cells(row, 1).Value",
  "    checked = checked + 1",
  "    RowLooksComplete = Len(customer) > 0",
  "End Function",
  "",
  "Public Sub CountBlankRows()",
  "    Dim blanks As Long",
  "    blanks = \"none\"",
  "    Debug.Print blanks",
  "End Sub",
  "",
].join("\r\n");

const SCENES = {
  /** The hero: two modules side by side, the tree unfolded to procedures, problems panel calm. */
  async editor(api) {
    await api.resetLayout();
    await api.act("expandWorkbook", { workbook: WORKBOOK, open: 1 });

    await api.pane("open", { module: "Reporting", project: WORKBOOK });
    await api.pane("open", { module: "Invoice", project: WORKBOOK });

    // Exactly the two: whatever else the session has been up to closes. Nothing here is
    // dirty, so no confirm can stand.
    for (const group of (await api.ui()).workspace.groups) {
      for (const tab of group.tabs) {
        if (tab.module !== "Reporting" && tab.module !== "Invoice") {
          await api.pane("close", { module: tab.module, project: tab.project });
        }
      }
    }

    if ((await api.ui()).workspace.groups.length < 2) {
      await api.act("activate", { module: "Invoice", project: WORKBOOK });
      await api.act("split", { direction: "right" });
    }

    await show(api, "Reporting", "grandTotal = grandTotal + inv.Total");

    // The tree's accordion holds ONE unfolded module, and several of the calls above have
    // already moved it. Unfold is a toggle, so only ask when it is not already Reporting.
    const unfolded = (await api.ui()).explorer.unfolded;
    if ((unfolded?.module ?? "").toLowerCase() !== "reporting") {
      await api.act("unfoldModule", { module: "Reporting" });
    }

    await beat();
    await snap(api, "editor");
  },

  /** The dot menu, open over the Invoice receiver, answered by the analyzer. */
  async completion(api) {
    await show(api, "Reporting", "Total\r\n        counted");
    await api.act("editorAction", { id: "editor.action.triggerSuggest" });
    await api.until("(() => { const w = document.querySelector('.suggest-widget'); return !!w && w.offsetParent !== null; })()", { waitMs: 8000 });
    await beat();
    await snap(api, "completion");
    await api.act("editorAction", { id: "hideSuggestWidget" }).catch(() => {});
  },

  /** Stopped at a breakpoint mid-loop, the Locals panel carrying that iteration's values. */
  async debugger(api) {
    await api.act("dock", { pane: "locals", side: "bottom" });

    const text = (await api.readModule("Reporting", WORKBOOK)).text;
    const stop = placeOf(text, "counted = counted + 1");

    await api.breakpoint("Reporting", stop.line, { project: WORKBOOK, state: "on" });
    await show(api, "Reporting", "Set invoices = LoadInvoices()");
    await api.command("run");
    await waitFor("break mode", async () => {
      if ((await api.breakpoints()).mode === "break") { return true; }
      const standing = (await api.dialogs()).dialogs;
      if (standing.length > 0) {
        throw new Error(`run raised a dialog instead of breaking: ${JSON.stringify(standing)}`);
      }
      return false;
    }, { budgetMs: 25000 });

    // Round the loop twice, so the panel shows totals mid-accumulation rather than zeros.
    await api.command("stepOver");
    await api.command("stepOver");
    await api.command("stepOver");
    await wait(600);

    await snap(api, "debugger");

    await api.command("reset").catch(() => {});
    await api.breakpoint("Reporting", stop.line, { project: WORKBOOK, state: "off" }).catch(() => {});
    await waitFor("design mode", async () => (await api.breakpoints()).mode === "design", { budgetMs: 10000 });
    await api.act("dock", { pane: "problems", side: "bottom" });
  },

  /** Findings as you type, and the quick-fix menu open on one. The module of deliberate
   *  defects arrives through the api and leaves the same way, unsaved. */
  async problems(api) {
    // A failed earlier run can leave the module behind; the scene owns its precondition.
    await api.component("remove", { name: "Validation", project: WORKBOOK }).catch(() => {});
    await api.component("add", { kind: 1, name: "Validation", project: WORKBOOK });
    await api.writeModule("Validation", VALIDATION_CODE, WORKBOOK);
    await api.pane("open", { module: "Validation", project: WORKBOOK });
    await waitFor("the analyzer's findings", async () =>
      ((await api.problems("Validation")).findings ?? []).length >= 2, { budgetMs: 15000 });

    await show(api, "Validation", "checked = checked");
    const fixes = await api.act("quickFixes", { word: "checked" });
    console.log(`  quick fixes standing: ${fixes.detail}`);

    // The sticky keyboard hover, not the quick-fix menu. The code-action menu takes DOM focus
    // when it opens, and in a window that does not hold OS focus - which a scripted shoot's
    // never does - that grab fails and the menu dismisses itself within a few hundred ms. The
    // hover takes no focus and stays; the lightbulb in the gutter tells the fix story beside it.
    await api.act("editorAction", { id: "editor.action.showHover" });
    await api.until("(() => { const w = document.querySelector('.monaco-hover'); return !!w && w.offsetParent !== null; })()", { waitMs: 8000 });
    await beat();
    await snap(api, "problems");

    await api.act("key", { code: "Escape", target: "document" }).catch(() => {});
    await api.pane("close", { module: "Validation", project: WORKBOOK }).catch(() => {});
    await api.component("remove", { name: "Validation", project: WORKBOOK });
  },
};

// ---------------------------------------------------------------------------- run

const asked = process.argv.slice(2);
const unknown = asked.filter((name) => !(name in SCENES));
if (unknown.length > 0) {
  console.error(`no scene named ${unknown.join(", ")}; there are: ${Object.keys(SCENES).join(", ")}`);
  process.exit(2);
}

const api = await open({ workbook: WORKBOOK });
mkdirSync(outDir, { recursive: true });

// A failed problems scene leaves Validation standing, and its findings would dirty every
// other scene's problems panel. Sweep it before anything shoots.
await api.component("remove", { name: "Validation", project: WORKBOOK }).catch(() => {});
await sizeFrame(api);

for (const name of asked.length > 0 ? asked : Object.keys(SCENES)) {
  console.log(`scene: ${name}`);
  await SCENES[name](api);
}

console.log("done");
