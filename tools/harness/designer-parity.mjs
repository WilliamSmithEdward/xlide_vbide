/*
 * The canvas against the RUNNING FORM, control by control.
 *
 * WHY THIS EXISTS. Every other designer check compares the canvas with the MODEL - the document,
 * the designer's collection - and the canvas has always agreed with the model to the point. What
 * nothing compared was the canvas with the thing the developer actually sees, and the two can
 * differ while both are "right": a Frame at top 112 has its rectangle at 112 in the model and
 * draws its rule four points lower on screen, because MSForms keeps a caption band above the
 * line. The canvas drew the rule at 112, so a button placed level with the frame looked level in
 * xlide and sat four points high in the form (the owner's side-by-side, 2026-08-16: "the xlide
 * designer doesn't accurately represent how the controls align with each other in actual
 * runtime", then "please validate geometry parity of all controls").
 *
 * HOW IT WORKS. The form is launched from its designer tab - which applies and saves, so both
 * surfaces are showing the same document - and photographed through `capture?window=form`, the
 * only picture of a running form anything can take: MSForms draws its controls windowless, so
 * there are no child handles to enumerate. The canvas is read from the DOM, which gives exact
 * numbers rather than pixels. Both are reduced to POINTS from the form's client origin and
 * compared landmark by landmark.
 *
 * WHAT A LANDMARK IS. The top edge each kind actually paints: a button's bevel, a text box's
 * border, a frame's rule, a multipage's body. Not the model's rectangle, which is what both
 * surfaces already agree about.
 *
 * It PRINTS rather than passing or failing, like the perf walk: what counts as a parity defect
 * is a judgement about what a person can see, and half a point is not the same finding as four.
 *
 *   node tools\harness\designer-parity.mjs
 */
import { open, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const FORM = "EntryForm";
const CAPTION = "Quarter Entry";

/** An answer that may already have been parsed for us, as an object either way. */
const asObject = (answer) => typeof answer === "string" ? JSON.parse(answer) : answer;

/** Device pixels per point in the capture, worked out from the form's own width. */
let scale = 2;

/** Where the form's client area begins in the capture, in device pixels. */
let originX = 0;
let originY = 0;

/**
 * The capture, decoded.
 *
 * The shim hands over a BMP because a bitmap is what PrintWindow gives, and forty bytes of
 * header plus a bottom-up array of BGRA is less trouble than an image dependency or a
 * PowerShell round trip per column - which this probe reached for first, and which the
 * machine's execution policy declined.
 */
function decodeBmp(bytes) {
  if (bytes.length < 54 || bytes[0] !== 0x42 || bytes[1] !== 0x4D) {
    throw new Error(`the capture is not a BMP: ${bytes.subarray(0, 80).toString("utf8")}`);
  }

  const offset = bytes.readUInt32LE(10);
  const width = bytes.readInt32LE(18);
  const height = bytes.readInt32LE(22);
  const bits = bytes.readUInt16LE(28);
  if (bits !== 32) {
    throw new Error(`the capture is ${bits} bits per pixel; this reads 32`);
  }

  const rows = Math.abs(height);
  const stride = width * 4;
  const upward = height > 0;

  return {
    width,
    height: rows,
    /** Perceived brightness at a pixel, 0 to 255. */
    lum(x, y) {
      const row = upward ? rows - 1 - y : y;
      const at = offset + row * stride + x * 4;
      return 0.299 * bytes[at + 2] + 0.587 * bytes[at + 1] + 0.114 * bytes[at];
    },
  };
}

/** The luminance profile down one column, as `{y, lum}` rows. */
function column(image, x, from, to) {
  const out = [];
  for (let y = Math.max(0, from); y <= Math.min(to, image.height - 1); y++) {
    out.push({ y, lum: image.lum(x, y) });
  }

  return out;
}

/**
 * The first row in the column that is not the form's ground - the top edge of whatever is
 * painted there. Both directions count: a raised bevel is BRIGHTER than the ground and a
 * border is darker, and an edge detector that only looks for dark misses every button.
 */
function edgeIn(profile, ground = 240, tolerance = 8) {
  for (const row of profile) {
    if (Math.abs(row.lum - ground) > tolerance) {
      return row.y;
    }
  }

  return null;
}

const asPoints = (y) => (y - originY) / scale;

console.log(`\nthe canvas against the running form, ${project.projectId.split(/[\\/]/).pop()}\n`);

try {
  // A form left standing by an earlier run holds its own designer shut, and the read below is
  // the first thing that needs it. Cheap to insist on a clean start rather than explain later.
  if (((await api.userforms()).forms ?? []).length > 0) {
    await api.userforms("close");
    await waitFor("the form somebody left standing to go", async () =>
      ((await api.userforms()).forms ?? []).length === 0, { budgetMs: 15000 });
  }

  await api.pane("open", { module: FORM, face: "design", project: project.projectId });
  await waitFor("the designer tab to stand", async () =>
    ((await api.ui()).workspace?.groups ?? []).flatMap((group) => group.tabs)
      .some((tab) => tab.module === FORM && tab.face === "design"), { budgetMs: 20000 });

  // What the canvas PAINTS, in points from the form's client origin. The landmark is the top
  // of the element that carries the visible edge, which for a container is not its box.
  //
  // The door parses an answer that looks like JSON, so this takes either shape rather than
  // assuming one - a string here and an object there is exactly the difference that turns a
  // working probe into `[object Object]` at the first parse.
  const painted = asObject(await api.ask(`(() => {
    const view = document.querySelector('.designer-view[data-module=${JSON.stringify(FORM)}]');
    const client = view.querySelector('.dc-form-client') || view.querySelector('.dc-form');
    const cb = client.getBoundingClientRect();
    const PT = 4 / 3;
    const out = {};
    for (const el of view.querySelectorAll('.dc[data-control]')) {
      const kind = el.dataset.kind;
      // The landmark is what the kind actually PAINTS at its top: a container's rule, a
      // tick box's or a radio's glyph, and the box itself for everything else.
      const edge = el.querySelector('.dc-frame-rule, .dc-page-rule, .dc-glyph-box, .dc-glyph-dot') || el;
      const r = edge.getBoundingClientRect();
      const box = el.getBoundingClientRect();
      out[el.dataset.control] = {
        kind,
        top: +((r.y - cb.y) / PT).toFixed(2),
        left: +((box.x - cb.x) / PT).toFixed(2),
        boxTop: +((box.y - cb.y) / PT).toFixed(2),
        width: +(box.width / PT).toFixed(2),
        bottom: +((box.bottom - cb.y) / PT).toFixed(2),
      };
    }
    return JSON.stringify(out);
  })()`));

  // The model, read BEFORE the launch: a loaded form has no designer to ask (the object goes
  // with the run and comes back a few hundred milliseconds after the close), and this is the
  // list every column below is aimed by.
  const form = await api.designer(FORM, project.projectId);

  // The form, launched from the tab - which applies and saves, so the picture below is of the
  // same document the canvas is drawing.
  const running = api.command("run");
  await waitFor("the form to stand", async () =>
    ((await api.userforms()).forms ?? []).some((title) => title.includes(CAPTION)),
    { budgetMs: 25000 });
  await running;

  const image = decodeBmp(await api.capture("form", CAPTION));

  // The capture is the whole window: the client's width sets the scale, and the topmost
  // control sets the vertical origin. Both are found from the image rather than assumed,
  // because a window border is a theme's business and the scale is the monitor's.
  scale = (image.width - 2) / form.form.width;
  originX = 1;

  // Only a control the FORM owns can calibrate: a container's child carries coordinates
  // relative to its parent's client, and the first cut of this took a checkbox inside a
  // MultiPage page for the topmost thing on the form, which put the origin 40 pixels wrong
  // and made every reading below it look like a defect.
  const highest = form.controls
    .filter((one) => one.top !== null && one.parent?.toLowerCase() === FORM.toLowerCase())
    .reduce((best, one) => (best === null || one.top < best.top ? one : best), null);
  const guessAt = Math.round(originX + (highest.left + highest.width / 2) * scale);
  const found = edgeIn(column(image, guessAt, 20, Math.round(image.height * 0.4)));
  originY = found === null ? 46 : found - highest.top * scale;

  console.log(`capture ${image.width}x${image.height}, ${scale.toFixed(3)} px per point, `
    + `client origin at ${originX},${originY.toFixed(1)} (from ${highest.name})\n`);

  console.log("control          kind            in       canvas   runtime   delta");
  console.log("-".repeat(70));

  const findings = [];
  for (const control of form.controls) {
    if (control.top === null || control.left === null || control.width < 6) {
      continue;
    }

    const drawn = painted[control.name];
    if (!drawn) {
      continue;
    }

    /*
     * THE AIM COMES FROM THE CANVAS, NOT FROM THE MODEL.
     *
     * A container's child carries coordinates relative to its parent's CLIENT, and where that
     * client begins is the very thing under test - so resolving a child's absolute position
     * from the model would bake the answer into the question. The canvas already knows where
     * it painted the thing; the scan starts there and looks a few points either way for what
     * the runtime painted. A canvas four points out is still inside the window, which is what
     * makes the reading a measurement rather than a confirmation.
     */
    // A CONTAINER is scanned at its far end, past its own caption and tabs: aim at a frame's
    // left and the column runs through "Freight", aim at a multipage's and it runs through
    // Page1, and what comes back is the top of the lettering rather than the rectangle. The
    // rule is the landmark for these kinds, so the column has to be somewhere the rule is the
    // only thing there.
    const container = drawn.kind === "Frame" || drawn.kind === "MultiPage" || drawn.kind === "TabStrip";
    const aimAt = container
      ? drawn.left + drawn.width - Math.max(8, drawn.width * 0.15)
      : drawn.left + Math.min(control.width, 40) * 0.35;
    const x = Math.round(originX + aimAt * scale);

    // Not a point above the nearest thing standing over this column. A window that reaches
    // eight points up runs into whatever is above - a list box's white interior, a frame's
    // rule - and the first reading it takes is that, not this control's edge. Every reading
    // in the first cut of this table was the top of the window it was given.
    const ceiling = Object.values(painted)
      .filter((other) => other !== drawn && other.bottom <= drawn.boxTop
        && aimAt >= other.left && aimAt <= other.left + other.width)
      .reduce((low, other) => Math.max(low, other.bottom), drawn.boxTop - 8);

    // A container's landmark is well down inside its box - past the caption band or the tab
    // strip - so the window has to reach that far or the scan comes back empty.
    const from = Math.round(originY + (ceiling + 0.5) * scale);
    const to = Math.round(originY + (drawn.boxTop + (container ? 26 : 10)) * scale);
    const at = edgeIn(column(image, x, Math.max(0, from), to));
    const runtime = at === null ? null : +asPoints(at).toFixed(2);
    const delta = runtime === null ? null : +(drawn.top - runtime).toFixed(2);
    const owner = control.parent?.toLowerCase() === FORM.toLowerCase() ? "form" : control.parent;

    // A Label paints no edge at all and a scroll bar's button face is the ground's own colour,
    // so nothing is there to find. Said plainly, because a dash in a table reads as a failure.
    console.log(control.name.padEnd(16) + String(control.type ?? drawn.kind).padEnd(14)
      + String(owner).padEnd(10) + String(drawn.top).padStart(7)
      + String(runtime ?? "no edge").padStart(10) + String(delta ?? "-").padStart(8));

    if (delta !== null && Math.abs(delta) >= 1) {
      findings.push(`${control.name} (${drawn.kind}): the canvas paints its top edge `
        + `${delta > 0 ? `${delta}pt lower` : `${-delta}pt higher`} than the form does`);
    }
  }

  console.log();
  if (findings.length === 0) {
    console.log("every control's painted top edge is within a point of the running form's.");
  } else {
    console.log(`${findings.length} control(s) a point or more out:`);
    for (const one of findings) {
      console.log(`  ${one}`);
    }
  }

  console.log("\nA point is about a device pixel and a half here, so under one is noise in the\n"
    + "edge detector rather than a difference anybody can see. Two or more is a real one.");
} finally {
  await api.userforms("close", CAPTION).catch(() => {});
  await waitFor("the form to unload and give its designer back", async () =>
    api.designer(FORM, project.projectId).then(() => true, () => false), { budgetMs: 20000 })
    .catch(() => {});
}
