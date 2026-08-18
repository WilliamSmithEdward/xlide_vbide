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

/*
 * NAME CONTROLS TO SEE THE PIXELS THE TABLE IS READING:
 *
 *   node tools\harness\designer-parity.mjs NameLabel PickGround
 *
 * The table reduces a column of the capture to one number, and one number cannot say WHAT the
 * edge detector found - a border, a bevel, the top of a letterform. That difference decides
 * whether a delta is a defect in the canvas or the instrument comparing two unlike things, and
 * getting it wrong is how four wrong picture fixes shipped in one day. So the profile is here
 * rather than in a throwaway probe: the next person asking "what is it actually seeing" should
 * find the answer in the tool that raised the question.
 */
const PROFILE = new Set(process.argv.slice(2).filter((one) => !one.startsWith("-")));

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
    /** The three channels at a pixel. Luminance cannot separate the logo's navy from the
     * black field behind it - `#002963` reads 35 and black reads 0 - and that separation is
     * the whole of the picture comparison below. */
    rgb(x, y) {
      const row = upward ? rows - 1 - y : y;
      const at = offset + row * stride + x * 4;
      return { r: bytes[at + 2], g: bytes[at + 1], b: bytes[at] };
    },
  };
}

/**
 * How a rectangle of a capture is MADE UP, in three buckets, as fractions of its area.
 *
 * Two surfaces photograph the same button at different scales - the runtime at the monitor's,
 * the canvas at the page's - so nothing can be compared pixel to pixel. What CAN be compared is
 * the mix, and for this defect the mix is the whole question: a black field drawn solid and the
 * same field keyed out differ by tens of percent of the button's area, which no amount of
 * rescaling hides.
 *
 *   black   the artwork's own background, opaque and unkeyed
 *   light   button face showing through, or around, the picture
 *   ink     everything else: the logo's colours, the caption's glyphs, the bevel's mid greys
 */
function pictureMix(image, x0, y0, w, h) {
  let black = 0;
  let light = 0;
  let counted = 0;

  for (let y = Math.max(0, y0); y < Math.min(y0 + h, image.height); y++) {
    for (let x = Math.max(0, x0); x < Math.min(x0 + w, image.width); x++) {
      const { r, g, b } = image.rgb(x, y);
      counted++;
      if (Math.max(r, g, b) < 24) {
        black++;
      } else if (0.299 * r + 0.587 * g + 0.114 * b > 200) {
        light++;
      }
    }
  }

  if (counted === 0) {
    return null;
  }

  const share = (n) => +(100 * n / counted).toFixed(1);
  return { black: share(black), light: share(light), ink: share(counted - black - light), counted };
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

  /*
   * THE PICTURE, NOT THE GEOMETRY - photographed BEFORE the launch, because the canvas does not
   * change when the form stands and a crop wants the element on screen.
   *
   * A caption picture and a surface picture are drawn two different ways by MSForms out of the
   * same pixels: on a button the runtime keys the artwork's background out and the button face
   * shows through, on an Image control it draws that same background solid. The canvas drew both
   * opaque, and FOUR fixes went at the wrong half of it - clipping, letterboxing, stretch in
   * place, absolute inset - each one read off a screenshot and inferred. Nothing here could catch
   * one, which is why four of them shipped. This is the row that can.
   */
  const PICTURED = "OkButton";
  const pictureAt = `.designer-view[data-module=${JSON.stringify(FORM)}] `
    + `.dc[data-control=${JSON.stringify(PICTURED)}]`;
  await api.ask(`(() => {
    const el = document.querySelector(${JSON.stringify(pictureAt)});
    if (el) { el.scrollIntoView({ block: "center", inline: "center" }); }
    return el ? "shown" : "missing";
  })()`);
  const canvasShot = await api.capture(undefined, undefined, { selector: pictureAt, pad: 0 })
    .then(decodeBmp, (why) => {
      console.log(`the canvas would not photograph ${PICTURED}: ${why.message}\n`);
      return null;
    });

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
  const profiles = [];
  for (const control of form.controls) {
    if (control.top === null || control.left === null || control.width < 6) {
      continue;
    }

    const drawn = painted[control.name];
    if (!drawn) {
      continue;
    }

    /*
     * A LABEL HAS NO EDGE TO COMPARE, and this table said it had one four points out.
     *
     * Profiled 2026-08-18 (`designer-parity.mjs NameLabel`): the runtime's column through
     * NameLabel is the form's ground - luminance 240 - from 11.25pt all the way to 17.27, and the
     * first thing that is not ground is at 18.03, running to 21.79 and back to ground. That is
     * the anti-aliased text of "Customer". There is no border and no fill: a Label's background
     * is the form's own, so the only ink in the column is the caption, four points below the
     * control's top because that is where the cap height of 8.25pt Tahoma starts in a 16pt box.
     *
     * The canvas's landmark for a Label is its RECTANGLE, so the table was subtracting a
     * rectangle from some lettering and reporting the leading as a defect. Both Labels sat in the
     * findings list for weeks looking like the worst rows in it.
     *
     * So a Label is excused the way a ScrollBar and a SpinButton already are - said out loud
     * rather than dropped, because a table that quietly skips a control is a table nobody can
     * check. What it would take to measure a Label honestly is a like-for-like comparison of INK
     * against INK, cropping the same rectangle off both surfaces; that is a bigger instrument
     * than this one and it is not pretended at here.
     */
    if (drawn.kind === "Label") {
      console.log(control.name.padEnd(16) + String(control.type ?? drawn.kind).padEnd(14)
        + String(control.parent?.toLowerCase() === FORM.toLowerCase() ? "form" : control.parent).padEnd(10)
        + String(drawn.top).padStart(7) + "  no edge".padStart(10)
        + "   caption ink only".padStart(8));
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

    if (PROFILE.has(control.name)) {
      profiles.push({ name: control.name, kind: drawn.kind, x, drawn, ceiling });
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

  /*
   * WHAT THE COLUMN ACTUALLY HOLDS, for a control named on the command line.
   *
   * The number in the table is the first row in this column that is not the form's ground, and it
   * cannot say what that row IS. A border, a bevel and the top of a letterform all read as "not
   * ground", and which one it found decides whether a delta is the canvas painting in the wrong
   * place or the table comparing a rectangle against some ink inside it.
   */
  for (const one of profiles) {
    const from = Math.round(originY + (one.drawn.boxTop - 3) * scale);
    const to = Math.round(originY + (one.drawn.boxTop + 14) * scale);
    console.log(`\n${one.name} (${one.kind}), the runtime's column at x=${one.x}`);
    console.log(`the canvas paints its landmark at ${one.drawn.top}pt; its BOX starts at ${one.drawn.boxTop}pt`);
    console.log("   pt    lum  ");
    for (const row of column(image, one.x, Math.max(0, from), to)) {
      const at = asPoints(row.y);
      const bar = "#".repeat(Math.max(0, Math.round((255 - row.lum) / 12)));
      console.log(`${at.toFixed(2).padStart(7)}${Math.round(row.lum).toString().padStart(7)}  ${bar}`
        + (Math.abs(at - one.drawn.boxTop) < 0.4 ? "   <- the canvas's box top" : "")
        + (Math.abs(at - one.drawn.top) < 0.4 && one.drawn.top !== one.drawn.boxTop
          ? "   <- the canvas's landmark" : ""));
    }
  }

  // ---- and the same button's PICTURE, off both surfaces ----
  const pictured = painted[PICTURED];
  const runtimeMix = pictured === undefined ? null : pictureMix(
    image,
    Math.round(originX + pictured.left * scale),
    Math.round(originY + pictured.boxTop * scale),
    Math.round(pictured.width * scale),
    Math.round((pictured.bottom - pictured.boxTop) * scale));
  const canvasMix = canvasShot === null
    ? null
    : pictureMix(canvasShot, 0, 0, canvasShot.width, canvasShot.height);

  console.log(`\n${PICTURED}'s picture, as each surface draws it\n` + "-".repeat(70));
  if (runtimeMix === null || canvasMix === null) {
    console.log("one of the two surfaces did not photograph, so there is nothing to compare.");
  } else {
    console.log("surface     black   light     ink   pixels");
    console.log(`runtime  ${String(runtimeMix.black).padStart(7)}`
      + `${String(runtimeMix.light).padStart(8)}${String(runtimeMix.ink).padStart(8)}`
      + `${String(runtimeMix.counted).padStart(9)}`);
    console.log(`canvas   ${String(canvasMix.black).padStart(7)}`
      + `${String(canvasMix.light).padStart(8)}${String(canvasMix.ink).padStart(8)}`
      + `${String(canvasMix.counted).padStart(9)}`);

    /*
     * WHICH BUCKET MOVED SAYS WHICH DEFECT IT IS, and that is the whole value of three buckets
     * rather than one number.
     *
     * BLACK is the artwork's own field. A canvas painting it where the runtime keys it out - the
     * defect four fixes chased - puts the canvas TENS of points above the runtime here, and no
     * amount of resampling does that.
     *
     * LIGHT and INK move together and the other way: the runtime downsamples 256x256 into 96x32
     * nearest-neighbour, which keeps every pixel either button face or logo, while the browser
     * interpolates and blends the two into mid greys. So a canvas that is short on LIGHT and long
     * on INK by the same amount is the same picture through a smoother filter, not a different
     * picture.
     */
    const gap = (bucket) => +(canvasMix[bucket] - runtimeMix[bucket]).toFixed(1);
    const signed = (n) => `${n > 0 ? "+" : ""}${n}`;
    console.log(`\nblack ${signed(gap("black"))}   light ${signed(gap("light"))}`
      + `   ink ${signed(gap("ink"))}  (points of the button's area, canvas less runtime)`);
    console.log(gap("black") >= 10
      ? "  THE BACKGROUND ITSELF: the canvas is painting the artwork's field where the runtime\n"
        + "  keys it out. Look at the two crops before touching fit, size or inset."
      : Math.abs(gap("light")) >= 10 || Math.abs(gap("ink")) >= 10
        ? "  RESAMPLING, not composition: light lost to ink is the browser interpolating where\n"
          + "  the runtime takes nearest-neighbour. Same picture, smoother filter."
        : "  The two surfaces draw this button the same way, to within the filter.");
  }
} finally {
  await api.userforms("close", CAPTION).catch(() => {});
  await waitFor("the form to unload and give its designer back", async () =>
    api.designer(FORM, project.projectId).then(() => true, () => false), { budgetMs: 20000 })
    .catch(() => {});
}
