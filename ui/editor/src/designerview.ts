/*
 * The form designer tab's body: the markup document and the visual form, side by side.
 *
 * One view per form, owned by the page and REPARENTED between editor groups as its tab moves,
 * so the markup editor's scroll, caret and undo ride along. Both halves are drawn from ONE
 * answer - the host walks the form once and sends the markup text with the spec it printed -
 * so the document and the picture cannot describe two different moments of the form.
 *
 * The visual is the honest canvas (docs/userform-designer.md, M3's opening stance): real
 * bounds, real captions, type-shaped placeholders, and nothing guessed. Bounds are points,
 * MSForMS's own unit, scaled 4/3 to CSS pixels; children sit INSIDE their container's element
 * the way they sit inside its client area, so containment is the browser's composition rather
 * than arithmetic here.
 *
 * Both halves EDIT, and they edit the same thing: a canvas gesture - a drag, a nudge - rewrites
 * the control's line in the document, one undoable edit each, exactly as the hand would have
 * typed it. That is what makes the document the designer's transaction log (M5): the draft
 * preview shows the gesture at once, the dirty dot says it has not reached the form, Ctrl+S
 * applies it, and Ctrl+Z takes it back from either half. The FORM is never written behind the
 * document's back.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import type { FormMarkupApplied, FormMarkupControl, FormMarkupDraft, FormMarkupKind, FormMarkupLintFinding, FormMarkupPayload } from "./bridge.js";
import { showContextMenu, type ContextMenuItem } from "./contextmenu.js";
import type { DocumentId } from "./documents.js";
import { installEdgeScroll, type EdgeScroll } from "./edgescroll.js";
import {
  FORM_MARKUP_LANGUAGE, TOOLBOX, completionsAt, headerHintAt, hoverAt,
  markupVocabulary, registerMarkupLanguage, setMarkupVocabulary,
} from "./formmarkuplang.js";
import { currentSettings, onSettingsApplied } from "./settings.js";
import {
  dressWithPicture, paintPictureSurface, pictureLayer,
  SURFACE_PICTURE_TYPES, type PictureLayer,
} from "./formpicture.js";
import { showTabOrder } from "./taborderdialog.js";

/** Points to CSS pixels at 96dpi: the designer's own unit, made visible at 100%. */
const PT = 4 / 3;

/*
 * A FRAME'S CAPTION BAND: TWO measurements, because it answers two questions.
 *
 * MSForms draws a group box the way Windows does: the caption occupies a band at the top of
 * the control's own rectangle, the rule runs through the MIDDLE of that band, and the client
 * area begins BELOW it. So a frame at top 112 shows its line at about 116 and its first child
 * lower still - which is why a button placed level with the frame looks level in a designer that
 * draws the line at 112 and is not (the owner's side-by-side, 2026-08-16).
 *
 * ONE CONSTANT WAS ANSWERING BOTH, and that is what left the children out (2026-08-18). The rule
 * was drawn at half the band and the client at the whole of it, which ties the two together at a
 * ratio the runtime does not keep: measured against the running form, the rule sits 4.69pt below
 * the control's top and the client begins 6.05pt below it - a gap of 1.36pt, not a factor of two.
 * With one number they could not both be right, and the client was 2.9pt low, so every control
 * inside a Frame sat almost three points below where it runs.
 *
 * HOW THEY WERE MEASURED, since the next person will want to redo it. designer-parity.mjs names
 * a control and prints the runtime's own column of pixels through it. An OptionButton was added
 * to the FORM to get the kind's own glyph offset with no container in the way - `ParityProbe` at
 * top 120 put its first ink at 124.22, so an OptionButton's glyph starts 4.22pt below its top.
 * PickGround, the same kind inside the Frame, put its ink at 136.27, so its control top is
 * 132.05; its Top is 14 within the frame, so the frame's client begins at 118.05 on a frame whose
 * own top is 112. The rule came off the same table. Neither number is reasoned about.
 */
const FRAME_RULE_TOP = 6.25;
const FRAME_CLIENT_TOP = 8;

/** How near an edge has to be, in POINTS, before a gesture lines up with it. Three is about
 * four device pixels at 100%: close enough that a hand aiming at an edge finds it, far enough
 * that a hand aiming between two of them is not dragged to either. */
const ALIGN_REACH = 3;

/** A line the canvas draws while a gesture is lining up with something: which axis it runs
 * along, and where it sits in POINTS from the container's own origin. */
interface Guide {
  axis: "x" | "y";
  at: number;
}

/** The tab strip's own height, as the stylesheet draws it. The rule below the tabs cannot sit
 * higher than this or the tabs hang through the line, and the runtime's band measures within
 * a point of it - so it is also the truer of the two numbers available. */
const PAGE_STRIP_HEIGHT = 20;

/*
 * AND HOW FAR INSIDE THE BODY A PAGE'S CHILDREN BEGIN, which is not where the body's edge is.
 *
 * The MultiPage had the Frame's defect in its own dialect (2026-08-18): the body's visible top
 * and the origin its children are placed from were one number, and the runtime keeps them about
 * two points apart. Measured the same way - a CheckBox on the FORM to get the kind's glyph offset
 * with no container in the way (Taxable, top 40, ink at 42.88, so 2.88pt), then the same kind on
 * a Page (Agree, ink at 216.1, so its top is 213.22; its Top is 8 within the page, so the page's
 * client begins at 205.22 on a MultiPage whose own top is 188). That is 17.22pt, against a body
 * edge measured at 15.3pt. The difference is the body's own inset and it is what this carries.
 */
const PAGE_CLIENT_INSET = 3;

/** How far a press must travel before it is a DRAG rather than a click, in CSS pixels. Below
 * it a press is a selection and nothing moves, which is what keeps a click a click. */
const DRAG_THRESHOLD = 3;

/** How long the lint waits after a KEYSTROKE: long enough that a word is not asked about
 * mid-typing, and the canvas's draft follows the same beat. A GESTURE does not wait it out -
 * see `lintNow` - because a drop or a delete is finished the moment it happens. */
const TYPING_DEBOUNCE = 350;

/** The arrow keys' vocabulary: one point a press, the finest move the document can spell. */
const NUDGE: Readonly<Record<string, { dx: number; dy: number }>> = {
  ArrowLeft: { dx: -1, dy: 0 },
  ArrowRight: { dx: 1, dy: 0 },
  ArrowUp: { dx: 0, dy: -1 },
  ArrowDown: { dx: 0, dy: 1 },
};

/** A name as a literal inside a pattern: control names are the developer's, and a dot in one
 * must not quietly become "any character". */
function escapeForRegExp(name: string): string {
  return name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** A declaring line in the document, split at its geometry: everything before the clauses, the
 * position and the size each clause spells (null where the line carries none), and whether this
 * is the FORM's line, which takes a size and never a position. The shape a move or a resize
 * rewrites. */
interface HeaderLine {
  line: number;
  /** The opening tag with its GEOMETRY attributes taken out and its closing removed, so a
   * gesture rebuilds the line as `head` + the geometry it wants + `close`. Geometry is removed
   * rather than left in place because the attributes can sit anywhere in the tag, and rewriting
   * them where they are means four ranges that all move under each other. */
  head: string;
  /** How the tag ended: `/>` for a childless element, `>` for one that opens a block. */
  close: string;
  at: { left: number; top: number } | null;
  size: { width: number; height: number } | null;
  form: boolean;
}

/** A box in POINTS, the document's unit: what a gesture is proposing. */
interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** A press that may become a gesture: what it grabbed, which handle if any, where it started,
 * the box it started from (in CSS pixels), and whether it has travelled far enough to count. */
interface CanvasDrag {
  name: string;
  element: HTMLElement;
  edge: string | null;
  pointerId: number;
  startX: number;
  startY: number;
  origin: Box;
  moved: boolean;
  /** The container the element sits in right now, by name - "" for the form's own ground. A move
   * that crosses into another container carries the element there as it goes, so this changes
   * mid-gesture and `where` is what it started as. */
  parent: string;
  where: string;
  /** Where inside the control the press landed, in CSS pixels. A reparent puts the control back
   * under the pointer by this offset, which is what a hand carrying something expects - and the
   * only way to place it that does not depend on how far the last pointer event jumped. */
  grabX: number;
  grabY: number;
}

/** The smallest a gesture may make something, in points. A control can be tiny - the renderer
 * floors its own boxes at four - but a form small enough to lose its title bar is a mistake
 * rather than a design. */
const MIN_CONTROL = 4;
const MIN_FORM = 24;

/*
 * The palette's kinds and their drop sizes moved to formmarkuplang.ts, because the language
 * service scaffolds a header from the same numbers: what a new Label looks like is one fact,
 * whether it arrives by a drag from the palette or by accepting a completion.
 */

/**
 * The palette's icons, 16x16, drawn for the palette rather than borrowed from the canvas.
 *
 * The first landing put a real canvas control in each button at 15px, on the theory that the
 * palette and the form should not hold two opinions about what a kind looks like. At that size
 * every kind is the same grey rectangle - what tells a ComboBox from a ListBox on the canvas is
 * its children, and children that small are noise (the owner, one look: "glyphs don't look
 * great"). So these are icons: the shape a person recognises, in `currentColor`, at the one size
 * they are ever drawn.
 */
const TOOL_ICON: Readonly<Record<string, string>> = {
  Label: '<path d="M2.5 5h11M2.5 8h8M2.5 11h5"/>',
  TextBox: '<rect x="1.5" y="4.5" width="13" height="7" rx="1"/><path d="M4 6.5v3M3 6.5h2M3 9.5h2"/>',
  ComboBox: '<rect x="1.5" y="4.5" width="13" height="7" rx="1"/><path d="M10.5 4.5v7"/><path d="M11.5 7.5l1 1.5 1-1.5z" fill="currentColor"/>',
  ListBox: '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><path d="M4 5.5h8M4 8h8M4 10.5h5"/>',
  CheckBox: '<rect x="1.5" y="3.5" width="9" height="9" rx="1"/><path d="M3.5 8l2.5 2.5 4-5"/>',
  OptionButton: '<circle cx="6" cy="8" r="4.5"/><circle cx="6" cy="8" r="1.8" fill="currentColor" stroke="none"/>',
  ToggleButton: '<rect x="1.5" y="4.5" width="13" height="7" rx="1.5" fill="currentColor" fill-opacity="0.28"/><path d="M4.5 6.5v3"/>',
  Frame: '<path d="M4.5 3.5H14v10H2v-10h1"/><path d="M3.5 3.5h1"/>',
  CommandButton: '<rect x="1.5" y="4" width="13" height="8" rx="1.5"/><path d="M5 8h6"/>',
  TabStrip: '<path d="M1.5 5.5h5v-2h4v2h5v8h-14z"/>',
  MultiPage: '<path d="M1.5 5.5h5v-2h4v2h5v8h-14z"/><path d="M2.5 5.5h4v-1h3v1" fill="currentColor" fill-opacity="0.35" stroke="none"/><path d="M4 9h8M4 11h5"/>',
  ScrollBar: '<rect x="5.5" y="1.5" width="5" height="13" rx="1"/><path d="M6.8 4.4L8 3l1.2 1.4zM6.8 11.6L8 13l1.2-1.4z" fill="currentColor"/>',
  SpinButton: '<rect x="5.5" y="3.5" width="5" height="9" rx="1"/><path d="M6.8 7L8 5.6 9.2 7zM6.8 9L8 10.4 9.2 9z" fill="currentColor"/>',
  Image: '<rect x="1.5" y="3.5" width="13" height="9" rx="1"/><circle cx="5" cy="6.5" r="1.2" fill="currentColor" stroke="none"/><path d="M2 11.5l3.5-3 2.5 2 2-1.5 4 2.5"/>',
};

/** What a palette button says when the kind's own name is too long for one: short enough to
 * read at a glance, and the full name is on the tooltip either way. */
const SHORT_KIND: Readonly<Record<string, string>> = {
  OptionButton: "Option",
  ToggleButton: "Toggle",
  CommandButton: "Button",
  ScrollBar: "Scroll",
  SpinButton: "Spin",
};

/** The cursor each handle's pull deserves, worn by the whole canvas while that pull runs. */
const EDGE_CURSOR: Readonly<Record<string, string>> = {
  nw: "nwse-resize",
  n: "ns-resize",
  ne: "nesw-resize",
  e: "ew-resize",
  se: "nwse-resize",
  s: "ns-resize",
  sw: "nesw-resize",
  w: "ew-resize",
};

export interface DesignerViewDeps {
  /** Ask the host for the form's projection; the answer arrives through watch. */
  request(): void;
  /** Subscribe to the form's projection answers; returns the unwatch. */
  watch(listener: (payload: FormMarkupPayload) => void): () => void;
  /** Apply the document to the form; the outcome arrives through watchApplied, and a fresh
   * projection follows it through watch. */
  apply(markup: string): void;
  /** Subscribe to the form's apply outcomes; returns the unwatch. */
  watchApplied(listener: (outcome: FormMarkupApplied) => void): () => void;
  /** The document's unapplied-edit state changed; the tab wears the dot. */
  dirtyChanged(dirty: boolean): void;
  /** Lint the document as it stands; findings return through watchLint. */
  lint(markup: string): void;
  /** Subscribe to the form's lint answers - squiggles, and the DRAFT the document parsed
   * to, which the canvas previews while the document is dirty; returns the unwatch. */
  watchLint(listener: (findings: FormMarkupLintFinding[], draft: FormMarkupDraft | null) => void): () => void;
  /** Bring to Front / Send to Back: the one canvas gesture that writes the MODEL rather than the
   * document, because MSForms' collection order is not z-order and the dialect cannot say it. */
  zorder(control: string, front: boolean): void;

  /** Writes one of a control's properties straight at the model, through the same host call the
   * Properties panel makes. The tab-order dialog's Move Up is a TabIndex write and nothing else. */
  setProperty(control: string, property: string, value: string): void;

  /** Asks the host for the markup language's vocabulary: the kinds and their properties, which
   * the document's completions and hovers answer from. Once per session is enough. */
  requestVocabulary(): void;
  /** Subscribe to the vocabulary; returns the unwatch. Not keyed by form - the language is the
   * same in every document, so whichever tab asked, every tab is answered. */
  watchVocabulary(listener: (kinds: FormMarkupKind[]) => void): () => void;
  /** Save the workbook, the host's own File Save - the second half of the designer's
   * Ctrl+S, after a successful apply. With `run`, the host launches the form after the
   * save, which is what F5 over a designer tab asks for. */
  saveWorkbook(run: boolean): void;
  /** Subscribe to the HOST's Ctrl+S and F5: neither accelerator reaches the page, so the host
   * asks the tab to apply-then-save through this - `run` is F5's - and returns the unwatch. */
  watchApplySave(listener: (run: boolean) => void): () => void;
  /** A canvas double-click: the host writes or shows the control's default event handler.
   * Null means the form itself. */
  eventStub(control: string | null): void;
  /** The selection changed: the Properties panel follows. Null is the form itself. */
  selection(control: string | null): void;
  /** Writes one setting through the host - the same call the settings dialog's controls make,
   * so the grid's switch here and its row there are two views of one fact. */
  changeSetting(key: string, value: unknown): void;

  /** Opens this form's code half at a line - the Sub an annotated handler is defined on. */
  openHandler(line: number): void;

  /** The form's own procedures, for showing which controls have handlers written against them.
   * Answers null when the host cannot say, which is a shrug rather than "there are none". */
  handlers(): Promise<{ name: string; kind: string; line: number }[] | null>;

  /** What MSForms' own AutoSize would make a control - the only thing that knows a control's
   * natural size, since a check box's glyph, a button's chrome and a picture at natural size are
   * none of them in the caption's ink. The host measures and puts the control back; null is a
   * control that has no natural size to give. */
  autoSize(control: string): Promise<{ width: number; height: number } | null>;

  /** Says something on the status line - the same place a refused host write says its piece.
   * For gestures that can decline: Ctrl+V with nothing copied, a Page pasted where no MultiPage
   * is. A gesture that silently does nothing is indistinguishable from one that is unwired. */
  notify(text: string): void;
}

export class DesignerView {
  readonly root: HTMLElement;
  readonly id: DocumentId;

  private readonly markupHost: HTMLElement;
  private readonly markupHalf: HTMLElement;
  private readonly splitter: HTMLElement;
  private readonly canvasScroll: HTMLElement;

  /** Stacked shows the form ABOVE the document, the design-surface convention; side-by-side
   * is the wide-monitor alternative a double-click on the divider reaches. */
  private orientation: "stacked" | "beside" = "stacked";

  /** One half at a time can leave: the divider's chevrons hide the document or the form and
   * bring them back, and the divider itself stays as the way home. */
  private collapsed: "markup" | "canvas" | null = null;
  private readonly hideMarkupButton: HTMLElement;
  private readonly hideCanvasButton: HTMLElement;

  private setCollapsed(which: "markup" | "canvas" | null): void {
    this.collapsed = which;
    this.root.classList.toggle("hide-markup", which === "markup");
    this.root.classList.toggle("hide-canvas", which === "canvas");
    this.refreshCollapseTitles();
    this.layout();
  }

  private refreshCollapseTitles(): void {
    this.hideMarkupButton.title = this.collapsed === "markup" ? "Show the document" : "Hide the document";
    this.hideMarkupButton.setAttribute("aria-label", this.hideMarkupButton.title);
    this.hideCanvasButton.title = this.collapsed === "canvas" ? "Show the form" : "Hide the form";
    this.hideCanvasButton.setAttribute("aria-label", this.hideCanvasButton.title);
  }
  private readonly notice: HTMLElement;
  private readonly errorStrip: HTMLElement;
  private readonly editor: monaco.editor.IStandaloneCodeEditor;
  private readonly model: monaco.editor.ITextModel;
  private readonly unwatch: () => void;
  private readonly unwatchApplied: () => void;
  private readonly unwatchLint: () => void;
  private readonly unwatchApplySave: () => void;
  private readonly unwatchVocabulary: () => void;
  private lintTimer: ReturnType<typeof setTimeout> | undefined;
  private readonly deps: DesignerViewDeps;

  /** The last APPLIED projection - the form's own truth, and the wardrobe a draft preview
   * borrows its display extras from. */
  private lastPayload: FormMarkupPayload | null = null;

  /** The projection the canvas is DRAWING - the draft while one stands, the applied one
   * otherwise - so a gesture that only changes which page is open can redraw the picture
   * without asking the host for it again. */
  private shownPayload: FormMarkupPayload | null = null;

  /**
   * Which page each MultiPage is showing, by the page's own NAME, and which tab each TabStrip
   * has selected, by index. The two are keyed differently because the two things are: a
   * MultiPage's pages are controls with names, and a TabStrip's tabs are not controls at all.
   *
   * VIEW state, and that is the one place this canvas departs from the native designer, which
   * writes the container's `Value` and dirties the form. Reaching page 2 must not rewrite the
   * developer's form: the document is the transaction log, so a switch that landed in it would
   * make LOOKING a change and Ctrl+S would carry it to the form. Navigation is not manipulation
   * here, the way scrolling and selection are not. A developer who wants the form to OPEN on a
   * page says so where every other unprinted property is said - the Properties panel.
   */
  private readonly shownPage = new Map<string, string>();
  private readonly shownTab = new Map<string, number>();

  /**
   * How large the canvas draws the form, 1 being MSForms' own points at 96dpi.
   *
   * VIEW state rather than a setting, like which page is open: it is about looking, not about
   * behaviour, and a big form wants Fit where the one beside it wants 100%. The picture is a CSS
   * transform on the form itself, so everything INSIDE it stays in the form's own coordinates -
   * only the places where screen pixels cross into points know about this at all, and they go
   * through `toPoints` and `toPixels`.
   */
  private zoom = 1;

  /** The FORM's own picture as a background layer, held rather than painted straight on: the
   * grid paints on the same element, and showGrid composes the two so that a form wearing a
   * picture still shows its grid and a grid toggle does not wipe the picture. */
  private formGround: PictureLayer | null = null;

  /** Screen pixels to the form's own points, at whatever zoom is showing. */
  private toPoints(pixels: number): number {
    return pixels / (PT * this.zoom);
  }

  /** The form's own points to screen pixels, the same conversion the other way. */
  private toPixels(points: number): number {
    return points * PT * this.zoom;
  }

  /** Whether the canvas is currently showing the DOCUMENT's draft rather than the form. */
  private draftShown = false;

  /** The selected control's name, "" for the FORM itself, null for no selection. With more than
   * one selected this is the ANCHOR - the native designer's primary control: the one the handles
   * dress, the one the Properties panel follows, and the one an alignment lines the rest up
   * with. */
  private selectedName: string | null = null;

  /** The rest of the selection, by name, never holding the anchor. Empty for the ordinary case
   * of one selected thing, which is what every gesture still reads as `selectedName`. */
  private readonly extras = new Set<string>();

  /** The rubber band in flight: a press on a container's own ground that has started to travel. */
  private band: {
    pointerId: number;
    host: HTMLElement;
    parent: string;
    startX: number;
    startY: number;
    element: HTMLElement | null;
  } | null = null;

  /** The press in flight, once it has grabbed something movable. */
  private drag: CanvasDrag | null = null;

  /** The palette, its edge arrows, and the kind being carried out of it right now. */
  private readonly toolbox: HTMLElement;
  private readonly toolboxEdges: EdgeScroll;
  private carrying: { kind: string; width: number; height: number; ghost: HTMLElement } | null = null;

  /** The grid's switch, at the end of the palette, and the unsubscribe for the setting it
   * follows: the button shows what the setting says rather than what it was last clicked to. */
  private readonly snapToggle: HTMLButtonElement;
  private readonly alignToggle: HTMLButtonElement;

  /** The zoom control at the end of the palette row: it wears the percentage it is showing. */
  private readonly zoomButton: HTMLButtonElement;
  private readonly unwatchSettings: () => void;

  /** The last text known to BE the form - what dirty is measured against. Null before the
   * first projection arrives, and while an apply's fresh projection is on its way. */
  private canonical: string | null = null;
  private dirty = false;
  /** An apply was accepted; the next projection is its canonical text and is adopted even
   * though the document was just marked clean. */
  private awaitingAdopt = false;

  constructor(id: DocumentId, deps: DesignerViewDeps) {
    this.id = id;

    this.root = document.createElement("div");
    this.root.className = "designer-view";
    this.root.dataset.module = id.module;
    this.root.dataset.project = id.project ?? "";

    this.markupHost = document.createElement("div");
    this.markupHost.className = "designer-markup";

    const splitter = document.createElement("div");
    splitter.className = "designer-splitter";
    splitter.setAttribute("role", "separator");
    splitter.setAttribute("aria-label", "Resize markup and form; double-click to switch the split");
    splitter.title = "Drag to resize; double-click to switch between stacked and side-by-side";
    splitter.tabIndex = 0;

    const grip = document.createElement("div");
    grip.className = "designer-splitter-grip";
    splitter.appendChild(grip);

    // The halves hide one at a time from the divider itself: a full-width canvas while
    // arranging, a full-width document while writing, one click each way.
    const hideMarkup = document.createElement("button");
    hideMarkup.type = "button";
    // A REAL chevron rather than a border-trick pseudo-element: the owner caught one of
    // the triangles rendering as an empty button in a state the headless frame does not
    // reproduce (2026-08-15), and an inline path is immune to the whole class - border
    // rounding, margin interplay, rotation - at any zoom. One glyph pointing UP; the
    // stylesheet rotates it per orientation and collapse state.
    const chevron = '<svg viewBox="0 0 10 10" aria-hidden="true"><path d="M2 6.5 L5 3.5 L8 6.5" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>';

    hideMarkup.className = "designer-collapse designer-collapse-markup";
    hideMarkup.innerHTML = chevron;
    hideMarkup.addEventListener("click", (event) => {
      event.stopPropagation();
      this.setCollapsed(this.collapsed === "markup" ? null : "markup");
    });
    hideMarkup.addEventListener("dblclick", (event) => event.stopPropagation());
    hideMarkup.addEventListener("pointerdown", (event) => event.stopPropagation());

    const hideCanvas = document.createElement("button");
    hideCanvas.type = "button";
    hideCanvas.className = "designer-collapse designer-collapse-canvas";
    hideCanvas.innerHTML = chevron;
    hideCanvas.addEventListener("click", (event) => {
      event.stopPropagation();
      this.setCollapsed(this.collapsed === "canvas" ? null : "canvas");
    });
    hideCanvas.addEventListener("dblclick", (event) => event.stopPropagation());
    hideCanvas.addEventListener("pointerdown", (event) => event.stopPropagation());

    splitter.append(hideMarkup, hideCanvas);
    this.hideMarkupButton = hideMarkup;
    this.hideCanvasButton = hideCanvas;

    const canvasHalf = document.createElement("div");
    canvasHalf.className = "designer-canvas-half";

    // The xlide toolbox: our own palette, docked in the tab rather than floating over it, in
    // the strip the draft banner used to occupy. The native Toolbox stays suppressed - this is
    // the thing it was suppressed FOR (docs/userform-designer.md, M5).
    //
    // ONE ROW that scrolls, with the same edge arrows the command strip and the tab strip wear
    // (the owner's call, 2026-08-15): wrapping to a second row eats the canvas on a narrow tab,
    // and a strip that scrolls the way the others do is one behaviour to learn, not three.
    const toolboxRow = document.createElement("div");
    toolboxRow.className = "designer-toolbox";
    const toolbox = document.createElement("div");
    toolbox.className = "designer-toolbox-strip";
    toolbox.setAttribute("role", "toolbar");
    toolbox.setAttribute("aria-label", "Controls to drag onto the form");
    toolboxRow.appendChild(toolbox);
    for (const tool of TOOLBOX) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "designer-tool";
      button.dataset.kind = tool.kind;
      button.title = `Drag a ${tool.kind} onto the form`;
      const glyph = document.createElement("span");
      glyph.className = "designer-tool-glyph";
      glyph.innerHTML = '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" '
        + 'fill="none" stroke="currentColor" stroke-width="1.1" stroke-linecap="round" '
        + `stroke-linejoin="round">${TOOL_ICON[tool.kind] ?? ""}</svg>`;
      button.append(glyph, document.createTextNode(SHORT_KIND[tool.kind] ?? tool.kind));
      toolbox.appendChild(button);
    }

    toolbox.addEventListener("pointerdown", (event) => this.onToolboxPress(event));

    /*
     * THE GRID'S SWITCH, at the end of the palette rather than buried in the settings dialog.
     *
     * It belongs where the work is: snapping is something a developer turns off for one
     * awkward control and back on immediately, and a switch two dialogs away is a switch that
     * stays where it was. It writes the SETTING, so the dialog, the api and this button are
     * three views of one fact rather than three states to keep in step, and it survives the
     * session because the setting does.
     */
    const switchFor = (mode: "grid" | "objects", glyph: string): HTMLButtonElement => {
      const control = document.createElement("button");
      control.type = "button";
      control.className = "designer-snap";
      control.dataset.mode = mode;
      control.innerHTML = '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" '
        + `fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round">${glyph}</svg>`;
      // Each one is a TOGGLE of its own mode, and the two together are the either/or: pressing
      // the one that is already on turns snapping off, pressing the other switches. Two plain
      // buttons rather than one that cycles, because a cycling button hides its third state.
      control.addEventListener("click", () => {
        deps.changeSetting("designerSnap", currentSettings().designerSnap === mode ? "off" : mode);
      });
      toolboxRow.appendChild(control);
      return control;
    };

    // Nine dots for the grid; a shape between two guide lines for the neighbours.
    this.snapToggle = switchFor("grid",
      '<g fill="currentColor" stroke="none">'
      + '<circle cx="4" cy="4" r="1"/><circle cx="8" cy="4" r="1"/><circle cx="12" cy="4" r="1"/>'
      + '<circle cx="4" cy="8" r="1"/><circle cx="8" cy="8" r="1"/><circle cx="12" cy="8" r="1"/>'
      + '<circle cx="4" cy="12" r="1"/><circle cx="8" cy="12" r="1"/><circle cx="12" cy="12" r="1"/>'
      + '</g>');
    this.alignToggle = switchFor("objects",
      '<path d="M3 1.5v13"/><path d="M13 1.5v13"/><rect x="5.5" y="5" width="5" height="6" rx="1"/>');

    /*
     * TAB ORDER, on the strip beside the switches.
     *
     * It was reachable only from the canvas's own right-click menu, which is a menu a developer
     * has to know is there - and unlike Align and Distribute beside it, tab order is not a
     * gesture with a visible equivalent on the canvas, so nothing else hints that it exists
     * (the owner, 2026-08-18: "please add tab order button to strip"). The native editor keeps
     * it on a MENU (View > Tab Order), and this product has no menu bar, so the strip is where
     * the equivalent affordance belongs.
     *
     * It opens for whatever the selection sits in, exactly as the menu item does - one call, so
     * the two entry points cannot drift into two behaviours.
     */
    const tabOrder = document.createElement("button");
    tabOrder.type = "button";
    tabOrder.className = "designer-taborder";
    tabOrder.title = "The order the Tab key reaches the controls in";
    tabOrder.innerHTML = '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" '
      + 'fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" '
      + 'stroke-linejoin="round">'
      + '<path d="M2.5 4h6"/><path d="M2.5 8h6"/><path d="M2.5 12h6"/>'
      + '<path d="M11 3.5v9"/><path d="M8.75 10.25 11 12.5l2.25-2.25"/>'
      + "</svg>";
    tabOrder.addEventListener("click", () => { this.showTabOrder(); });
    toolboxRow.appendChild(tabOrder);

    /*
     * The ZOOM, at the end of the row beside the snap switches: a button wearing the percentage,
     * opening the same menu the rest of this product's menus use. A control, not a setting -
     * which page is open is view state for the same reason, and a big form wants Fit where the
     * one beside it wants 100%.
     */
    this.zoomButton = document.createElement("button");
    this.zoomButton.type = "button";
    this.zoomButton.className = "designer-zoom";
    this.zoomButton.title = "How large the form is drawn";
    this.zoomButton.textContent = "100%";
    this.zoomButton.addEventListener("click", (event) => {
      const box = this.zoomButton.getBoundingClientRect();
      showContextMenu(Math.round(box.left), Math.round(box.bottom + 2), [
        ...[0.5, 0.75, 1, 1.5, 2].map((factor) => ({
          label: `${Math.round(factor * 100)}%`,
          run: () => { this.setZoom(factor); },
        })),
        {},
        { label: "Fit", run: () => { this.setZoom("fit"); } },
      ]);
      event.preventDefault();
    });
    toolboxRow.appendChild(this.zoomButton);

    canvasHalf.appendChild(toolboxRow);
    this.toolbox = toolbox;
    // Inserted into the strip's own parent, which is why the row exists: the arrows sit beside
    // the palette rather than inside it, and appear only when there is something past the edge.
    this.toolboxEdges = installEdgeScroll(toolbox, "toolbar-edge");

    this.canvasScroll = document.createElement("div");
    this.canvasScroll.className = "designer-canvas-scroll";
    canvasHalf.appendChild(this.canvasScroll);

    // The wheel scrolls the canvas wherever it lands on it, unconditionally - and this is
    // also what gives the harness a wheel it can drive, because a synthesised WheelEvent
    // never triggers native scrolling (the owner, 2026-08-15: "need to be able to scroll
    // form designer window, currently cant").
    canvasHalf.addEventListener("wheel", (event) => {
      this.canvasScroll.scrollTop += event.deltaY;
      this.canvasScroll.scrollLeft += event.deltaX;
      event.preventDefault();
    }, { passive: false });

    // A press gives the canvas focus, so PageUp, arrows and Ctrl+S work from either half:
    // a keydown routes through the view only when something inside it holds focus.
    this.canvasScroll.tabIndex = -1;

    // M3's gesture, now carrying M5's: a press SELECTS - a control by name, the form by its
    // ground - the markup caret follows to the selected thing's line, and the same press
    // arms a drag. Selection on PRESS rather than release is the native designer's timing,
    // and it is what lets one gesture both pick and move. A double-click asks the host for
    // the default event handler, the native gesture; it also fires the presses, which is
    // native too.
    this.canvasScroll.addEventListener("pointerdown", (event) => this.onCanvasPress(event));
    this.canvasScroll.addEventListener("pointermove", (event) => this.onCanvasDragMove(event));
    this.canvasScroll.addEventListener("pointerup", (event) => this.onCanvasDrop(event));
    this.canvasScroll.addEventListener("pointercancel", () => this.cancelDrag());
    this.canvasScroll.addEventListener("keydown", (event) => this.onCanvasKey(event));
    this.canvasScroll.addEventListener("contextmenu", (event) => this.onCanvasMenu(event));
    this.canvasScroll.addEventListener("dblclick", (event) => {
      const control = (event.target as HTMLElement).closest<HTMLElement>(".dc");
      if (control?.dataset.control) {
        // A PAGE gets its own handler, and that is measured rather than assumed: a page raises
        // Click, the host writes `Page1_Click` for one, and the VBE's own object list carries
        // pages beside the controls. So the page's ground - double-clickable since the body took
        // the page's identity - asks for the page's handler, exactly as the native designer does.
        this.deps.eventStub(control.dataset.control);
      } else if ((event.target as HTMLElement).closest(".dc-form")) {
        this.deps.eventStub(null);
      }
    });

    // Ctrl+S anywhere in the view is the same save. Capture, so it runs ahead of the
    // markup editor's own binding and nothing double-fires.
    this.root.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && !event.altKey
        && (event.key === "s" || event.key === "S")) {
        event.preventDefault();
        event.stopPropagation();
        this.applyNow();
      }
    }, { capture: true });

    this.notice = document.createElement("div");
    this.notice.className = "designer-notice";
    this.notice.hidden = true;
    canvasHalf.appendChild(this.notice);

    // Where a refusal lands: at the document that earned it, with the host's own wording -
    // the line number for a parse, "what landed first" for a stop partway.
    this.errorStrip = document.createElement("div");
    this.errorStrip.className = "designer-error";
    this.errorStrip.setAttribute("role", "alert");
    this.errorStrip.hidden = true;

    this.markupHalf = document.createElement("div");
    this.markupHalf.className = "designer-markup-half";
    this.markupHalf.append(this.markupHost, this.errorStrip);

    this.root.append(this.markupHalf, splitter, canvasHalf);
    this.splitter = splitter;

    registerMarkupLanguage();
    this.model = monaco.editor.createModel("", FORM_MARKUP_LANGUAGE,
      monaco.Uri.parse(`xlide-form:/${encodeURIComponent((id.project ?? "").toLowerCase())}/${encodeURIComponent(id.module)}`));
    this.editor = monaco.editor.create(this.markupHost, {
      model: this.model,
      // The view is BUILT detached - the workspace mounts it a beat later - so the editor's
      // creation-time measure is zero. The observer picks up the real box on mount, on group
      // moves, and on the splitter, the same way the group editors track their containers.
      automaticLayout: true,
      // The language's own unit: four spaces per level, the printer's and the parser's.
      tabSize: 4,
      insertSpaces: true,
      detectIndentation: false,
      minimap: { enabled: false },
      lineNumbers: "on",
      folding: false,
      wordWrap: "off",
      scrollBeyondLastLine: false,
      renderLineHighlight: "none",
      occurrencesHighlight: "off",
      fixedOverflowWidgets: true,
    });

    // Ctrl+S applies THIS document, while focus is in it; the code editors' save flows are
    // untouched because the command binds to this editor instance alone.
    this.editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS, () => this.applyNow());

    /*
     * THE CARET PICKS THE CONTROL, which is the canvas-to-markup link run backwards.
     *
     * Only when the caret MOVED by a hand - a click, an arrow, a search landing - because the
     * other direction sets the caret too, and a selection that re-selects on its own echo is a
     * loop. Monaco names the reason, so the loop is cut by asking rather than by a flag that
     * has to be cleared on every path out.
     */
    this.editor.onMouseDown((event) => this.onMarkupClick(event));
    this.editor.onDidChangeCursorPosition((event) => {
      const byHand = event.reason === monaco.editor.CursorChangeReason.Explicit
        || event.reason === monaco.editor.CursorChangeReason.NotSet;
      if (byHand && !this.followingSelection) {
        this.selectFromMarkup(event.position.lineNumber);
      }
    });

    this.model.onDidChangeContent(() => {
      const nowDirty = this.canonical !== null && this.model.getValue() !== this.canonical;
      if (nowDirty !== this.dirty) {
        this.dirty = nowDirty;
        deps.dirtyChanged(nowDirty);
      }

      // The squiggles follow the typing, debounced past the keystroke rate: the lint is a
      // host round trip, and a request per keystroke asks the same question mid-word.
      clearTimeout(this.lintTimer);
      this.lintTimer = setTimeout(() => deps.lint(this.model.getValue()), TYPING_DEBOUNCE);
    });

    this.installSplitter(splitter);

    this.deps = deps;
    this.unwatch = deps.watch((payload) => this.update(payload));
    this.unwatchApplied = deps.watchApplied((outcome) => this.onApplied(outcome));
    this.unwatchLint = deps.watchLint((findings, draft) => {
      this.showLint(findings);
      this.previewDraft(findings, draft);
    });
    // The REAL Ctrl+S: a host accelerator the page never sees as a key. The host's Save,
    // finding this tab active, asks for the apply-then-save here.
    this.unwatchApplySave = deps.watchApplySave((run) => this.applyNow(run));

    // The language's vocabulary: the host measures it, the page holds it for its whole life,
    // and no keystroke waits on it. Asked for only while the table is empty - the language is
    // the same in every document, so the second form to open inherits the first one's answer,
    // and an ask that went unanswered is retried by the next tab rather than never again.
    this.unwatchVocabulary = deps.watchVocabulary((kinds) => setMarkupVocabulary(kinds));
    if (markupVocabulary().length === 0) {
      deps.requestVocabulary();
    }

    // The grid follows the SETTING, from wherever it was changed: this tab's switch, the
    // settings dialog, the api, another designer tab. One fact, several views of it.
    this.unwatchSettings = onSettingsApplied(() => this.showGrid());
    this.showGrid();

    this.request = deps.request;

    // A tab that is ALREADY OPEN gets no projection just because the page reloaded under it,
    // which is precisely the state a developer is in after a deploy. The ticket above is what
    // makes this safe to fire alongside the one `update` starts.
    this.refreshHandlers();


    // LAST, because it lays the editor out and the editor must exist: the first version ran
    // this beside the DOM build and the constructor died reaching for an editor not yet
    // created - a designer tab that stood in the strip and mounted nothing.
    this.setOrientation("stacked");
    this.refreshCollapseTitles();
  }

  private readonly request: () => void;

  /** Set when an apply should be followed by the host's save - every Ctrl+S is, because
   * Ctrl+S means "save the workbook" everywhere else in the product and the designer must
   * not quietly mean less (the owner, 2026-08-15: "CTRL+S in designer isn't saving"). `run`
   * carries F5's extra half through the apply, so the launch waits for the same save. */
  private pendingSave: { run: boolean } | null = null;

  /** Ctrl+S: the document to the form, then the workbook to disk. A clean document has
   * nothing to apply and just saves. F5 comes through here too, with `run`, because the
   * form the developer is about to see must be the document they are looking at. */
  applyNow(run = false): void {
    this.errorStrip.hidden = true;
    if (!this.dirty) {
      this.deps.saveWorkbook(run);
      return;
    }

    this.pendingSave = { run };
    this.deps.apply(this.model.getValue());
  }

  /** For the debug surface: set the document and apply, answering the outcome - the same
   * path Ctrl+S takes, save included, which is the point of driving it from a suite. The
   * apply is unconditional here: the act's contract is an outcome, and the clean-document
   * shortcut above answers with a save instead of one. */
  applyDocument(markup: string): Promise<FormMarkupApplied> {
    this.model.setValue(markup);
    return new Promise((settle) => {
      this.pendingActOutcome = settle;
      this.errorStrip.hidden = true;
      this.pendingSave = { run: false };
      this.deps.apply(this.model.getValue());
    });
  }

  /** The document as it stands, for the debug surface's read side. */
  markupText(): string {
    return this.model.getValue();
  }

  /**
   * Puts the markup caret on a line, for the debug surface: the drive side of the selection
   * link that runs markup-to-canvas. It goes through the editor's own position, so the cursor
   * listener answers it exactly as it answers a click - which is the point of driving it here
   * rather than calling the selection directly.
   */
  markupCaret(line: number): string {
    const lines = this.model.getLineCount();
    if (!Number.isFinite(line) || line < 1 || line > lines) {
      return `line ${line} is outside the document's ${lines}`;
    }

    this.editor.focus();
    this.editor.setPosition({ lineNumber: line, column: 1 });
    this.editor.revealLineInCenterIfOutsideViewport(line);
    return `caret on line ${line}, selecting ${this.selectedName === "" ? "the form" : this.selectedName}`;
  }

  private pendingActOutcome: ((outcome: FormMarkupApplied) => void) | null = null;

  /** The squiggles: the host's tolerant parse, drawn on the document's own model. */
  private showLint(findings: FormMarkupLintFinding[]): void {
    monaco.editor.setModelMarkers(this.model, "xlide-form", findings.map((finding) => {
      const line = Math.min(Math.max(1, finding.line), this.model.getLineCount());
      const content = this.model.getLineContent(line);
      const first = Math.max(1, content.length - content.trimStart().length + 1);
      return {
        startLineNumber: line,
        startColumn: first,
        endLineNumber: line,
        endColumn: Math.max(first + 1, content.length + 1),
        message: finding.message,
        severity: finding.severity === "warning"
          ? monaco.MarkerSeverity.Warning
          : monaco.MarkerSeverity.Error,
      };
    }));
  }

  /** The current squiggles as data, for the debug surface. */
  lintMarkers(): { line: number; message: string; severity: string }[] {
    return monaco.editor.getModelMarkers({ resource: this.model.uri }).map((marker) => ({
      line: marker.startLineNumber,
      message: marker.message,
      severity: marker.severity === monaco.MarkerSeverity.Warning ? "warning" : "error",
    }));
  }

  /** For the debug surface: set the document WITHOUT applying - the typing path. */
  setDocument(markup: string): void {
    this.model.setValue(markup);
  }

  /** Whether this tab's form is the thing a Properties row is about: the form itself, or one of
   * the controls the projection knows. Anything else - a module, a worksheet, a control on some
   * other form - is not this document's to write. */
  owns(target: string): boolean {
    return target.toLowerCase() === this.id.module.toLowerCase() || this.kindOf(target) !== "";
  }

  /**
   * A Properties row, written to the DOCUMENT.
   *
   * THE PANEL JOINS THE TRANSACTION instead of going round it. Every other gesture on this tab
   * edits the document and lets Ctrl+S carry it to the form; the panel used to write the FORM
   * over COM and let the re-projection echo back into the text. Two things fell out of that, and
   * the owner reported both (2026-08-17): "ctrl+z is not undoing it", because the echo is pushed
   * with no stack element of its own and so leaves nothing for an undo to reverse; and "ctrl+z
   * updates the markdown editor, but not the designer", because undoing the echo left the FORM
   * still holding the value, and a draft wears the applied projection's colours over its own.
   * Writing here instead means the form never diverges, so neither can happen.
   *
   * The attribute is replaced where it stands, added before the close where it is absent, and
   * REMOVED when the value is the kind's own default - which is the dialect's rule that an
   * untouched property stays unspoken. Answers false when this document has no line to write,
   * and the caller then falls back to the host.
   */
  writeProperty(target: string, property: string, value: string): boolean {
    const line = target.toLowerCase() === this.id.module.toLowerCase()
      ? 1
      : this.headerLineOf(target);
    if (line === 0 || line > this.model.getLineCount()) {
      return false;
    }

    if (!this.spells(target, property)) {
      return false;
    }

    const text = this.model.getLineContent(line);
    const open = text.indexOf("<");
    if (open < 0) {
      return false;
    }

    // Quote-aware, because an attribute value may hold `>` and the tag does not end there.
    let scan = open + 1;
    let quoted = false;
    let end = -1;
    while (scan < text.length) {
      const character = text[scan];
      if (character === '"') {
        quoted = !quoted;
      } else if (character === ">" && !quoted) {
        end = scan;
        break;
      }
      scan += 1;
    }

    if (end < 0 || quoted) {
      return false;
    }

    // A value with a quote in it cannot be spelled in this dialect, so the host keeps that one.
    if (value.includes('"')) {
      return false;
    }

    const body = text.slice(open, end);
    const selfClosing = body.trimEnd().endsWith("/");
    const inner = selfClosing ? body.slice(0, body.lastIndexOf("/")) : body;

    const kind = target.toLowerCase() === this.id.module.toLowerCase() ? "Form" : this.kindOf(target);
    const spelled = this.isDefaultFor(kind, property, value) ? "" : ` ${property}="${value}"`;
    const already = new RegExp(`\\s*\\b${escapeForRegExp(property)}\\s*=\\s*"[^"]*"`, "i");
    const rebuilt = already.test(inner)
      ? inner.replace(already, spelled)
      : inner.trimEnd() + spelled;

    const written = text.slice(0, open) + rebuilt.trimEnd()
      + (selfClosing ? " />" : ">") + text.slice(end + 1);
    if (written === text) {
      return true;
    }

    this.model.pushStackElement();
    this.model.pushEditOperations([], [{
      range: new monaco.Range(line, 1, line, this.model.getLineMaxColumn(line)),
      text: written,
    }], () => null);

    this.lintNow();
    return true;
  }

  /** The GEOMETRY four, which the printer writes itself rather than taking from the vocabulary's
   * property list - so they are spellable without appearing there. */
  private static readonly PLACEMENT = new Set(["left", "top", "width", "height"]);

  /**
   * Whether a Properties row about this target is one this document can carry - which is the
   * question BOTH the routing and the debug act have to answer the same way, so they ask here.
   *
   * ONLY WHAT THE DIALECT CAN SPELL. A panel row is not always an attribute: `Font.Name` is a row
   * and is not a property of the markup - the dialect prints the font differently and its own
   * vocabulary is the list of what it knows - and an attribute called `Font.Name` is not even a
   * legal name, so writing one would make a document the apply refuses and a control that never
   * wears the font (caught by the suite's font-picker row, 2026-08-17). Anything not named here
   * falls through to the host, which is where it always went.
   */
  spells(target: string, property: string): boolean {
    if (!this.owns(target) || !/^[A-Za-z_]\w*$/.test(property)) {
      return false;
    }

    if (DesignerView.PLACEMENT.has(property.toLowerCase())) {
      return true;
    }

    const kind = target.toLowerCase() === this.id.module.toLowerCase() ? "Form" : this.kindOf(target);
    const known = markupVocabulary().find((row) => row.kind.toLowerCase() === kind.toLowerCase());
    return known?.properties.some((row) => row.name.toLowerCase() === property.toLowerCase()) === true;
  }

  /** Whether a value is what an untouched control of that kind already holds, as the host
   * measured it. Only an exact match counts: spelling a default costs a line the canonical print
   * takes back out at the next save, while dropping a property that was NOT the default would
   * lose the developer's edit. */
  private isDefaultFor(kind: string, property: string, value: string): boolean {
    const known = markupVocabulary().find((row) => row.kind.toLowerCase() === kind.toLowerCase());
    const shape = known?.properties.find((row) => row.name.toLowerCase() === property.toLowerCase());
    return shape?.default !== null && shape?.default !== undefined
      && shape.default.toLowerCase() === value.toLowerCase();
  }

  /*
   * The language service, read at a place in the document. Each of these puts the CARET there
   * first, because that is what a developer asking the same question has done - Ctrl+Space and
   * a hover both happen somewhere - and then asks the real provider. No copy of the answer
   * exists for probes to read: what comes back is what the widget would show.
   */

  /** An aim inside the document, whatever a caller asked for: monaco throws on a line past the
   * end, and a probe that mistyped a number deserves an answer about the nearest real place
   * rather than a stack trace from the surface. */
  private aim(line: number, column: number): monaco.IPosition {
    const lineNumber = Math.min(Math.max(1, Math.trunc(line) || 1), this.model.getLineCount());
    return {
      lineNumber,
      column: Math.min(Math.max(1, Math.trunc(column) || 1), this.model.getLineMaxColumn(lineNumber)),
    };
  }

  /** What the completion widget would offer here. */
  completions(line: number, column: number): {
    label: string; detail: string | null; documentation: string | null; insert: string;
    replaces: { from: number; to: number };
  }[] {
    const position = this.aim(line, column);
    this.editor.setPosition(position);
    return completionsAt(this.model, position).suggestions.map((item) => {
      // What accepting would REPLACE, in columns. A suggestion is an insert and a range, and the
      // range is half the answer: a font face offered where the developer has already typed
      // `"Tah` has to take the quote with it or the line gains a second one.
      const range = "startColumn" in item.range
        ? item.range
        : item.range.insert;
      return {
        label: typeof item.label === "string" ? item.label : item.label.label,
        detail: item.detail ?? null,
        documentation: typeof item.documentation === "string"
          ? item.documentation
          : item.documentation?.value ?? null,
        insert: item.insertText,
        replaces: { from: range.startColumn, to: range.endColumn },
      };
    });
  }

  /** What a hover here would say, one entry per block of the card. */
  hover(line: number, column: number): string[] {
    const position = this.aim(line, column);
    this.editor.setPosition(position);
    const answer = hoverAt(this.model, position);
    return (answer?.contents ?? []).map((content) => content.value);
  }

  /** The header hint here: the grammar it shows, and which clause it is pointing at. */
  headerHint(line: number, column: number): { label: string; active: number; parameter: string } | null {
    const position = this.aim(line, column);
    this.editor.setPosition(position);
    const help = headerHintAt(this.model, position);
    const signature = help?.value.signatures[0];
    if (!help || !signature) {
      return null;
    }

    const active = help.value.activeParameter ?? 0;
    const parameter = signature.parameters[active]?.label;
    return {
      label: signature.label,
      active,
      parameter: typeof parameter === "string" ? parameter : "",
    };
  }

  private onApplied(outcome: FormMarkupApplied): void {
    this.pendingActOutcome?.(outcome);
    this.pendingActOutcome = null;

    const saveNext = this.pendingSave;
    this.pendingSave = null;

    if (outcome.ok) {
      this.errorStrip.hidden = true;
      // The fresh projection that follows is this apply's canonical print; adopt it even
      // though the dot just cleared.
      this.awaitingAdopt = true;
      if (this.dirty) {
        this.dirty = false;
        this.deps.dirtyChanged(false);
      }

      // The second half of Ctrl+S: what just landed on the form reaches the file. Only
      // after an OK - a refused apply saves nothing, because the file would then hold a
      // form the developer was just told did not take their document, and a refused F5
      // launches nothing rather than launching yesterday's form.
      if (saveNext) {
        this.deps.saveWorkbook(saveNext.run);
      }

      return;
    }

    // The document keeps the developer's text; the canvas will show what actually landed
    // (the projection follows a partial apply too); the strip explains the gap.
    this.errorStrip.textContent = outcome.refused ?? "the apply was refused";
    this.errorStrip.hidden = false;
  }

  /** Called when the tab goes on screen: every show asks the host again, so the picture
   * follows edits made elsewhere (the api, the native designer) while it was hidden. */
  shown(): void {
    this.layout();
    this.request();
  }

  layout(): void {
    this.editor.layout();
  }

  /** Adopts one projection answer: the text into the document, the spec onto the canvas.
   * The CANVAS always follows - it is the truth of the form. The DOCUMENT follows unless the
   * developer holds unapplied edits in it, because an echo must not eat their typing; their
   * dirty document is then measured against the NEW canonical, so typing the form's own text
   * back clears the dot honestly. */
  update(payload: FormMarkupPayload): void {
    // The handlers are re-asked on EVERY projection, not only when the controls change: a
    // projection is what arrives after a rename, an add, a delete and an apply, and a Sub written
    // in the code half lands on the next one. Cheap enough to be unconditional - the outline is
    // the same call the tree's unfolded list makes, and the answer is a list of names.
    void this.paintHandlers();

    if (payload.markup === null) {
      this.notice.textContent = payload.reason ?? "the form could not be read";
      this.notice.hidden = false;
      return;
    }

    this.notice.hidden = true;

    const adopt = this.awaitingAdopt || !this.dirty;
    const first = this.canonical === null;
    // Whether this projection is an APPLY'S OWN canonical print, read before the flag is
    // consumed below. `onApplied` is the only thing that sets it. See the panel refresh at the end.
    const afterApply = this.awaitingAdopt;
    this.awaitingAdopt = false;
    this.canonical = payload.markup;

    if (adopt && this.model.getValue() !== payload.markup) {
      /*
       * THE PROJECTION MUST NOT MOVE THE SELECTION.
       *
       * Replacing the text moves the caret - monaco restores it by POSITION, and a form whose
       * lines have shifted puts that position on somebody else's line - and a caret landing on
       * another control's block selects that control. So a projection arriving while a developer
       * had a control selected quietly re-selected whatever was now at that line: measured on the
       * suite's own form, where the tab-order dialog opened on the form after a row had selected
       * a control inside a Frame.
       *
       * The guard is the one the caret-follow already respects, and the caret is put back on the
       * selected thing's own line afterwards.
       */
      const held = this.selectedName;
      this.followingSelection = true;
      const state = this.editor.saveViewState();
      if (first) {
        // The FIRST projection fills an empty document, and it lands by setValue - which
        // clears the undo stack - so there is no "back to empty" step underneath everything
        // else. As an edit it left one, and a Ctrl+Z at the floor of the stack blanked the
        // whole document while the canvas kept showing the form (found live, and by the
        // owner in the same minute, 2026-08-15).
        this.model.setValue(payload.markup);
      } else {
        // A single undoable edit rather than setValue: the developer's undo stack survives the
        // canonical print landing after their own apply. It deliberately gets NO stack stop of
        // its own, so it rides in the element of whatever edit it is echoing - the print of a
        // form the machine reads back as 50.000025 must not become an undo step the developer
        // has to walk through to reach their own move.
        this.model.pushEditOperations([], [{
          range: this.model.getFullModelRange(),
          text: payload.markup,
        }], () => null);
      }

      if (state) {
        this.editor.restoreViewState(state);
      }

      // Back on the selected thing's own line, whatever the restore made of the position, and
      // only then does the caret start choosing selections again.
      if (held !== null && held !== "") {
        this.revealInMarkup(held);
      }

      this.followingSelection = false;
    }

    const nowDirty = this.model.getValue() !== this.canonical;
    if (nowDirty !== this.dirty) {
      this.dirty = nowDirty;
      this.deps.dirtyChanged(nowDirty);
    }

    /*
     * AND THE PANEL IS REFRESHED WHEN AN APPLY LANDS, because nothing else refreshes it now.
     *
     * The Properties panel used to be republished by the host as a side effect of the write it
     * received; since the panel writes the DOCUMENT instead (#68) the host is never asked, so the
     * row went on showing whatever the developer typed. Type `&H00C0FFC0&` into a colour and the
     * row kept saying that long after the save had made it `#c0ffc0` - the panel stale against the
     * form it is supposed to be showing, swatch and all.
     *
     * MEASURED OFF `awaitingAdopt`, which is the one flag that means an apply asked for this
     * print. Two earlier readings of "an apply landed" were both wrong and both silent, and the
     * log said so each time - `round` unmoved across an apply while an explicit reselect corrected
     * the row at once. The dirty EDGE in this method never fires, because replacing the text above
     * raises the content handler synchronously and that handler takes the edge first; and `dirty`
     * read on ENTRY never fires either, because `onApplied` clears it the moment the apply is
     * accepted, well before the projection arrives.
     */
    if (afterApply) {
      this.deps.selection(this.selectedName === "" ? null : this.selectedName);
    }

    this.lastPayload = payload;
    this.setDraftShown(false);
    this.renderCanvas(payload);
  }

  /**
   * The canvas follows the DOCUMENT while it is dirty (the owner, 2026-08-14: "if I update
   * in the markdown pane, it doesn't reflect in the xlide form designer"): the draft the
   * lint round trip parsed - the strict parser, the apply's own grammar, host-side so there
   * is exactly one - renders in place of the applied projection, dressed in that
   * projection's display extras by name so the preview stays steady instead of flickering
   * between dressed and bare. A draft that does not parse keeps the last picture, because a
   * half-typed line must not blank the form; a document back at canonical puts the applied
   * projection back. The FORM is untouched throughout - Ctrl+S is still the only apply.
   */
  private previewDraft(findings: FormMarkupLintFinding[], draft: FormMarkupDraft | null): void {
    if (!this.dirty) {
      if (this.draftShown && this.lastPayload) {
        this.setDraftShown(false);
        this.renderCanvas(this.lastPayload);
      }
      return;
    }

    if (draft === null || findings.some((finding) => finding.severity === "error")) {
      return;
    }

    this.setDraftShown(true);
    this.renderCanvas(this.dressDraft(draft));
  }

  /** The draft state wears TWO cues and no third: the tab's own unsaved dot, and the amber
   * dashed outline this class puts around the form. A banner saying the same thing in words
   * stood here until 2026-08-15, when the owner retired it - the dot already says it, and a
   * strip across the top of the canvas is a lot of room for a sentence nobody needs twice. */
  private setDraftShown(shown: boolean): void {
    this.draftShown = shown;
    this.canvasScroll.classList.toggle("draft", shown);
  }

  /** A draft wearing the last applied projection's display extras, matched by name. */
  private dressDraft(draft: FormMarkupDraft): FormMarkupPayload {
    const applied = this.lastPayload;
    const wardrobe = new Map((applied?.controls ?? []).map((row) => [row.name.toLowerCase(), row]));

    const controls = draft.controls.map((row) => {
      const worn = wardrobe.get(row.name.toLowerCase());
      return worn
        ? {
          ...row,
          // THE DRAFT WINS WHERE IT SPEAKS. The rule was always "an UNSPOKEN property keeps the
          // applied one" - that is the dialect's own rule and the sentence below says so - but
          // this took the applied value either way, so a property the document DID spell could
          // never reach the canvas while the document was dirty. Typing a colour showed nothing,
          // and so did the Properties panel once its edits became document edits (2026-08-17).
          fontName: row.fontName ?? worn.fontName ?? null,
          fontSize: row.fontSize ?? worn.fontSize ?? null,
          fontBold: row.fontBold ?? worn.fontBold ?? null,
          fontItalic: row.fontItalic ?? worn.fontItalic ?? null,
          backColor: row.backColor ?? worn.backColor ?? null,
          foreColor: row.foreColor ?? worn.foreColor ?? null,
          insideWidth: worn.insideWidth ?? null,
          insideHeight: worn.insideHeight ?? null,
          tabs: row.tabs ?? worn.tabs ?? null,
          // The PICTURE is worn like the rest of them, and it has to be: a drag makes the
          // document dirty, the draft renders in place of the applied projection, and a draft
          // that does not carry a picture blanks every image on the form for as long as the
          // document is unsaved (measured 2026-08-16 - the model kept it, the canvas lost it).
          picture: row.picture ?? worn.picture ?? null,
        }
        : row;
    });

    // The dialect's own rule dresses the form box: an unspoken colour keeps the applied
    // one, exactly as an apply would. The derived chrome (insides) carries over only while
    // the draft keeps the applied OUTER size - resized, the real insets are unknown until
    // the apply answers, and the canvas falls back to its floating title bar honestly.
    const appliedForm = applied?.form ?? null;
    const sameSize = appliedForm !== null && draft.form !== null
      && appliedForm.width === draft.form.width && appliedForm.height === draft.form.height;
    const form = draft.form
      ? {
        ...draft.form,
        backColor: draft.form.backColor ?? appliedForm?.backColor ?? null,
        foreColor: draft.form.foreColor ?? appliedForm?.foreColor ?? null,
        insideWidth: sameSize ? appliedForm?.insideWidth ?? null : null,
        insideHeight: sameSize ? appliedForm?.insideHeight ?? null : null,
        picture: appliedForm?.picture ?? null,
      }
      : appliedForm;

    return { markup: null, reason: null, form, controls };
  }

  /** The canvas as rendered, for the harness: which picture stands (draft or applied), whether
   * the document holds unapplied edits, every control's name with its placed position in
   * POINTS - the document's own unit, so a row can compare the two - what is selected, where
   * the markup caret sits, and which LINES the selection's block covers. */
  canvasSnapshot(): {
    draft: boolean;
    dirty: boolean;
    undoable: boolean;
    selected: string | null;
    markupLine: number;
    markupBlock: { from: number; to: number } | null;
    group: string[];
    controls: { name: string; left: number; top: number; width: number; height: number }[];
    containers: { name: string; kind: string; tabs: string[]; open: number; page: string }[];
    pictures: { name: string; bytes: number; where: string }[];
  } {
    return {
      // Every tabbed container the canvas drew, its tabs as the strip labels them, and which
      // one is OPEN - read off the strip rather than off the map behind it, so a row sees what
      // the developer sees rather than what the view intended.
      containers: [...this.canvasScroll.querySelectorAll<HTMLElement>(".dc-page-strip")].map((strip) => {
        const tabs = [...strip.querySelectorAll<HTMLElement>(".dc-page-tab")];
        const open = tabs.findIndex((tab) => tab.classList.contains("current"));
        return {
          name: strip.dataset.container ?? "",
          kind: strip.dataset.kind ?? "",
          tabs: tabs.map((tab) => tab.textContent ?? ""),
          open,
          page: tabs[open]?.dataset.page ?? "",
        };
      }),
      draft: this.draftShown,
      dirty: this.dirty,
      // Whether the DOCUMENT has a gesture to give back. The canvas's own Ctrl+Z is only as
      // good as this, and a row that waits for an undo to land wants to know the difference
      // between "undone and re-rendered" and "there was nothing on the stack".
      undoable: this.model.canUndo(),
      selected: this.selectedName,
      markupLine: this.editor.getPosition()?.lineNumber ?? 0,
      // The block the canvas selection lights up, asked of the DOCUMENT rather than counted
      // off the screen: monaco renders a decoration on the next frame and only for the lines
      // in view, so a DOM count measures its renderer and the scroll position (which is how
      // the first version of this row read two lines for a three-line block).
      markupBlock: this.selectedName === null ? null : this.blockOf(this.selectedName),
      // The whole selection, anchor first. `selected` stays the anchor's own name, because that
      // is what every row written before groups existed asks about - and what the panel follows.
      group: this.selection(),
      controls: [...this.canvasScroll.querySelectorAll<HTMLElement>(".dc")].map((el) => ({
        name: el.dataset.control ?? "",
        ...this.inPoints({
          left: Number.parseFloat(el.style.left || "0"),
          top: Number.parseFloat(el.style.top || "0"),
          width: el.offsetWidth,
          height: el.offsetHeight,
        }),
      })),
      pictures: this.picturesDrawn(),
    };
  }

  /**
   * Every picture the canvas is actually drawing, by the control wearing it.
   *
   * SIZE rather than the bytes, for the reason the properties snapshot gives: a data URI is a
   * whole bitmap in base64 and a snapshot is read on every wait. The size proves a picture is
   * there and that it CHANGED, `where` proves it was placed rather than dropped in the middle,
   * and the pixels themselves are proved against the running form's photograph, which is the
   * only place that can prove them at all.
   *
   * The form's own picture answers to the empty name, the way the form does everywhere else.
   */
  private picturesDrawn(): { name: string; bytes: number; where: string }[] {
    const drawn: { name: string; bytes: number; where: string }[] = [];

    const surfaces = this.canvasScroll.querySelectorAll<HTMLElement>(
      ".dc-form-client, .dc-frame-client, .dc-page-body, .dc");
    for (const element of surfaces) {
      const source = element.style.backgroundImage;
      if (!source.startsWith("url(")) {
        continue;
      }

      drawn.push({
        name: element.classList.contains("dc-form-client")
          ? ""
          : element.dataset.control ?? element.closest<HTMLElement>(".dc")?.dataset.control ?? "",
        bytes: source.length,
        where: `${element.style.backgroundSize} ${element.style.backgroundPosition}`.trim(),
      });
    }

    // The caption pictures are elements rather than backgrounds, and the control they belong to
    // is the box they sit in.
    for (const image of this.canvasScroll.querySelectorAll<HTMLImageElement>("img.dc-picture")) {
      const box = image.closest<HTMLElement>(".dc");
      drawn.push({
        name: box?.dataset.control ?? "",
        bytes: image.src.length,
        where: `${box?.style.flexDirection ?? ""} ${box?.style.alignItems ?? ""}`.trim(),
      });
    }

    return drawn;
  }

  /** Selects a control by name - "" is the FORM itself - dresses it in the native handles,
   * and puts the markup caret on its line: the two halves pointing at one thing. The same
   * entry a canvas click takes, so the act and the click stay one gesture. */
  /** Set while the markup is being moved to FOLLOW a selection, so the caret it lands does not
   * bounce straight back as a fresh selection. */
  private followingSelection = false;

  /**
   * @param reveal Move the markup caret onto the selected thing. TRUE for every gesture that
   * selects from somewhere else - a canvas click, the tree, an act - and FALSE when the caret is
   * where the selection CAME from, because moving it then is moving it out from under the hand
   * that is typing. Renaming a control in the markup did exactly that: a half-typed name is not
   * a name the projection knows, the walk falls back to the form, and the form's block starts at
   * line 1 - so every keystroke threw the caret to the end of the first line (the owner,
   * 2026-08-17: "it's popping carret to the end of the firs tline which is kind of jarring").
   */
  select(name: string, extend = false, reveal = true): void {
    if (extend) {
      this.extendSelection(name);
      return;
    }

    this.extras.clear();
    this.selectedName = name;

    // Opening the page first, because the element the selection dresses only exists once the
    // page holding it is the one being drawn - and the redraw dresses it on the way out.
    if (name !== "" && this.openPagesFor(name)) {
      this.redraw();
    } else {
      this.dressSelection();
    }

    this.followingSelection = true;
    try {
      this.revealInMarkup(name, reveal);
    } finally {
      this.followingSelection = false;
    }

    /*
     * The Properties panel follows the selection host-side - M4's bridgehead - but only to a name
     * THE FORM HAS.
     *
     * The panel reads the live control, so a name the document has only just invented is a name it
     * cannot find, and its answer for a control it cannot find is to fall back to the form. Rename
     * a control in the markup and that is what happened: the caret and the canvas stayed on the
     * control while the panel jumped to the UserForm (the owner, 2026-08-17, with a screenshot of
     * `NameLabel2` selected and EntryForm's rows showing). Holding the panel where it is keeps it
     * on the same physical control under the name the form still knows it by, which is the honest
     * answer until Ctrl+S makes the rename real.
     */
    const known = name === ""
      || (this.lastPayload?.controls ?? []).some((row) => row.name.toLowerCase() === name.toLowerCase());
    if (known) {
      this.deps.selection(name === "" ? null : name);
    }
  }

  /**
   * Ctrl+click: adds a control to the selection, or takes it back out.
   *
   * The FORM is never part of a group - it is the ground the group stands on, and every gesture
   * a group offers (move them together, line them up, size them alike) means nothing applied to
   * it. Taking the anchor out promotes one of the others, so a selection that still holds
   * something always has a primary.
   */
  private extendSelection(name: string): void {
    const anchor = this.selectedName;
    if (name === "" || anchor === null || anchor === "") {
      this.select(name);
      return;
    }

    const key = name.toLowerCase();
    const held = [...this.extras].find((one) => one.toLowerCase() === key);
    if (held !== undefined) {
      this.extras.delete(held);
    } else if (anchor.toLowerCase() === key) {
      const next = [...this.extras][0];
      if (next === undefined) {
        return;
      }

      this.extras.delete(next);
      this.selectedName = next;
    } else {
      // A group is one container's business: controls in different boxes are measured from
      // different origins, so moving them together or lining them up would mean nothing on
      // screen. Selecting across containers starts a new selection instead of a nonsense group.
      if (this.parentOf(name).toLowerCase() !== this.parentOf(anchor).toLowerCase()) {
        this.select(name);
        return;
      }

      this.extras.add(name);
    }

    this.dressSelection();
    this.deps.selection(this.selectedName === "" ? null : this.selectedName);
  }

  /** Everything selected, anchor first. One name is the ordinary case; the form answers as
   * itself and is never in a group. */
  private selection(): string[] {
    return this.selectedName === null ? [] : [this.selectedName, ...this.extras];
  }

  /** Which container a control belongs to in the drawn projection, by name - "" for the form's
   * own ground. Spelled as the projection spells it; callers comparing two of them fold the case
   * themselves. */
  private parentOf(name: string): string {
    return (this.shownPayload?.controls ?? [])
      .find((one) => one.name.toLowerCase() === name.toLowerCase())?.parent ?? "";
  }

  /** A canvas double-click, for the debug surface: the host's event-stub gesture. */
  requestEventStub(control: string | null): void {
    this.deps.eventStub(control);
  }

  /**
   * A press on the palette picks a kind up. From here the pointer carries a ghost of the
   * control's real size until it is dropped on the form - the toolbox gesture the native
   * designer has and the suppressed palette was suppressed for.
   *
   * The whole drag lives on the DOCUMENT window rather than on the palette, because the
   * interesting part of it happens over the canvas: capture on the button would keep the
   * events but hide which container the pointer is over, and that container is the answer the
   * drop needs.
   */
  private onToolboxPress(event: PointerEvent): void {
    const button = (event.target as HTMLElement).closest<HTMLElement>(".designer-tool");
    const tool = TOOLBOX.find((one) => one.kind === button?.dataset.kind);
    if (!tool || event.button !== 0) {
      return;
    }

    event.preventDefault();
    const ghost = document.createElement("div");
    ghost.className = "designer-tool-ghost";
    ghost.style.width = `${this.toPixels(tool.width)}px`;
    ghost.style.height = `${this.toPixels(tool.height)}px`;
    ghost.textContent = tool.kind;
    document.body.appendChild(ghost);
    this.carrying = { kind: tool.kind, width: tool.width, height: tool.height, ghost };
    this.moveGhost(event.clientX, event.clientY);

    const onMove = (moved: PointerEvent): void => this.moveGhost(moved.clientX, moved.clientY);
    const onUp = (dropped: PointerEvent): void => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      this.dropTool(dropped.clientX, dropped.clientY, dropped.altKey);
    };

    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  /** The ghost rides the pointer by its TOP-LEFT, because that is where the control will land:
   * what the eye follows is the box that is about to exist, not a cursor with a label. */
  private moveGhost(x: number, y: number): void {
    if (!this.carrying) {
      return;
    }

    this.carrying.ghost.style.left = `${x}px`;
    this.carrying.ghost.style.top = `${y}px`;
    const over = this.containerAt(x, y);
    this.carrying.ghost.classList.toggle("over", over !== null);
  }

  /**
   * Where a drop would land: the deepest CONTAINER under the pointer - a Frame's client, a
   * MultiPage's page body, or the form's own client area - with the name it goes under and the
   * point inside it, in points. Null when the pointer is not over this form at all, which is
   * how a drop outside becomes nothing rather than a control at 0,0.
   */
  private containerAt(x: number, y: number): { parent: string; left: number; top: number } | null {
    const hit = document.elementFromPoint(x, y);
    const client = hit?.closest<HTMLElement>(".dc-frame-client, .dc-page-body, .dc-form-client");
    if (!client || !this.canvasScroll.contains(client)) {
      return null;
    }

    // A Frame's client belongs to the Frame; a page body IS the page whose content it draws and
    // says so on itself, so a drop lands on the page the developer is LOOKING at; the form's
    // client belongs to the form.
    const box = client.getBoundingClientRect();
    const parent = client.classList.contains("dc-page-body")
      ? client.dataset.control ?? ""
      : client.closest<HTMLElement>(".dc")?.dataset.control ?? "";
    return { parent, left: this.toPoints(x - box.left), top: this.toPoints(y - box.top) };
  }

  /** Which of a container's tabs is open, as an index into its headers: the page the developer
   * opened while it is still there, the tab they picked while the strip is still that long, and
   * the first one otherwise. */
  private openTabOf(row: FormMarkupControl, pages: FormMarkupControl[], count: number): number {
    const key = row.name.toLowerCase();
    if (row.type === "TabStrip") {
      return Math.min(Math.max(0, this.shownTab.get(key) ?? 0), Math.max(0, count - 1));
    }

    const wanted = (this.shownPage.get(key) ?? "").toLowerCase();
    return Math.max(0, pages.findIndex((page) => page.name.toLowerCase() === wanted));
  }

  /**
   * Paints the current zoom: the form scaled from its top-left corner, and a stage sized to what
   * that comes to, so the canvas scrolls over the whole of it.
   */
  private applyZoom(): void {
    const form = this.canvasScroll.querySelector<HTMLElement>(".dc-form");
    const stage = this.canvasScroll.querySelector<HTMLElement>(".dc-stage");
    if (!form || !stage) {
      return;
    }

    form.style.transformOrigin = "top left";
    form.style.transform = this.zoom === 1 ? "" : `scale(${this.zoom})`;
    stage.style.width = `${form.offsetWidth * this.zoom}px`;
    stage.style.height = `${form.offsetHeight * this.zoom}px`;
    this.zoomButton.textContent = `${Math.round(this.zoom * 100)}%`;
  }

  /**
   * Sets the zoom - a factor, or "fit" for the largest that shows the whole form in the canvas as
   * it stands. Answers what it settled on, which for Fit is the only way to know.
   */
  setZoom(what: number | "fit"): string {
    const form = this.canvasScroll.querySelector<HTMLElement>(".dc-form");
    if (!form) {
      return "the canvas is not showing a form";
    }

    if (what === "fit") {
      const room = this.canvasScroll.getBoundingClientRect();
      // The margin the stage carries, both sides, taken off before the ratio - a form fitted to
      // the raw box would sit with its right edge under the scroll bar.
      const margin = 48 + 16;
      const wide = (room.width - margin) / Math.max(1, form.offsetWidth);
      const tall = (room.height - margin) / Math.max(1, form.offsetHeight);
      this.zoom = Math.max(0.25, Math.min(2, Math.min(wide, tall)));
    } else {
      this.zoom = Math.max(0.25, Math.min(4, what));
    }

    this.applyZoom();
    return `zoom ${Math.round(this.zoom * 100)}%`;
  }

  /** The zoom the canvas is drawing at, as a percentage - for the harness and the button. */
  zoomPercent(): number {
    return Math.round(this.zoom * 100);
  }

  /** Re-draws the picture that is up - the draft while one stands, the applied projection
   * otherwise - with no round trip: opening a page changes what is DRAWN, never what the form
   * or the document holds. */
  private redraw(): void {
    if (this.shownPayload) {
      this.renderCanvas(this.shownPayload);
    }
  }

  /**
   * Opens whatever pages a thing sits inside, so that selecting it shows it - a control on page
   * two of a MultiPage nested in page one of another opens both. Answers whether anything
   * changed, because the caller has to redraw if it did.
   *
   * This is what keeps the two halves pointing at one thing (the markup caret and the canvas):
   * a caret landing in page two's block opens page two rather than selecting nothing, which is
   * what happened while the canvas only ever drew the first.
   */
  private openPagesFor(name: string): boolean {
    const rows = this.shownPayload?.controls ?? [];
    const byName = new Map(rows.map((row) => [row.name.toLowerCase(), row]));

    let changed = false;
    let row = byName.get(name.toLowerCase());
    // Bounded by the chain itself: every step moves to a parent, and a cycle would have to be
    // in the projection, which is a tree by construction.
    for (let guard = 0; row && guard < 32; guard++) {
      const parent = row.parent ? byName.get(row.parent.toLowerCase()) : undefined;
      if (row.type === "Page" && parent) {
        const key = parent.name.toLowerCase();
        if ((this.shownPage.get(key) ?? "").toLowerCase() !== row.name.toLowerCase()) {
          this.shownPage.set(key, row.name);
          changed = true;
        }
      }

      row = parent;
    }

    return changed;
  }

  /**
   * Opens the page a tab stands for, and selects it - which is the native designer's own
   * gesture: clicking a tab shows that page and puts the PAGE in the Properties window.
   *
   * A TabStrip's tab selects the TabStrip instead, because there is no page object behind it,
   * and shows nothing new: the runtime draws the same controls under every tab of a TabStrip,
   * so a canvas that swapped content there would be inventing a control it does not have.
   */
  private openTab(tab: HTMLElement): string {
    const strip = tab.closest<HTMLElement>(".dc-page-strip");
    const container = strip?.dataset.container;
    if (!container) {
      return "that tab belongs to nothing on the canvas";
    }

    const page = tab.dataset.page;
    if (page) {
      this.shownPage.set(container.toLowerCase(), page);
    } else {
      this.shownTab.set(container.toLowerCase(), Number(tab.dataset.tab ?? 0));
    }

    this.redraw();
    this.select(page ?? container);
    return `opened ${page ?? `tab ${Number(tab.dataset.tab ?? 0) + 1} of ${container}`}`;
  }

  /**
   * The tab strip's own menu, where the native designer keeps New Page and Delete Page.
   *
   * A TabStrip gets no menu at all: its tabs are not in the dialect - they ride the walk as
   * strings for painting and there is no line in the document to add or take away - so every
   * item would be a lie about what this can do. The gap is the dialect's, not the menu's.
   */
  private onCanvasMenu(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    const strip = target.closest<HTMLElement>(".dc-page-strip");
    if (!strip?.dataset.container) {
      this.showArrangeMenu(event);
      return;
    }

    event.preventDefault();
    const container = strip.dataset.container;

    // The right-click OPENS the tab under it first, the way every menu in this product acts on
    // the thing it marked: Delete Page then names the page the developer is looking at.
    const tab = target.closest<HTMLElement>(".dc-page-tab");
    if (tab?.dataset.tab !== undefined) {
      this.openTab(tab);
    }

    // Delete Page acts on the page that is OPEN, whether the right-click landed on its tab or
    // on the empty end of the strip: what a developer means by "this page" is the one they can
    // see. Disabled only for a MultiPage that has no pages at all, which is a document the
    // dialect allows and the canvas draws as bare chrome.
    //
    // Read off the canvas AFTER the open above, never off `strip`: opening a tab redraws, which
    // leaves the element this event arrived on detached and still wearing the old picture's
    // marks - so the menu named the previously open page and deleted that one instead.
    // A MultiPage holds PAGES and a TabStrip holds TABS, and since the dialect learned both they
    // are one gesture with two words: a line of the child's kind under the container's, one
    // undoable edit, on the form at Ctrl+S through the apply's own diff.
    const what = strip.dataset.kind === "TabStrip" ? "Tab" : "Page";
    const open = this.openChildOf(container);
    showContextMenu(event.clientX, event.clientY, [
      { label: `New ${what}`, run: () => { this.addChild(container, what); } },
      {
        label: `Delete ${what}`,
        enabled: open !== "",
        run: () => { this.deleteFromDocument(open); },
      },
    ]);
  }

  /**
   * The canvas's own menu over a selection: the native Format menu's arrange group, which has no
   * other home here - this product has no menu bar, and the editor's own Format menu stays
   * suppressed because it would act on the NATIVE designer's selection rather than on ours.
   *
   * Offered only where it means something. One control cannot be lined up with anything, so the
   * menu does not appear for it: an item that cannot run is worse than no menu at all when the
   * whole gesture is "do this to these".
   */
  private showArrangeMenu(event: MouseEvent): void {
    const target = event.target as HTMLElement;
    const control = target.closest<HTMLElement>(".dc")?.dataset.control;
    if (control !== undefined
      && !this.selection().some((one) => one.toLowerCase() === control.toLowerCase())) {
      // A right-click somewhere else in the form picks that control first, the way every menu in
      // this product acts on the thing it marked.
      this.select(control);
    }

    const chosen = this.selection().filter((one) => one !== "");
    if (chosen.length === 0) {
      return;
    }

    event.preventDefault();
    const item = (label: string, how: string): ContextMenuItem =>
      ({ label, run: () => { this.arrange(how); } });

    // DEPTH is offered for one control as well as for a group, and it is the one item here that
    // does not write the document: MSForms' collection is not in z-order, so the dialect has no
    // way to say this and the model takes it directly. See the bridge call's own note.
    const depth: ContextMenuItem[] = [
      {
        label: "Bring to Front",
        run: () => { this.zorder(chosen, true); },
      },
      {
        label: "Send to Back",
        run: () => { this.zorder(chosen, false); },
      },
    ];

    /*
     * THE CLIPBOARD FOUR, on the menu as well as on the keys.
     *
     * This menu offered depth, tab order and the align set and none of these, and Delete was on
     * the keyboard alone - the same discoverability problem the tab-order strip button had, and
     * the owner named both in one breath. A gesture reachable only by a key a developer has to
     * already know is a gesture most developers never find.
     */
    const clipboard: ContextMenuItem[] = [
      {},
      { label: "Cut", run: () => { this.say(this.cutSelection()); } },
      { label: "Copy", run: () => { this.say(this.copySelection()); } },
      {
        label: "Paste",
        enabled: this.clipboard !== null,
        run: () => { this.say(this.pasteClipboard()); },
      },
      { label: "Duplicate", run: () => { this.say(this.duplicateSelection()); } },
      { label: "Delete", run: () => { this.deleteSelection(); } },
    ];

    // The three that act on each control on its own, so they are offered for ONE as much as for
    // a group - unlike align and distribute below, which need something to line up against.
    const shaping: ContextMenuItem[] = [
      {},
      { label: "Center in Container (across)", run: () => { void this.sayLater(this.format("centreX")); } },
      { label: "Center in Container (down)", run: () => { void this.sayLater(this.format("centreY")); } },
      { label: "Size to Fit", run: () => { void this.sayLater(this.format("fit")); } },
      { label: "Size to Grid", run: () => { void this.sayLater(this.format("grid")); } },
    ];

    const order: ContextMenuItem[] = [
      {},
      { label: "Tab Order...", run: () => { this.showTabOrder(); } },
    ];

    if (chosen.length < 2) {
      showContextMenu(event.clientX, event.clientY, [...depth, ...clipboard, ...shaping, ...order]);
      return;
    }

    showContextMenu(event.clientX, event.clientY, [
      ...depth,
      ...clipboard,
      ...shaping,
      ...order,
      {},
      item("Align Left", "left"),
      item("Align Center", "centreX"),
      item("Align Right", "right"),
      item("Align Top", "top"),
      item("Align Middle", "centreY"),
      item("Align Bottom", "bottom"),
      {},
      item("Same Width", "width"),
      item("Same Height", "height"),
      {},
      item("Space Across", "across"),
      item("Space Down", "down"),
    ]);
  }

  /**
   * The tab-order dialog for the container the selection sits in - the form's own controls when
   * nothing or the form is selected.
   *
   * ONE container at a time, because tab order is per container in MSForms: a Frame's children
   * have their own 0..n and the Tab key walks into the frame and out again. Showing a form's
   * whole tree in one list would be a list whose numbers repeat.
   */
  showTabOrder(): string {
    const anchor = this.selectedName;
    const container = anchor === null || anchor === ""
      ? ""
      : this.kindOf(anchor) === "Page" || this.isContainer(anchor)
        ? anchor
        : this.parentOf(anchor);

    // A control with NO TabIndex is not in the tab order at all - an Image cannot take focus, and
    // MSForms answers null rather than a number for one. Listing it would put a row in the
    // sequence that the Tab key never visits, and every row below it would be numbered wrong.
    const controls = (this.shownPayload?.controls ?? [])
      .filter((row) => (row.parent ?? "").toLowerCase() === container.toLowerCase()
        && row.type !== "Page"
        && typeof row.tabIndex === "number")
      .map((row) => ({
        name: row.name,
        type: row.type,
        tabIndex: row.tabIndex ?? 0,
        // A control in the order that Tab does not STOP on - a Label, or anything the developer
        // has set TabStop false on to take it out of the walk without taking it out of the
        // sequence. The dialog dims it rather than dropping it, because the place is real: an
        // accelerator moves focus to whatever comes after.
        stops: row.tabStop !== false,
      }));

    if (controls.length === 0) {
      return `${container === "" ? "the form" : container} holds nothing to order`;
    }

    // Which container this decided on, and from what: a row that expected a Frame's list and got
    // the form's has no way to tell whether the selection or the projection was the surprise.
    const from = anchor === null || anchor === ""
      ? "the form is selected"
      : `${anchor} is a ${this.kindOf(anchor) || "control the canvas does not know"}`;

    showTabOrder({
      container,
      controls,
      setIndex: (control, index) => { this.deps.setProperty(control, "TabIndex", String(index)); },
    });
    return `tab order for ${container === "" ? "the form" : container}: ${controls.length} control(s)`
      + ` (${from})`;
  }

  /** Whether the drawn projection calls this a container - a Frame or a MultiPage. */
  private isContainer(name: string): boolean {
    const kind = this.kindOf(name);
    return kind === "Frame" || kind === "MultiPage";
  }

  /**
   * Bring to Front / Send to Back, straight at the MODEL.
   *
   * Every other canvas gesture writes the document and waits for Ctrl+S. This one cannot: the
   * Controls collection a projection walks is not in z-order and does not move when ZOrder is
   * called (measured 2026-08-16 - two overlapping labels swapped which was on top on the running
   * form, while the walk's order never changed), so there is nothing for the dialect to print and
   * nothing for the canvas to draw. It behaves like a Properties panel edit instead: the model
   * takes it at once, and the developer sees it when the form runs.
   */
  zorder(names: string[], front: boolean): string {
    if (names.length === 0) {
      return "nothing is selected";
    }

    // Front-most last, so a group keeps its own order when it arrives at the front - and the
    // reverse when it goes to the back, for the same reason.
    for (const name of front ? names : [...names].reverse()) {
      this.deps.zorder(name, front);
    }

    return `${names.join(", ")} to the ${front ? "front" : "back"}`;
  }

  /**
   * Lines a selection up on its ANCHOR, sizes it to the anchor, or spreads it evenly - the
   * native designer's arrange vocabulary, applied to the DOCUMENT as one undoable edit like
   * every other canvas gesture.
   *
   * The anchor is the reference and never moves, which is the native rule: what a developer
   * means by "align left" is "line these up with THAT one", and the one they mean is the one
   * they picked first (or the one the marquee caught first).
   */
  /**
   * The three Format gestures that are not about lining controls up WITH EACH OTHER, and so are
   * not `arrange`: each acts on every selected control independently, and each is meaningful for
   * ONE (2026-08-18 - the native Format menu has all three and this had none of them).
   *
   *   centreX / centreY  centres within the control's own CONTAINER's client area, which for a
   *                      control inside a Frame is the Frame - the same containment rule align
   *                      and distribute keep, and the reason this is not "centre in the form".
   *   fit                the control's size to what it DRAWS. Measured off the canvas's own
   *                      rendering with a Range over the element's contents, which catches the
   *                      glyph of a check box as well as the caption beside it, plus the padding
   *                      and border the kind wears. MSForms calls this AutoSize; going through
   *                      the host to ask it would take the gesture out of the document, so the
   *                      picture the canvas already draws in the right font is what answers.
   *   grid               rounds the box out to the grid the snap setting sizes, whether or not
   *                      snapping is switched ON - the grid has a size either way, and "size to
   *                      grid" is a gesture rather than a mode.
   */
  async format(how: string): Promise<string> {
    const names = this.selection().filter((one) => one !== "");
    if (names.length === 0) {
      return "nothing is selected";
    }

    const step = Math.max(1, currentSettings().designerGridSize);
    const moves: { name: string; box: Box; sized: boolean }[] = [];

    for (const name of names) {
      const header = this.headerOf(name);
      const element = this.elementFor(name);
      if (!header || !element) {
        return "a line in the selection is not one this can rewrite";
      }

      const box = this.baseBox(header, element);
      const wanted = { ...box };

      if (how === "centreX" || how === "centreY") {
        const room = this.roomFor(element);
        if (!room) {
          return `${name} is not inside anything to centre it in`;
        }

        if (how === "centreX") {
          wanted.left = Math.max(0, (room.width - box.width) / 2);
        } else {
          wanted.top = Math.max(0, (room.height - box.height) / 2);
        }
      } else if (how === "fit") {
        /*
         * THE HOST MEASURES THIS ONE, and the numbers are why. A page-side read of the caption's
         * ink was written first and driven against MSForms' own AutoSize; it agreed for exactly
         * one kind, put "Hold" at 18pt wide - narrower than its own caption - and could not see a
         * picture at all, because the canvas draws an oversized one scaled down. See
         * FormDesignService.MeasureAutoSize for the table.
         *
         * IT COPIES THE RUNTIME EXACTLY, including where the runtime is absurd: the OK button
         * with its 256-square logo becomes 220.5x198.1, and that is the answer this writes. The
         * owner chose that over fitting the caption instead (2026-08-18), on the grounds that
         * settled the icon frame - a surface faithful in one place and clever in another is one
         * nobody can predict.
         */
        const natural = await this.deps.autoSize(name);
        if (!natural) {
          return `${name} has no natural size to fit to`;
        }

        wanted.width = Math.max(MIN_CONTROL, natural.width);
        wanted.height = Math.max(MIN_CONTROL, natural.height);
      } else if (how === "grid") {
        // OUT to the grid rather than to the nearest line, so a control never shrinks under a
        // gesture whose name says nothing about making it smaller.
        wanted.left = Math.floor(box.left / step) * step;
        wanted.top = Math.floor(box.top / step) * step;
        wanted.width = Math.max(step, Math.ceil((box.left + box.width) / step) * step - wanted.left);
        wanted.height = Math.max(step, Math.ceil((box.top + box.height) / step) * step - wanted.top);
      } else {
        return `${how} is not a format this knows`;
      }

      moves.push({
        name,
        // A FIT IS NOT CLAMPED TO THE CONTAINER. Every other gesture keeps a control inside the
        // thing that holds it, and this one must not: the whole point of copying the runtime is
        // that a button sizing to a 256-square picture comes out bigger than the form, and a
        // clamp would quietly turn the faithful answer back into a plausible one.
        box: how === "fit" ? wanted : this.clamp(wanted, this.roomFor(element), MIN_CONTROL),
        sized: how !== "centreX" && how !== "centreY",
      });
    }

    if (!this.writeBoxes(moves)) {
      return "the document refused the change";
    }

    this.lintNow();
    return `${how} on ${names.join(", ")}`;
  }

  arrange(how: string): string {
    const names = this.selection().filter((one) => one !== "");
    if (names.length < 2) {
      return "lining up takes two or more controls";
    }

    const placed = names.map((name) => {
      const header = this.headerOf(name);
      const element = this.elementFor(name);
      return header && element ? { name, box: this.baseBox(header, element), element } : null;
    });
    if (placed.some((one) => one === null)) {
      return "a line in the selection is not one this can rewrite";
    }

    const all = placed as { name: string; box: Box; element: HTMLElement }[];
    const anchor = all[0]?.box;
    if (!anchor) {
      return "lining up takes two or more controls";
    }

    const sized = how === "width" || how === "height";

    const moves = all.map(({ name, box, element }) => {
      const wanted = { ...box };
      switch (how) {
        case "left": wanted.left = anchor.left; break;
        case "right": wanted.left = anchor.left + anchor.width - box.width; break;
        case "centreX": wanted.left = anchor.left + (anchor.width - box.width) / 2; break;
        case "top": wanted.top = anchor.top; break;
        case "bottom": wanted.top = anchor.top + anchor.height - box.height; break;
        case "centreY": wanted.top = anchor.top + (anchor.height - box.height) / 2; break;
        case "width": wanted.width = anchor.width; break;
        case "height": wanted.height = anchor.height; break;
        default: break;
      }

      return {
        name,
        box: this.clamp(wanted, this.roomFor(element), MIN_CONTROL),
        sized,
      };
    });

    if (how === "across" || how === "down") {
      const across = how === "across";
      // Equal GAPS between the boxes, which is what "make spacing equal" means: the two on the
      // ends stay where they are and everything between them is spread across what is left over.
      const order = [...all].sort((a, b) =>
        across ? a.box.left - b.box.left : a.box.top - b.box.top);
      const first = order[0]?.box ?? anchor;
      const last = order[order.length - 1]?.box ?? anchor;
      const span = across
        ? (last.left + last.width) - first.left
        : (last.top + last.height) - first.top;
      const filled = order.reduce((sum, one) => sum + (across ? one.box.width : one.box.height), 0);
      const gap = (span - filled) / Math.max(1, order.length - 1);

      let run = across ? first.left : first.top;
      for (const one of order) {
        const move = moves.find((each) => each.name === one.name);
        if (move) {
          move.box = across
            ? { ...one.box, left: run }
            : { ...one.box, top: run };
        }

        run += (across ? one.box.width : one.box.height) + gap;
      }
    }

    if (!this.writeBoxes(moves)) {
      return "the document refused the move";
    }

    this.dressSelection();
    return `${how} across ${names.length} control(s)`;
  }

  /** The page or tab a container is showing, by NAME, read off the canvas as it stands now. */
  private openChildOf(container: string): string {
    const strip = [...this.canvasScroll.querySelectorAll<HTMLElement>(".dc-page-strip")]
      .find((one) => (one.dataset.container ?? "").toLowerCase() === container.toLowerCase());
    return strip?.querySelector<HTMLElement>(".dc-page-tab.current")?.dataset.name ?? "";
  }

  /** What the drawn projection calls a thing - the canvas's own answer, not the document's. */
  private kindOf(name: string): string {
    return (this.shownPayload?.controls ?? [])
      .find((row) => row.name.toLowerCase() === name.toLowerCase())?.type ?? "";
  }

  /**
   * Whether the PROJECTION holds a control of this name - which is a different question from
   * whether the canvas is drawing one.
   *
   * A control on a page that is not open is perfectly real and selecting it is what OPENS that
   * page; a control deleted three gestures ago is not, and selecting it must be refused. Asking
   * the drawing conflated the two, and the refusal that was meant for the second case caught the
   * first (measured 2026-08-16: selecting `Agree` while page two was showing answered "not on the
   * canvas to select", where a click on its markup line opens page one and selects it).
   */
  knows(name: string): boolean {
    return name === "" || this.kindOf(name) !== "";
  }

  /**
   * A new Page under a MultiPage, or a new Tab under a TabStrip: one line at the end of the
   * container's block, named and captioned the way MSForms names one, as a single undoable edit -
   * the same text a hand would have typed, which is what every canvas gesture writes.
   *
   * Opened and selected on arrival, because a page nobody can see is not what New Page means.
   * The container does not have it until Ctrl+S carries the document through the apply.
   */
  private addChild(container: string, what: "Page" | "Tab"): string {
    const kind = this.kindOf(container);
    if (what === "Page" ? kind !== "MultiPage" : kind !== "TabStrip") {
      return `${container} is not a ${what === "Page" ? "MultiPage" : "TabStrip"}`;
    }

    const host = this.headerOf(container);
    if (!host) {
      return `${container} is not a line this can add to`;
    }

    const lines = this.model.getLinesContent();
    const indent = (text: string): number => text.length - text.trimStart().length;
    const level = indent(lines[host.line - 1] ?? "");
    let after = host.line;
    while (after < lines.length && indent(lines[after] ?? "") > level) {
      after += 1;
    }

    const name = this.freeName(what);
    const column = this.model.getLineMaxColumn(after);
    this.model.pushStackElement();
    this.model.pushEditOperations([], [{
      range: new monaco.Range(after, column, after, column),
      text: `\n${" ".repeat(level + 4)}<${what} Name="${name}" Caption="${name}" />`,
    }], () => null);

    // Marked open before it is drawn: the child does not exist in the picture until the lint
    // round trip brings the parsed draft back, and the render then finds it already chosen. A
    // TabStrip is opened by index rather than by name, and the new tab is the last one.
    if (what === "Page") {
      this.shownPage.set(container.toLowerCase(), name);
    } else {
      this.shownTab.set(container.toLowerCase(), Number.MAX_SAFE_INTEGER);
    }
    this.select(name);
    this.lintNow();
    return `added ${name}`;
  }

  /** A press on the canvas: a HANDLE resizes what is already selected, anything else picks
   * what is under the pointer and arms a move. */
  private onCanvasPress(event: PointerEvent): void {
    this.canvasScroll.focus();
    if (event.button !== 0) {
      return;
    }

    const target = event.target as HTMLElement;

    // The handles come first, because they stand ON the selection's boundary: a press there is
    // a resize of the selected thing, never a pick of whatever the boundary crosses.
    const handle = target.closest<HTMLElement>(".dc-handle");
    if (handle?.dataset.edge && this.selectedName !== null) {
      this.arm(this.selectedName, handle.dataset.edge, event);
      return;
    }

    // A TAB is a switch and never a pick-up: the press opens that page and no drag is armed,
    // so a container is carried by its body or by the empty end of its strip rather than by
    // the tab a click means to open.
    const tab = target.closest<HTMLElement>(".dc-page-tab");
    if (tab?.dataset.tab !== undefined) {
      this.openTab(tab);
      return;
    }

    const control = target.closest<HTMLElement>(".dc");
    if (!control?.dataset.control) {
      // ANYWHERE on the canvas selects the FORM, not just the form's own ground. A press on the
      // grey around it used to select nothing at all and leave the last control dressed, so the
      // panel and the handles went on describing something the developer had just clicked away
      // from - and there was no way to get back to the form's own properties except to find a
      // bare patch of it (the owner, 2026-08-16: "if I click inside of the designer canvas, but
      // not on the form, I'd like the form to be selected").
      //
      // A press on a container's own ground also ARMS A RUBBER BAND: travel makes it a marquee
      // over that container's children, and a press that never travels is still the click that
      // selects the form.
      this.armBand(event);
      return;
    }

    const name = control.dataset.control;
    const extend = event.ctrlKey || event.metaKey;
    // A press on something already selected KEEPS the group, which is what lets a group be
    // dragged by any of its members - re-selecting would throw the rest away on the way to
    // moving them.
    if (extend || !this.selection().some((one) => one.toLowerCase() === name.toLowerCase())) {
      this.select(name, extend);
    }

    // Ctrl+click is a selection gesture and nothing else: no drag arms behind it, so a hand
    // gathering a group cannot nudge one of them by accident.
    if (extend) {
      return;
    }

    // A Page is not placed by coordinates - it fills its MultiPage - so there is nothing to
    // drag it by. Everything else only ARMS here: the gesture begins past the threshold, which
    // is what keeps a plain click a plain click.
    if (control.dataset.kind === "Page") {
      return;
    }

    this.arm(name, null, event);
  }

  /**
   * A press on a container's own ground: the rubber band the native designer draws, which
   * selects what it touches when it is let go.
   *
   * Armed rather than begun, like every other gesture here - below the threshold this is the
   * click that selects the form, and the band only exists once the hand has travelled.
   */
  private armBand(event: PointerEvent): void {
    const host = (event.target as HTMLElement)
      .closest<HTMLElement>(".dc-frame-client, .dc-page-body, .dc-form-client");
    if (!host || !this.canvasScroll.contains(host)) {
      this.select("");
      return;
    }

    this.band = {
      pointerId: event.pointerId,
      host,
      parent: host.classList.contains("dc-page-body")
        ? host.dataset.control ?? ""
        : host.closest<HTMLElement>(".dc")?.dataset.control ?? "",
      startX: event.clientX,
      startY: event.clientY,
      element: null,
    };

    try {
      this.canvasScroll.setPointerCapture(event.pointerId);
    } catch {
      // A synthesised pointer has no capture to take; bubbling carries the rest.
    }
  }

  /** The band follows the pointer, drawn inside the container it belongs to so it cannot stretch
   * over ground its own selection could never cover. */
  private trackBand(event: PointerEvent): void {
    const band = this.band;
    if (!band) {
      return;
    }

    const travelled = Math.abs(event.clientX - band.startX) >= DRAG_THRESHOLD
      || Math.abs(event.clientY - band.startY) >= DRAG_THRESHOLD;
    if (!band.element && !travelled) {
      return;
    }

    if (!band.element) {
      band.element = document.createElement("div");
      band.element.className = "dc-marquee";
      band.host.appendChild(band.element);
    }

    const room = band.host.getBoundingClientRect();
    const left = Math.min(band.startX, event.clientX) - room.left;
    const top = Math.min(band.startY, event.clientY) - room.top;
    band.element.style.left = `${left}px`;
    band.element.style.top = `${top}px`;
    band.element.style.width = `${Math.abs(event.clientX - band.startX)}px`;
    band.element.style.height = `${Math.abs(event.clientY - band.startY)}px`;
    event.preventDefault();
  }

  /**
   * The band is let go: everything of that container's own children the band TOUCHES is
   * selected, in the order the document holds them, and the first is the anchor.
   *
   * Touching rather than enclosing, which is MSForms' own rule and the more forgiving one - a
   * band that has to swallow a control whole cannot pick up a row of them without also swallowing
   * whatever sits below.
   */
  private dropBand(event: PointerEvent): void {
    const band = this.band;
    this.band = null;
    if (!band) {
      return;
    }

    try {
      this.canvasScroll.releasePointerCapture(band.pointerId);
    } catch {
      // Nothing was captured; nothing to give back.
    }

    if (!band.element) {
      this.select("");
      return;
    }

    const box = band.element.getBoundingClientRect();
    band.element.remove();

    const caught = [...band.host.children]
      .filter((one): one is HTMLElement => one instanceof HTMLElement
        && one.classList.contains("dc") && one.dataset.control !== undefined
        && one.dataset.kind !== "Page")
      .filter((one) => {
        const at = one.getBoundingClientRect();
        return at.left < box.right && at.right > box.left
          && at.top < box.bottom && at.bottom > box.top;
      })
      .map((one) => one.dataset.control ?? "");

    if (caught.length === 0) {
      this.select("");
      return;
    }

    this.select(caught[0] ?? "");
    for (const also of caught.slice(1)) {
      this.extras.add(also);
    }

    this.dressSelection();
    event.preventDefault();
  }

  /** Arms a gesture on a named thing - a move with no edge, a resize with one - recording the
   * box it starts from so every later frame measures from the press rather than from the last
   * one, which is what keeps a slow drag from accumulating rounding. */
  private arm(name: string, edge: string | null, event: PointerEvent): void {
    const element = this.elementOf(name);
    const header = this.headerOf(name);
    if (!element || !header) {
      return;
    }

    const parent = this.hostNameOf(element);
    const grabbed = element.getBoundingClientRect();
    this.drag = {
      name,
      element,
      edge,
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      // In POINTS, off the DOCUMENT - never off the painted box, whose width carries whatever
      // border the renderer drew. Measured live: a form dragged 20 points wider grew 21,
      // because its element is two pixels bigger than the size its own line spells.
      origin: this.baseBox(header, element),
      moved: false,
      parent,
      where: parent,
      grabX: event.clientX - grabbed.left,
      grabY: event.clientY - grabbed.top,
    };

    try {
      this.canvasScroll.setPointerCapture(event.pointerId);
    } catch {
      // A synthesised pointer has no capture to take; the gesture then runs on bubbling alone,
      // which is all the harness's own pointer sequence needs.
    }
  }

  /** The thing follows the pointer, in the picture only - the document is written once, at the
   * drop, so a whole gesture is one undo step rather than one per pixel. */
  private onCanvasDragMove(event: PointerEvent): void {
    if (this.band && event.pointerId === this.band.pointerId) {
      this.trackBand(event);
      return;
    }

    const drag = this.drag;
    if (!drag || event.pointerId !== drag.pointerId) {
      return;
    }

    const dx = event.clientX - drag.startX;
    const dy = event.clientY - drag.startY;
    if (!drag.moved) {
      if (Math.abs(dx) < DRAG_THRESHOLD && Math.abs(dy) < DRAG_THRESHOLD) {
        return;
      }

      drag.moved = true;
      drag.element.classList.add("dc-dragging");
      this.canvasScroll.classList.add("dragging");
      // The gesture's own cursor, worn by the whole canvas while it runs, so passing over
      // another control does not change what the hand is doing.
      this.canvasScroll.style.cursor = drag.edge ? EDGE_CURSOR[drag.edge] ?? "move" : "move";
    }

    // Reparenting happens WHILE the drag runs, not at the drop: the control has to be carried
    // into the container it is heading for, or a box that has left its parent is simply clipped
    // away (every control box hides its overflow) and the developer drags something invisible.
    this.carryInto(drag, event);

    this.paintBox(drag.element, this.proposal(drag, this.toPoints(dx), this.toPoints(dy), event.altKey));
    event.preventDefault();
  }

  /**
   * Moves the dragged control into whatever container the pointer is over, rebasing the gesture
   * on the way so the arithmetic carries on unbroken: the origin becomes the box's position in
   * the NEW container's coordinates and the press moves to here, which is exactly what
   * `proposal` measures from.
   *
   * Only a MOVE reparents. A resize by a handle stays where it is - pulling an edge is a
   * statement about size, not about belonging - and the form has nowhere to go.
   */
  private carryInto(drag: CanvasDrag, event: PointerEvent): void {
    if (drag.edge !== null || drag.name === "") {
      return;
    }

    const over = this.containerAt(event.clientX, event.clientY);
    if (over === null || over.parent.toLowerCase() === drag.parent.toLowerCase()) {
      return;
    }

    const host = this.clientOf(over.parent);
    // A container cannot be carried into itself or into anything it holds: that is a document
    // whose own parent is inside it, which the parser would take and the model never could.
    if (!host || drag.element.contains(host)) {
      return;
    }

    // Placed by the POINTER and the grab offset rather than by where the box had got to: the box
    // is wherever the old container's clamp left it, which for a gesture that crosses on its last
    // move is the edge it was pressed against - the control would land at the boundary instead of
    // under the hand. This is the same arithmetic a hand feels: the control stays where it was
    // picked up, relative to the pointer.
    const room = host.getBoundingClientRect();
    const left = event.clientX - drag.grabX - room.left;
    const top = event.clientY - drag.grabY - room.top;
    drag.element.style.left = `${left}px`;
    drag.element.style.top = `${top}px`;
    host.appendChild(drag.element);

    // The overlay is a SIBLING of the control it dresses, so it travels with it or dresses an
    // empty patch of the container the control just left.
    const overlay = this.canvasScroll.querySelector<HTMLElement>(".dc-selection");
    if (overlay) {
      drag.element.insertAdjacentElement("afterend", overlay);
    }

    drag.parent = over.parent;
    drag.origin = { ...drag.origin, left: this.toPoints(left), top: this.toPoints(top) };
    drag.startX = event.clientX;
    drag.startY = event.clientY;
  }

  /** The element that holds a container's children - a Frame's client, a page's body, the form's
   * own ground for "" - which is where a control carried into it belongs in the DOM. */
  private clientOf(parent: string): HTMLElement | null {
    if (parent === "") {
      return this.canvasScroll.querySelector<HTMLElement>(".dc-form-client");
    }

    const element = this.elementFor(parent);
    if (!element) {
      return null;
    }

    return element.classList.contains("dc-page-body")
      ? element
      : element.querySelector<HTMLElement>(":scope > .dc-frame-client");
  }

  /** Which container an element sits in, by name: the page whose body holds it, the Frame whose
   * client does, or "" for the form's own ground. */
  private hostNameOf(element: HTMLElement): string {
    const host = element.parentElement;
    if (!host) {
      return "";
    }

    if (host.classList.contains("dc-page-body")) {
      return host.dataset.control ?? "";
    }

    return host.classList.contains("dc-frame-client")
      ? host.closest<HTMLElement>(".dc")?.dataset.control ?? ""
      : "";
  }

  /** The drop writes the document: the box the gesture ended on, in points, as one edit. */
  private onCanvasDrop(event: PointerEvent): void {
    if (this.band && event.pointerId === this.band.pointerId) {
      this.dropBand(event);
      return;
    }

    const drag = this.drag;
    if (!drag || event.pointerId !== drag.pointerId) {
      return;
    }

    this.releaseDrag(drag);
    if (!drag.moved) {
      return;
    }

    // A drop the document refuses - a line this cannot rewrite - puts the thing back the way
    // the document still describes it, rather than leaving a picture nothing agrees with.
    const box = this.proposal(
      drag, this.toPoints(event.clientX - drag.startX), this.toPoints(event.clientY - drag.startY),
      event.altKey);

    // A control that ended up somewhere else is REPARENTED: its whole block moves under the new
    // container and its position is read in that container's coordinates. A refusal redraws from
    // the document, because the element has already been carried across the DOM and putting it
    // back by hand would be a second, worse copy of the renderer.
    if (drag.parent.toLowerCase() !== drag.where.toLowerCase()) {
      if (!this.reparentInDocument(drag.name, drag.parent, box)) {
        this.redraw();
      }

      return;
    }

    // A GROUP moves together: the anchor takes the gesture's box and everything else takes the
    // distance the anchor actually travelled, all in one edit and one undo step. A resize is the
    // anchor's alone - eight handles cannot promise anything about the others.
    const moves = this.extras.size > 0 && drag.edge === null
      ? this.groupMoves(drag.name, box, drag.origin)
      : [{ name: drag.name, box, sized: drag.edge !== null }];

    if (!this.writeBoxes(moves)) {
      this.paintBox(drag.element, drag.origin);
    }
  }

  /**
   * Moves a control's whole BLOCK under another container: the header with its new position, its
   * property lines and, for a container, its children, re-indented a level under their new home
   * and appended at the end of it - which is where the model appends, and the markup's order is
   * the model's.
   *
   * One undoable edit, like every other canvas gesture, and it writes only the DOCUMENT: the
   * apply's name-keyed diff calls a changed container a remove-and-add, "which is also the truth
   * of what it means" (the dialect's own rule), so the form catches up at Ctrl+S.
   */
  private reparentInDocument(name: string, parent: string, box: Box): boolean {
    const header = this.headerOf(name);
    const host = this.headerOf(parent);
    if (!header || !host || parent.toLowerCase() === name.toLowerCase()) {
      return false;
    }

    const lines = this.model.getLinesContent();
    const indent = (text: string): number => text.length - text.trimStart().length;
    const own = indent(lines[header.line - 1] ?? "");
    let last = header.line;
    while (last < lines.length && indent(lines[last] ?? "") > own) {
      last += 1;
    }

    // Refused rather than mangled: a container cannot move inside its own block, and the lines
    // that would have to move are the ones it would move into.
    if (host.line >= header.line && host.line <= last) {
      return false;
    }

    // The header, rewritten where it lands. A line that spelled no size still spells none: a
    // move is a move, in a new box or the old one.
    const level = parent === "" ? 0 : indent(lines[host.line - 1] ?? "");
    const shift = (level + 4) - own;
    const block = lines.slice(header.line - 1, last).map((line, at) => at === 0
      ? `${" ".repeat(level + 4)}${header.head.trimStart()}`
        + ` Left="${Math.round(box.left)}" Top="${Math.round(box.top)}"`
        + (header.size !== null
          ? ` Width="${Math.round(box.width)}" Height="${Math.round(box.height)}"`
          : "")
        + (header.close === "/>" ? " />" : ">")
      : shift >= 0 ? " ".repeat(shift) + line : line.slice(Math.min(-shift, indent(line))));

    // The whole document, spliced, as ONE edit: the block leaves and arrives in the same
    // operation, which is one undo step. Two ranges would be fewer characters and would collide
    // the moment the block being moved is the last one in the document - the end of the host's
    // block and the start of the deletion are then the same line.
    const rest = [...lines.slice(0, header.line - 1), ...lines.slice(last)];

    // Where it goes: the end of the new parent's block, which is where the model appends. Found
    // in the text WITHOUT the block, so the arithmetic cannot be thrown by lines that are on
    // their way out.
    const hostLine = parent === ""
      ? 0
      : rest.findIndex((line) => line === lines[host.line - 1]);
    let after = hostLine + 1;
    while (after < rest.length && (parent === "" || indent(rest[after] ?? "") > level)) {
      after += 1;
    }

    this.model.pushStackElement();
    this.model.pushEditOperations([], [{
      range: this.model.getFullModelRange(),
      text: [...rest.slice(0, after), ...block, ...rest.slice(after)].join("\n"),
    }], () => null);

    this.lintNow();
    return true;
  }

  /**
   * What a gesture is proposing, in points: the origin box with the pointer's travel applied -
   * moved whole for a drag, or with the grabbed edges pushed for a resize - then clamped so
   * nothing leaves its parent or collapses to nothing. Measured from the PRESS every frame, so
   * a slow drag cannot accumulate rounding, and rounded to whole points because that is what
   * the document will carry.
   */
  private proposal(drag: CanvasDrag, dx: number, dy: number, free = false): Box {
    const origin = drag.origin;
    const room = this.roomFor(drag.element);
    const floor = drag.name === "" ? MIN_FORM : MIN_CONTROL;

    /*
     * ALIGNMENT COMES FIRST AND THE GRID SECOND, when the two disagree.
     *
     * A neighbour's edge is a place a developer MEANT; a grid line is only a place that was
     * available. So a gesture within reach of an edge takes the edge and draws the guide, and
     * the grid gets what is left over - which is every gesture that is not near anything.
     * Doing it the other way round produced a control that jumped to the grid a point away
     * from the button it was clearly being lined up with.
     */
    // The chrome above this thing's own client, if it has any: a container's grid snap goes
    // to the line it draws rather than to the top of its caption band.
    const chrome = drag.element.querySelector<HTMLElement>(
      ":scope > .dc-frame-rule, :scope > .dc-page-rule");
    const band = chrome ? Number.parseFloat(chrome.style.top || "0") / PT : 0;

    const settle = (box: Box, edge: string | null): Box => {
      const lining = currentSettings().designerSnap === "objects";
      if (free || !lining || drag.name === "") {
        this.showGuides([]);
        return this.snapped(box, edge, free, band);
      }

      const lined = this.aligned(box, drag.name, edge);
      this.showGuides(lined.guides);
      return lined.guides.length > 0
        ? lined.box
        : { left: Math.round(box.left), top: Math.round(box.top),
            width: Math.round(box.width), height: Math.round(box.height) };
    };

    // Snapped BEFORE the clamp, so an edge of the parent beats the grid rather than the other
    // way about: a control pushed into a corner belongs in the corner.
    if (drag.edge === null) {
      return this.clamp(
        settle({ ...origin, left: origin.left + dx, top: origin.top + dy }, null),
        room, floor);
    }

    // The grabbed edges move and the opposite ones stay: west and north change the origin as
    // well as the extent, which is what makes a north-west drag feel like a corner rather than
    // a move. A push past the opposite edge stops at the floor instead of inverting the box.
    const box = { ...origin };
    if (drag.edge.includes("w")) {
      const travel = Math.min(dx, origin.width - floor);
      box.left = origin.left + travel;
      box.width = origin.width - travel;
    } else if (drag.edge.includes("e")) {
      box.width = Math.max(floor, origin.width + dx);
    }

    if (drag.edge.includes("n")) {
      const travel = Math.min(dy, origin.height - floor);
      box.top = origin.top + travel;
      box.height = origin.height - travel;
    } else if (drag.edge.includes("s")) {
      box.height = Math.max(floor, origin.height + dy);
    }

    return this.clamp(settle(box, drag.edge), room, floor);
  }

  /**
   * Draws the alignment guides for the gesture in flight, and clears them when it ends.
   *
   * Inside the moving control's OWN container, because that is the box its coordinates are
   * measured in - a guide drawn on the form for a control inside a Frame would be a line
   * through the right pixels for the wrong reason, and would move when the Frame did. They are
   * two divs, reused, because a gesture repaints them on every pointer move.
   */
  private showGuides(guides: Guide[]): void {
    // Kept after the gesture ends, so the acts can answer what they lined up WITH. The lines
    // themselves come down on release; this is the record of them, and a hand driving the api
    // has no other way to tell an alignment from a coincidence.
    this.lastGuides = guides;

    const host = this.selectedName === null ? null : this.elementOf(this.selectedName)?.parentElement;
    if (!host) {
      return;
    }

    const wanted = new Map(guides.map((guide) => [guide.axis, guide.at]));
    for (const axis of ["x", "y"] as const) {
      const found = host.querySelector<HTMLElement>(`.dc-guide-${axis}`);
      const at = wanted.get(axis);
      if (at === undefined) {
        found?.remove();
        continue;
      }

      const line = found ?? document.createElement("div");
      line.className = `dc-guide dc-guide-${axis}`;
      if (axis === "x") {
        line.style.left = `${at * PT}px`;
      } else {
        line.style.top = `${at * PT}px`;
      }

      if (!found) {
        host.appendChild(line);
      }
    }
  }

  /** What the last gesture lined up with, for the acts to report. */
  private lastGuides: Guide[] = [];

  /** The alignment an act should mention, or nothing when the gesture lined up with nothing. */
  private guideNote(): string {
    return this.lastGuides.length === 0
      ? ""
      : `, lined up with ${this.lastGuides.map((one) => `${one.axis}=${one.at}`).join(" and ")}`;
  }

  /** Takes every guide down, wherever it was drawn: a gesture that ends anywhere - dropped,
   * cancelled, or its control re-rendered underneath it - leaves none behind. */
  private clearGuides(): void {
    for (const line of this.canvasScroll.querySelectorAll(".dc-guide")) {
      line.remove();
    }
  }

  /** Abandons a gesture in flight - Escape, or the pointer taken away - and puts the thing back
   * the way the document still describes it, because nothing was written yet. */
  private cancelDrag(): void {
    const drag = this.drag;
    if (!drag) {
      return;
    }

    this.releaseDrag(drag);
    if (drag.moved) {
      this.paintBox(drag.element, drag.origin);
    }
  }

  private releaseDrag(drag: CanvasDrag): void {
    drag.element.classList.remove("dc-dragging");
    this.canvasScroll.classList.remove("dragging");
    // Every way out of a gesture comes through here - dropped, cancelled, taken away - which
    // is the one place a guide can be guaranteed to be taken down with it.
    this.clearGuides();
    this.canvasScroll.style.cursor = "";
    try {
      this.canvasScroll.releasePointerCapture(drag.pointerId);
    } catch {
      // Never captured, or the pointer is already gone: either way there is nothing to give
      // back, and the drag is over regardless.
    }

    this.drag = null;
  }

  /** The canvas's own keys: arrows nudge the selection a point at a time and Shift+arrow
   * resizes it - through the same commit a drag takes, so the picture, the dot and undo behave
   * identically - Delete takes the selection out of the document, Escape abandons a gesture in
   * flight, and undo/redo reach the document from this half too, because this half edits it. */
  private onCanvasKey(event: KeyboardEvent): void {
    if (event.key === "Escape") {
      this.cancelDrag();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && !event.altKey) {
      const key = event.key.toLowerCase();
      if (key === "z" || key === "y") {
        // The MODEL's own undo rather than `editor.trigger("undo")`. The undo stack belongs to
        // the document both halves share, and this half has no editor of its own to route
        // through: going straight at the model says what it means and does not depend on
        // monaco's idea of which editor is active.
        if (key === "y" || event.shiftKey) {
          this.model.redo();
        } else {
          this.model.undo();
        }

        event.preventDefault();
      }

      // The clipboard four, on the keys every designer uses for them. Ctrl+D duplicates without
      // touching what Ctrl+C last took, which is the whole reason it is a separate key.
      if (key === "c" || key === "x" || key === "v" || key === "d") {
        const said = key === "c"
          ? this.copySelection()
          : key === "x"
            ? this.cutSelection()
            : key === "v" ? this.pasteClipboard() : this.duplicateSelection();

        // A refusal is said on the status line rather than swallowed: Ctrl+V with an empty
        // clipboard doing nothing at all is indistinguishable from Ctrl+V being unwired.
        if (!/^(copied|cut|pasted)/.test(said)) {
          this.deps.notify(said);
        }

        event.preventDefault();
      }

      return;
    }

    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      this.deleteSelection();
      return;
    }

    const step = NUDGE[event.key];
    if (!step || event.altKey || this.selectedName === null) {
      return;
    }

    // The FORM has no position of its own - it is where the canvas puts it - so a bare arrow
    // has nothing to do there, while a Shift+arrow resizes it like anything else.
    if (this.selectedName === "" && !event.shiftKey) {
      return;
    }

    // Without this the canvas scrolls under the thing the developer is placing.
    event.preventDefault();

    // SHIFT resizes where a bare arrow moves, the native designer's own pairing: the same two
    // deltas, applied to the extent rather than to the origin.
    const header = this.headerOf(this.selectedName);
    const element = this.elementOf(this.selectedName);
    if (!header || !element) {
      return;
    }

    const base = this.baseBox(header, element);
    const room = this.roomFor(element);
    const floor = this.selectedName === "" ? MIN_FORM : MIN_CONTROL;
    const box = this.clamp(event.shiftKey
      ? { ...base, width: base.width + step.dx, height: base.height + step.dy }
      : { ...base, left: base.left + step.dx, top: base.top + step.dy }, room, floor);

    // A nudge carries the whole group, like a drag; a Shift+arrow resizes the anchor alone, for
    // the same reason its handles do.
    this.writeBoxes(this.extras.size > 0 && !event.shiftKey
      ? this.groupMoves(this.selectedName, box, base)
      : [{ name: this.selectedName, box, sized: event.shiftKey }]);
  }

  /**
   * Writes boxes into the document: each thing's own line rewritten with the position and size it
   * now has, rounded to whole points, as ONE undoable edit, and painted at once so the gesture
   * lands before the round trip answers.
   *
   * This is the single commit every canvas gesture takes - drag, resize, nudge, align, and the
   * harness's acts - and it writes the DOCUMENT, never the form: the draft preview shows it
   * immediately, the dot says it has not reached the form yet, and Ctrl+S is still the only
   * apply. Whole points because a hand-placed control wants round numbers. The FORM's line takes
   * a size and never a position.
   *
   * Several at once is what makes a GROUP gesture one undo step rather than one per control. Each
   * is a different line, so the ranges cannot overlap and monaco applies them in one operation.
   * All or nothing: a line this cannot rewrite - half-typed, a caption whose quote never closes -
   * refuses the whole gesture rather than moving the controls it happened to understand.
   */
  private writeBoxes(moves: { name: string; box: Box; sized: boolean }[]): boolean {
    const edits: monaco.editor.IIdentifiedSingleEditOperation[] = [];
    const painted: { element: HTMLElement; box: Box }[] = [];

    for (const move of moves) {
      const element = this.elementOf(move.name);
      const header = this.headerOf(move.name);
      if (!element || !header) {
        return false;
      }

      const box = {
        left: Math.round(move.box.left),
        top: Math.round(move.box.top),
        width: Math.round(move.box.width),
        height: Math.round(move.box.height),
      };

      // A line that spelled no size keeps spelling none unless this gesture is a resize: a move
      // must not quietly pin a size the developer left to the control.
      const spellSize = move.sized || header.size !== null;
      edits.push({
        range: new monaco.Range(header.line, 1, header.line, this.model.getLineMaxColumn(header.line)),
        text: header.head
          + (header.form ? "" : ` Left="${box.left}" Top="${box.top}"`)
          + (spellSize ? ` Width="${box.width}" Height="${box.height}"` : "")
          + (header.close === "/>" ? " />" : ">"),
      });
      painted.push({ element, box });
    }

    if (edits.length === 0) {
      return false;
    }

    // A stack stop BEFORE the edit and none after. Monaco appends edits to whichever element is
    // open, so the stop is what keeps one gesture from joining the gesture before it; leaving
    // the element open afterwards is equally deliberate, because the canonical print that
    // follows a save belongs to this gesture - with it riding along, one Ctrl+Z after a save
    // reaches the text from before the move rather than a step that only differs by the
    // machine's own rounding (both halves measured live, 2026-08-15).
    this.model.pushStackElement();
    this.model.pushEditOperations([], edits, () => null);
    for (const { element, box } of painted) {
      this.paintBox(element, box);
    }

    this.lintNow();
    return true;
  }

  /**
   * The same delta applied to everything else in the selection, each clamped inside its own
   * container - the group move. The anchor's own box is the gesture's; the rest follow it by the
   * distance it actually travelled, which is what keeps a snapped group in formation.
   */
  private groupMoves(anchor: string, box: Box, from: Box): { name: string; box: Box; sized: boolean }[] {
    const dLeft = box.left - from.left;
    const dTop = box.top - from.top;
    const moves = [{ name: anchor, box, sized: false }];

    for (const name of this.extras) {
      const header = this.headerOf(name);
      const element = this.elementFor(name);
      if (!header || !element) {
        continue;
      }

      const base = this.baseBox(header, element);
      moves.push({
        name,
        box: this.clamp(
          { ...base, left: base.left + dLeft, top: base.top + dTop },
          this.roomFor(element), MIN_CONTROL),
        sized: false,
      });
    }

    return moves;
  }

  /**
   * The drop: a new control's line into the document, under the container the pointer was over,
   * at the point it was let go - clamped inside that container, on whole points, named the way
   * the native toolbox names one (the kind plus the first free number).
   *
   * One undoable edit, like every other canvas gesture, and it writes the DOCUMENT: the draft
   * preview draws the control at once, the dot says it is not on the form yet, and Ctrl+S adds
   * it there through the apply's name-keyed diff - which is the SAME add the `designer` route
   * makes, so a control born on the canvas is indistinguishable from one born through the api.
   */
  private dropTool(x: number, y: number, free = false): boolean {
    const carrying = this.carrying;
    this.carrying = null;
    carrying?.ghost.remove();
    if (!carrying) {
      return false;
    }

    const where = this.containerAt(x, y);
    if (!where) {
      return false;
    }

    const host = where.parent === "" ? this.headerOf("") : this.headerOf(where.parent);
    if (!host) {
      return false;
    }

    // Inside the container, and never off its edges: the same floor every gesture keeps - and
    // onto the grid on the way, because a control dropped from the palette is placed by the
    // pointer like any other.
    // A container dropped from the palette gets the same courtesy a dragged one does: the grid
    // takes the line it will DRAW, not the top of the caption band above it. There is nothing
    // rendered yet to measure, so the band is the one the stylesheet will give it.
    const room = this.roomOfContainer(where.parent);
    const dropBand = carrying.kind === "Frame" || carrying.kind === "MultiPage"
      || carrying.kind === "TabStrip"
      ? (carrying.kind === "Frame" ? FRAME_RULE_TOP : PAGE_STRIP_HEIGHT) / PT
      : 0;
    const box = this.clamp(
      this.snapped(
        { left: where.left, top: where.top, width: carrying.width, height: carrying.height },
        null, free, dropBand),
      room, MIN_CONTROL);
    const name = this.freeName(carrying.kind);
    const lines = this.model.getLinesContent();
    const indent = (text: string): number => text.length - text.trimStart().length;
    const level = where.parent === "" ? 0 : indent(lines[host.line - 1] ?? "");

    // At the END of the container's block, which is where a control added last belongs: the
    // markup's order is the model's order, and the model appends.
    /*
     * STOPPING AT THE FORM'S CLOSE, which a tagged document has and a line-oriented one did not.
     *
     * Dropping on the form walked to the LAST line of the document, because under the old dialect
     * the form's block ran to the end of the file. It does not any more: `</Form>` is the last
     * line, so the new element landed AFTER it and the apply refused the whole document - "one
     * Form per document; nothing follows </Form>" - which took the drop, the save and every row
     * after them down with it (2026-08-17). A container's own walk is unchanged; it already stops
     * on its closing tag, which sits at its own indent.
     */
    let after = host.line;
    while (after < lines.length
      && !/^\s*<\/Form>/.test(lines[after] ?? "")
      && (where.parent === "" || indent(lines[after] ?? "") > level)) {
      after += 1;
    }

    const line = `${" ".repeat(level + 4)}<${carrying.kind} Name="${name}"`
      + ` Left="${Math.round(box.left)}" Top="${Math.round(box.top)}"`
      + ` Width="${Math.round(box.width)}" Height="${Math.round(box.height)}" />`;

    /*
     * AN EMPTY CONTAINER HAS TO BE OPENED FIRST, because a self-closing element has no body to
     * put anything in. A childless Page prints as `<Page Name="Page2" ... />`, so the walk above
     * finds nothing deeper than it and stops on the page's own line - and the control went in as
     * the page's SIBLING, a control directly inside the MultiPage. The lint says exactly that ("a
     * MultiPage holds Pages; controls go on a Page"), the strict parse then refuses the document,
     * and the canvas quietly stops previewing because a draft that will not parse keeps the last
     * picture (2026-08-17). So the tag is split into a pair and the child goes between them.
     */
    const hostText = this.model.getLineContent(host.line);
    const hostEmpty = where.parent !== "" && after === host.line && /\/>\s*$/.test(hostText);
    this.model.pushStackElement();
    if (hostEmpty) {
      const opened = hostText.replace(/\s*\/>\s*$/, ">");
      this.model.pushEditOperations([], [{
        range: new monaco.Range(host.line, 1, host.line, this.model.getLineMaxColumn(host.line)),
        text: `${opened}\n${line}\n${" ".repeat(level)}</${this.kindOf(where.parent) || "Page"}>`,
      }], () => null);
    } else {
      this.model.pushEditOperations([], [{
        range: new monaco.Range(after, this.model.getLineMaxColumn(after), after, this.model.getLineMaxColumn(after)),
        text: `\n${line}`,
      }], () => null);
    }

    // Selected on arrival, the way a control dropped from the native toolbox is: the panel
    // targets it and the next gesture is about the thing just made.
    this.select(name);
    this.lintNow();
    return true;
  }

  /** The inside of a container, in points, for a drop that has no element to measure yet. */
  /**
   * TWO CASES, because a container's client is either INSIDE it or IS it.
   *
   * A Frame draws a `.dc-frame-client` within its own rectangle. A PAGE is its own client - the
   * body element carries the page's name and kind, which is what lets a press on it select the
   * page and a drop land in it - so there is nothing inside it to look for.
   *
   * This was three branches until 2026-08-16, the third a `closest(".dc-page-body")` off a
   * `data-control` lookup. It was not redundant, which is the interesting part: since the body
   * started carrying its own identity, the second branch found the body and then searched INSIDE
   * it for a client, found nothing, and the third branch quietly did every page's work. Two
   * cases that say which is which beat three where the load-bearing one looks like a fallback.
   */
  private roomOfContainer(parent: string): { width: number; height: number } | null {
    const element = parent === "" ? null : this.elementFor(parent);
    const client = parent === ""
      ? this.canvasScroll.querySelector<HTMLElement>(".dc-form-client")
      : element?.classList.contains("dc-page-body")
        ? element
        : element?.querySelector<HTMLElement>(".dc-frame-client");
    return client ? { width: client.clientWidth / PT, height: client.clientHeight / PT } : null;
  }

  /** The name the native toolbox would give: the kind, then the first number the document is
   * not already using - counted across the WHOLE document, because a form's names are one
   * namespace however deeply a control is nested. */
  /** The same, for a gesture that has to ask the host before it knows. */
  private async sayLater(said: Promise<string>): Promise<void> {
    this.say(await said);
  }

  /** A gesture's answer, on the status line when it is a refusal and nowhere when it worked -
   * the canvas itself is the report for a paste that landed or a control that resized. */
  private say(said: string): void {
    if (!/^(copied|cut|pasted|centreX|centreY|fit|grid)/.test(said)) {
      this.deps.notify(said);
    }
  }

  /** Every name the document currently spells, lowercased, for allocating ones that are free. */
  private takenNames(): Set<string> {
    return new Set(this.model.getLinesContent()
      .map((text) => /\bName="([^"]*)"/.exec(text)?.[1]?.toLowerCase())
      .filter((name): name is string => name !== undefined));
  }

  private freeName(kind: string, taken: Set<string> = this.takenNames()): string {
    // THE NAMES THE DOCUMENT SPELLS, read off the attribute that carries them. This matched the
    // second word of a line - `Kind Name` in the old dialect - and under tags that word is
    // `Name="Page1"` itself, so nothing ever counted as taken and every new control came back as
    // `Kind1`. New Page on a MultiPage that already had Page1 and Page2 produced a second Page1
    // (2026-08-17), and the apply then had two elements of one name.
    for (let at = 1; ; at++) {
      const name = `${kind}${at}`;
      if (!taken.has(name.toLowerCase())) {
        // Claimed as it is handed out, so a caller allocating SEVERAL against one set - a paste
        // renaming a container and all of its children - does not get one name twice.
        taken.add(name.toLowerCase());
        return name;
      }
    }
  }

  /**
   * COPY, PASTE, CUT AND DUPLICATE, on the document rather than on the model.
   *
   * Laying out a form is making one OptionButton and then making five more like it, and until
   * 2026-08-18 that was five trips to the toolbox: there was no clipboard on this canvas at all
   * (the owner, having been shown the gap: "please fix these gaps").
   *
   * It rides the transaction log like every other gesture - a paste is new element lines, one
   * undoable edit, reaching the form at Ctrl+S - so nothing touches MSForms until the apply and
   * a mistaken paste is one Ctrl+Z away.
   *
   * The clipboard is THIS PRODUCT'S, not the system's. A system clipboard would carry the markup
   * between forms and even between machines, which sounds better until a paste of some text a
   * developer copied out of a mail lands as controls. The document is already the thing that can
   * be copied as text, by selecting it in the markup half and using the editor's own clipboard.
   */
  private clipboard: string[] | null = null;

  /**
   * How far a pasted copy lands from its original, in POINTS.
   *
   * OURS, not measured off the native designer: its own offset cannot be read without driving a
   * designer window this product keeps shut. What it has to achieve is that a copy is visibly a
   * copy - a paste that lands exactly on the original reads as nothing having happened, and the
   * developer drags the wrong one.
   */
  private static readonly PASTE_OFFSET = 6;

  /**
   * An element header's Name, in the POSITION THE PRINTER PUTS IT: first attribute after the
   * tag. The paste rename used to take the leftmost `Name="..."` anywhere in the line, and a
   * hand-written caption containing that text - `Caption="Call me Name="` is legal markup -
   * put the replacement inside the caption and mangled the line (the 2026-08-19 hunt, round
   * three). Anchored, a hand-reordered line simply keeps its name, and the duplicate-name
   * refusal that follows names the problem instead of shredding the text.
   */
  private static readonly HEADER_NAME = /^(\s*<[A-Za-z][\w.]*\s+)Name="([^"]*)"/;

  /** The selection's blocks, as the lines they are written on. Answers null when the selection
   * holds nothing that has a block - the form itself, above all, which cannot be copied. */
  private selectionBlocks(): string[] | null {
    const names = this.selection().filter((one) => one !== "");
    if (names.length === 0) {
      return null;
    }

    const lines = this.model.getLinesContent();
    const blocks: string[] = [];
    for (const name of names) {
      const span = this.blockOf(name);
      if (!span) {
        return null;
      }

      blocks.push(lines.slice(span.from - 1, span.to).join("\n"));
    }

    return blocks;
  }

  /** Copies the selection. The document is untouched, so this is the one gesture here that is
   * not an edit. */
  copySelection(): string {
    const blocks = this.selectionBlocks();
    if (!blocks) {
      return "nothing is selected that can be copied";
    }

    this.clipboard = blocks;
    const names = this.selection().filter((one) => one !== "");
    return `copied ${names.join(", ")}`;
  }

  /** Copy, then delete: one clipboard load and one undoable edit. */
  cutSelection(): string {
    const copied = this.copySelection();
    if (!copied.startsWith("copied")) {
      return copied;
    }

    return this.deleteSelection() ? copied.replace("copied", "cut") : "the document would not take the removal";
  }

  /** Copy and paste in one gesture, without disturbing what the clipboard already holds - which
   * is what Ctrl+D means everywhere else and what makes it a different key from Ctrl+C Ctrl+V. */
  duplicateSelection(): string {
    const blocks = this.selectionBlocks();
    if (!blocks) {
      return "nothing is selected that can be duplicated";
    }

    return this.pasteBlocks(blocks);
  }

  /** Pastes whatever the last Copy or Cut took. */
  pasteClipboard(): string {
    return this.clipboard === null
      ? "nothing has been copied yet"
      : this.pasteBlocks(this.clipboard);
  }

  /**
   * The write behind Paste and Duplicate both.
   *
   * WHERE IT LANDS is the container the selection sits in, which is how the toolbox drop resolves
   * its host too: with a control selected the copy joins it as a sibling, and with a CONTAINER
   * selected the copy goes inside. Nothing selected means the form.
   *
   * EVERY name in a block is renamed, not just the top one, because a Frame carries its children
   * and two controls of one name is a document the apply refuses. They are allocated against a
   * single growing set so a container and its three children cannot be handed the same name.
   *
   * ONLY the top-level element is offset. A child's Left and Top are its parent's coordinates, so
   * moving those as well would shift the children twice and take them out of the copy.
   */
  private pasteBlocks(blocks: string[]): string {
    const anchor = this.selectedName;

    /*
     * COPYING A CONTAINER LEAVES IT SELECTED, so "paste into the selected container" put the copy
     * INSIDE the original - select a Frame, Ctrl+C, Ctrl+V, and the Frame grows a Frame. Caught
     * by driving it (2026-08-18) rather than by reading the rule, which sounded right.
     *
     * So the container rule holds for pasting something ELSE into a container, and a paste of the
     * selected thing itself lands beside it. Which is what the gesture means: Ctrl+C then Ctrl+V
     * asks for another one of these, not for one of these inside this one.
     */
    const copyingItself = anchor !== null && anchor !== ""
      && blocks.some((block) =>
        (DesignerView.HEADER_NAME.exec(block)?.[2] ?? "").toLowerCase() === anchor.toLowerCase());

    const host = anchor === null || anchor === ""
      ? ""
      : copyingItself
        ? this.parentOf(anchor)
        : this.kindOf(anchor) === "Page" || this.isContainer(anchor)
          ? anchor
          : this.parentOf(anchor);

    // A Page belongs to a MultiPage and a Tab to a TabStrip, and nowhere else. The parser would
    // refuse the document and the canvas would stop previewing, so the refusal is said here where
    // it can name the reason.
    const hostKind = host === "" ? "Form" : this.kindOf(host);
    for (const block of blocks) {
      const kind = /^\s*<([A-Za-z][\w.]*)/.exec(block)?.[1] ?? "";
      const wants = kind === "Page" ? "MultiPage" : kind === "Tab" ? "TabStrip" : null;
      if (wants !== null && hostKind !== wants) {
        return `a ${kind} can only be pasted into a ${wants}`;
      }

      if (wants === null && (hostKind === "MultiPage" || hostKind === "TabStrip")) {
        return `a ${hostKind} holds only ${hostKind === "MultiPage" ? "Pages" : "Tabs"}`;
      }
    }

    const lines = this.model.getLinesContent();
    const indent = (text: string): number => text.length - text.trimStart().length;

    // The insertion point: the end of the host's body, which for the form is the line before
    // `</Form>` and for a container the last line indented under it.
    let after: number;
    let level: number;
    if (host === "") {
      const closing = lines.findIndex((text) => /^\s*<\/Form\s*>/.test(text));
      after = closing < 0 ? lines.length : closing;
      level = 4;
    } else {
      const header = this.headerOf(host);
      if (!header) {
        return `${host} is not a line this can paste into`;
      }

      level = indent(lines[header.line - 1] ?? "");
      after = header.line;
      while (after < lines.length && indent(lines[after] ?? "") > level) {
        after += 1;
      }

      // A container written self-closing has no body to paste into; it has to be opened first,
      // the same courtesy the toolbox drop gives an empty page.
      if (/\/>\s*$/.test(lines[header.line - 1] ?? "")) {
        return `${host} is empty and self-closed; drop a control into it first`;
      }

      level += 4;
    }

    const taken = this.takenNames();
    const landed: string[] = [];
    const written: string[] = [];

    for (const block of blocks) {
      const rows = block.split("\n");
      const base = indent(rows[0] ?? "");
      const renamed = rows.map((row, at) => {
        let text = " ".repeat(Math.max(0, indent(row) - base) + level) + row.trimStart();

        text = text.replace(DesignerView.HEADER_NAME, (_whole, lead: string) => {
          const kind = /^\s*<([A-Za-z][\w.]*)/.exec(row)?.[1] ?? "Control";
          const now = this.freeName(kind, taken);
          if (at === 0) {
            landed.push(now);
          }

          return `${lead}Name="${now}"`;
        });

        if (at === 0) {
          for (const axis of ["Left", "Top"] as const) {
            text = text.replace(new RegExp(`\\b${axis}="(-?[\\d.]+)"`), (whole, value: string) =>
              whole.replace(`"${value}"`, `"${Number(value) + DesignerView.PASTE_OFFSET}"`));
          }
        }

        return text;
      });

      written.push(renamed.join("\n"));
    }

    const column = this.model.getLineMaxColumn(after);
    this.model.pushStackElement();
    this.model.pushEditOperations([], [{
      range: new monaco.Range(after, column, after, column),
      text: `\n${written.join("\n")}`,
    }], () => null);

    // The copies are what is selected afterwards, which is what makes a paste-then-drag one
    // motion rather than two, and what every other designer does.
    for (const [at, name] of landed.entries()) {
      this.select(name, at > 0);
    }

    this.lintNow();
    return `pasted ${landed.join(", ")}`;
  }

  /**
   * Takes the selected control out of the document - its line and everything indented under it,
   * which is its properties and, for a container, its children. One undoable edit, like every
   * other canvas gesture, so a mistaken Delete is one Ctrl+Z away and the form itself is
   * untouched until the save carries the removal through the apply's name-keyed diff.
   *
   * The FORM cannot be deleted from its own canvas: removing a component is the tree's gesture,
   * with the confirmation the product asks there. Selection lands back on the form, which is
   * where the native designer leaves it and what returns the Properties panel to the component.
   */
  private deleteSelection(): boolean {
    const names = this.selection().filter((one) => one !== "");
    if (names.length === 0) {
      return false;
    }

    return names.length === 1
      ? this.deleteFromDocument(names[0] ?? "")
      : this.deleteGroup(names);
  }

  /**
   * A whole selection out of the document at once - every block, one edit, one undo step.
   *
   * Spliced out of the line array rather than removed range by range: the ranges of several
   * blocks are easy to compute and awkward to apply, because each removal moves the ones after
   * it, and a group can hold a container whose children are blocks of their own.
   */
  private deleteGroup(names: string[]): boolean {
    const lines = this.model.getLinesContent();
    const indent = (text: string): number => text.length - text.trimStart().length;
    const doomed = new Set<number>();

    for (const name of names) {
      const header = this.headerOf(name);
      if (!header) {
        return false;
      }

      const own = indent(lines[header.line - 1] ?? "");
      doomed.add(header.line);
      let at = header.line;
      for (; at < lines.length && indent(lines[at] ?? "") > own; at++) {
        doomed.add(at + 1);
      }

      // AND ITS CLOSING TAG, which the indent walk cannot see: `</Page>` sits at the SAME indent as
      // the `<Page>` that opened it, so a walk that takes everything DEEPER stops one line short and
      // leaves the close behind. An orphaned close is a document the parser refuses, and a document
      // the parser refuses stops the canvas previewing at all - deleting a container's pages looked
      // like the strip simply not updating (2026-08-17).
      if (at < lines.length && /^\s*<\//.test(lines[at] ?? "") && indent(lines[at] ?? "") === own) {
        doomed.add(at + 1);
      }
    }

    this.model.pushStackElement();
    this.model.pushEditOperations([], [{
      range: this.model.getFullModelRange(),
      text: lines.filter((_, at) => !doomed.has(at + 1)).join("\n"),
    }], () => null);

    this.select("");
    this.lintNow();
    return true;
  }

  /** The commit behind Delete, by name, so the tab strip's Delete Page and the Delete key are
   * one edit rather than two spellings of it. */
  private deleteFromDocument(name: string): boolean {
    if (!name) {
      return false;
    }

    const header = this.headerOf(name);
    if (!header) {
      return false;
    }

    const lines = this.model.getLinesContent();
    const indent = (text: string): number => text.length - text.trimStart().length;
    const own = indent(lines[header.line - 1] ?? "");
    let last = header.line;
    while (last < lines.length && indent(lines[last] ?? "") > own) {
      last += 1;
    }

    // AND ITS CLOSING TAG, which the indent walk cannot see: `</Page>` sits at the SAME indent as
    // the `<Page>` that opened it, so a walk that takes everything DEEPER stops one line short and
    // leaves the close behind. An orphaned close is a document the parser refuses, and a document
    // the parser refuses stops the canvas previewing at all - deleting a container's pages looked
    // like the strip simply not updating (2026-08-17).
    if (last < lines.length && /^\s*<\//.test(lines[last] ?? "")
      && indent(lines[last] ?? "") === own) {
      last += 1;
    }

    // Eat the line break BEFORE the block when it runs to the end of the document, and the one
    // after it otherwise: either way the document is left without a blank line where the
    // control used to be, which is a line the parser would refuse.
    const range = last < this.model.getLineCount()
      ? new monaco.Range(header.line, 1, last + 1, 1)
      : new monaco.Range(header.line - 1, this.model.getLineMaxColumn(header.line - 1),
        last, this.model.getLineMaxColumn(last));

    this.model.pushStackElement();
    this.model.pushEditOperations([], [{ range, text: "" }], () => null);
    this.select("");
    this.lintNow();
    return true;
  }

  /**
   * Asks for the lint NOW, cancelling the typing debounce.
   *
   * A GESTURE is finished the moment it happens - a drop, a delete, a drag let go - so waiting
   * out a delay built for keystrokes is a delay for nothing. Measured before this existed: a
   * toolbox drop took 347ms to appear on the canvas and a delete 348ms, which is the debounce
   * almost exactly, and which the owner felt immediately (2026-08-15: "some delay when dragging
   * an element onto the UI, or deleting"). A drag and a resize paint themselves as they go, so
   * they never showed it; the two gestures that change the SHAPE of the document did.
   */
  private lintNow(): void {
    clearTimeout(this.lintTimer);
    this.deps.lint(this.model.getValue());
  }

  /** Where a gesture starts from: the DOCUMENT's own numbers when the line spells them, and the
   * picture's only where it does not, so the arithmetic stays right even when the canvas is a
   * parse behind. */
  private baseBox(header: HeaderLine, element: HTMLElement): Box {
    const painted = this.inPoints({
      left: Number.parseFloat(element.style.left || "0"),
      top: Number.parseFloat(element.style.top || "0"),
      width: element.offsetWidth,
      height: element.offsetHeight,
    });
    return {
      left: header.at?.left ?? painted.left,
      top: header.at?.top ?? painted.top,
      width: header.size?.width ?? painted.width,
      height: header.size?.height ?? painted.height,
    };
  }

  /** The floor a gesture works on, in points: the inside of the box that holds this thing, or
   * nothing at all for the FORM, which is held by the canvas and may grow as it likes. */
  private roomFor(element: HTMLElement): { width: number; height: number } | null {
    const host = element.classList.contains("dc-form") ? null : element.parentElement;
    return host ? { width: host.clientWidth / PT, height: host.clientHeight / PT } : null;
  }

  /** Keeps a gesture inside the box that holds it and above the floor size: a gesture may
   * reposition or resize a control, never lose it behind an edge or collapse it to nothing.
   * Reparenting - dragging a control out of its Frame - is a later gesture, so the parent's
   * client area is the whole floor here. A control already bigger than its parent keeps a zero
   * origin rather than being yanked to a negative one. */
  private clamp(box: Box, room: { width: number; height: number } | null, floor: number): Box {
    const width = Math.max(floor, room ? Math.min(box.width, room.width) : box.width);
    const height = Math.max(floor, room ? Math.min(box.height, room.height) : box.height);
    const maxLeft = room ? Math.max(0, room.width - width) : Number.POSITIVE_INFINITY;
    const maxTop = room ? Math.max(0, room.height - height) : Number.POSITIVE_INFINITY;
    return {
      left: Math.min(Math.max(0, box.left), maxLeft),
      top: Math.min(Math.max(0, box.top), maxTop),
      width,
      height,
    };
  }

  private inPoints(box: Box): Box {
    return { left: box.left / PT, top: box.top / PT, width: box.width / PT, height: box.height / PT };
  }

  /**
   * Draws the grid on the form's ground, and puts the switch in the state the setting is in.
   *
   * On the FORM rather than on the canvas behind it, because the grid a developer is placing
   * against starts at the form's own origin - a grid drawn on the scroll box would drift with
   * every scroll and line up with nothing. Two one-pixel dots per cell, painted as a repeating
   * background so the DOM carries no grid elements at all: a 360x320 form at a six-point grid
   * is 3,180 cells, and that many divs is a canvas that stutters when it moves.
   */
  private showGrid(): void {
    const settings = currentSettings();
    const on = settings.designerSnap === "grid";
    const lining = settings.designerSnap === "objects";

    const dress = (control: HTMLButtonElement, pressed: boolean, title: string): void => {
      control.setAttribute("aria-pressed", pressed ? "true" : "false");
      control.classList.toggle("on", pressed);
      control.title = title;
      control.setAttribute("aria-label", title);
    };

    dress(this.snapToggle, on, on
      ? `Snapping to a ${settings.designerGridSize}-point grid. Alt overrides it; arrow keys `
        + "always move by one point."
      : "Snap to grid: place controls on an even grid.");
    dress(this.alignToggle, lining, lining
      ? "Lining up with the other controls - their edges and centers, with a guide where it "
        + "lands. Alt overrides it."
      : "Snap to the other controls: line up with a neighbor's edge or center.");

    const client = this.canvasScroll.querySelector<HTMLElement>(".dc-form-client")
      ?? this.canvasScroll.querySelector<HTMLElement>(".dc-form");
    if (!client) {
      return;
    }

    // A sixth of the form's own ink, which is a dot you can place against and stop seeing while
    // you work. At full strength it read as texture on the form rather than as a grid under it
    // (the owner, 2026-08-16: "can you make the dots more subtle?").
    //
    // AND SHIFTED BACK BY HALF A CELL, because a radial gradient is centred in its tile: drawn
    // at the origin the dots sit half a step from every coordinate this snaps to, so a control
    // that landed exactly on the grid appeared to land exactly between it ("should controls
    // align exactly on the dots in snap to grid mode?" - yes, and now they do). The shift puts a
    // dot on 0,0 and on every multiple after it, measured from the client's own corner, which
    // is where a control's Left and Top are measured from too.
    const step = Math.max(2, settings.designerGridSize) * PT;
    const half = -step / 2;

    // TWO LAYERS, composed rather than one overwriting the other: the grid on top, the form's
    // own picture under it. They share an element, and until 2026-08-16 there was only ever one
    // of them to draw - a form with a picture would have lost it to the next grid toggle, or
    // painted over the grid, depending on which ran last.
    const layers: PictureLayer[] = [];
    if (on) {
      layers.push({
        // A DOT BIG ENOUGH TO RASTERISE. This asked for a 0.5px radius with both colour stops at
        // the same offset - a half-pixel disc with no ramp - and Chromium painted NOTHING at all:
        // measured off the live canvas 2026-08-17, the form's ground a uniform 240,240,240 with
        // the toggle lit and the gradient present in the computed style (the owner: "grid snap
        // indicator is lit, but grid snap not showing on form"). A whole pixel of solid centre and
        // a short ramp out to 1.4px is the smallest thing that reliably lands on a device pixel,
        // and it keeps the dot subtle the way the owner asked for on 2026-08-16 - subtle was never
        // the problem, invisible was.
        image: "radial-gradient(color-mix(in srgb, currentColor 18%, transparent) 0 1px, transparent 1.4px)",
        size: `${step}px ${step}px`,
        position: `${half}px ${half}px`,
        repeat: "repeat",
      });
    }

    if (this.formGround) {
      layers.push(this.formGround);
    }

    client.style.backgroundImage = layers.map((layer) => layer.image).join(", ");
    client.style.backgroundSize = layers.map((layer) => layer.size).join(", ");
    client.style.backgroundPosition = layers.map((layer) => layer.position).join(", ");
    client.style.backgroundRepeat = layers.map((layer) => layer.repeat).join(", ");
  }

  /**
   * A number on the grid, in points, when the grid is on - and a whole point when it is not.
   *
   * ONLY POINTER GESTURES COME THROUGH HERE. An arrow key moves one point whatever the setting
   * says, because the hand that reaches for the keyboard has already decided the grid is not
   * where this control belongs, and a nudge that jumps six points is not a nudge. That is the
   * editor's own division of labour and the reason its Align Controls to Grid never fought
   * anyone: the mouse is the fast tool and the keyboard is the exact one.
   */
  private onGrid(value: number): number {
    const settings = currentSettings();
    if (settings.designerSnap !== "grid") {
      return Math.round(value);
    }

    const step = Math.max(2, settings.designerGridSize);
    return Math.round(value / step) * step;
  }

  /**
   * The edges and centres of everything the gesture could line up WITH: the siblings sharing
   * this control's container, plus the container's own inside edges and middle.
   *
   * Siblings only, because lining up with a control in a different box means nothing on screen -
   * a button inside a Frame and a button beside it are measured from different origins, and a
   * guide drawn between them would be pointing at a coincidence.
   */
  private neighbours(name: string): { x: number[]; y: number[] } {
    const element = this.elementOf(name);
    const host = element?.parentElement;
    if (!element || !host) {
      return { x: [], y: [] };
    }

    const x: number[] = [];
    const y: number[] = [];

    for (const other of host.children) {
      if (!(other instanceof HTMLElement) || other === element
        || !other.classList.contains("dc") || !other.dataset.control) {
        continue;
      }

      // Rounded to a hundredth of a point, because these come back through the browser's own
      // pixel arithmetic: a control the document puts at 85 measures 84.99975, and a guide
      // reported at that is a true number nobody can read.
      const round = (value: number): number => Math.round(value * 100) / 100;
      const left = round(Number.parseFloat(other.style.left || "0") / PT);
      const top = round(Number.parseFloat(other.style.top || "0") / PT);
      const width = round(other.offsetWidth / PT);
      const height = round(other.offsetHeight / PT);

      /*
       * A CONTAINER OFFERS THE EDGE IT PAINTS, not the one its rectangle has.
       *
       * A Frame's rectangle starts at the top of its caption band and its RULE is about four
       * points lower, so a guide at the rectangle's top runs through the caption and the
       * control lining up with it looks aligned to the lettering (the owner, 2026-08-16: "the
       * button should snap to the frame's edge, not the label"). A developer means the line
       * they can see. The same is true of a MultiPage, whose rectangle starts at the tabs.
       */
      const rule = other.querySelector<HTMLElement>(":scope > .dc-frame-rule, :scope > .dc-page-rule");
      const band = rule ? round(Number.parseFloat(rule.style.top || "0") / PT) : 0;

      x.push(left, round(left + width / 2), left + width);
      y.push(round(top + band), round(top + band + (height - band) / 2), top + height);
    }

    // The container itself: its inside edges, and its middle for centring against.
    const room = this.roomFor(element);
    if (room) {
      const round = (value: number): number => Math.round(value * 100) / 100;
      x.push(0, round(room.width / 2), round(room.width));
      y.push(0, round(room.height / 2), round(room.height));
    }

    return { x, y };
  }

  /**
   * A gesture's box, aligned to the nearest neighbour edge within reach - and the guides to
   * draw for it.
   *
   * The candidates on each axis are the moving box's own leading edge, centre and trailing
   * edge, so a control lines up left-to-left, centre-to-centre and right-to-left alike, which
   * is what makes this feel like a design surface rather than like magnets. The NEAREST match
   * inside the tolerance wins and the others are ignored: two guides on one axis would be two
   * different answers to where the thing goes.
   */
  private aligned(box: Box, name: string, edge: string | null): { box: Box; guides: Guide[] } {
    const near = this.neighbours(name);
    const guides: Guide[] = [];
    const out = { ...box };

    const pick = (mine: number[], theirs: number[]): { at: number; shift: number } | null => {
      let best: { at: number; shift: number } | null = null;
      for (const candidate of mine) {
        for (const line of theirs) {
          const shift = line - candidate;
          if (Math.abs(shift) <= ALIGN_REACH
            && (best === null || Math.abs(shift) < Math.abs(best.shift))) {
            best = { at: line, shift };
          }
        }
      }

      return best;
    };

    // A resize moves the edges the hand holds; a move takes the whole box, so all three of its
    // lines are candidates.
    const holdsWest = edge?.includes("w") ?? false;
    const holdsEast = edge?.includes("e") ?? false;
    const holdsNorth = edge?.includes("n") ?? false;
    const holdsSouth = edge?.includes("s") ?? false;

    const mineX = edge === null
      ? [box.left, box.left + box.width / 2, box.left + box.width]
      : [...(holdsWest ? [box.left] : []), ...(holdsEast ? [box.left + box.width] : [])];
    const mineY = edge === null
      ? [box.top, box.top + box.height / 2, box.top + box.height]
      : [...(holdsNorth ? [box.top] : []), ...(holdsSouth ? [box.top + box.height] : [])];

    const onX = pick(mineX, near.x);
    if (onX) {
      if (edge === null) {
        out.left = box.left + onX.shift;
      } else if (holdsWest) {
        out.left = box.left + onX.shift;
        out.width = box.width - onX.shift;
      } else {
        out.width = box.width + onX.shift;
      }

      guides.push({ axis: "x", at: onX.at });
    }

    const onY = pick(mineY, near.y);
    if (onY) {
      if (edge === null) {
        out.top = box.top + onY.shift;
      } else if (holdsNorth) {
        out.top = box.top + onY.shift;
        out.height = box.height - onY.shift;
      } else {
        out.height = box.height + onY.shift;
      }

      guides.push({ axis: "y", at: onY.at });
    }

    return { box: out, guides };
  }

  /**
   * A gesture's box, snapped. A move takes its ORIGIN to the grid and keeps its size, so a
   * control does not change shape by being moved; a resize takes the edges the hand is holding
   * and leaves the opposite ones exactly where they were, which is what keeps the other side of
   * the control still while one side is pulled.
   */
  private snapped(box: Box, edge: string | null, free = false, band = 0): Box {
    // ALT IS THE OVERRIDE, held rather than toggled: the one control that will not sit on the
    // grid is a reason to escape it for a moment, not to turn it off and forget to turn it back
    // on. Read per pointer event, so pressing or releasing Alt part-way through a drag changes
    // what the next movement does - which is how every design surface a developer already knows
    // behaves (the owner, 2026-08-16: "holding alt key should override any snapping").
    if (free) {
      return {
        left: Math.round(box.left),
        top: Math.round(box.top),
        width: Math.round(box.width),
        height: Math.round(box.height),
      };
    }

    /*
     * `band` IS THE CHROME ABOVE A CONTAINER'S CLIENT - a Frame's caption strip, a MultiPage's
     * tabs - and the grid is applied to the edge BELOW it.
     *
     * A Frame's rectangle starts at the top of its caption and the line it draws is four
     * points lower, so snapping the rectangle puts the caption on a grid dot and the visible
     * edge between two of them: the frame is on the grid and looks like the only thing on the
     * form that is not (the owner, 2026-08-16: "in snap to grid, the snapping for container
     * should be to the frame, not the label"). Everything else has a band of zero and is
     * unaffected.
     */
    if (edge === null) {
      return {
        ...box,
        left: this.onGrid(box.left),
        top: this.onGrid(box.top + band) - band,
      };
    }

    const out = { ...box };
    if (edge.includes("w")) {
      const right = box.left + box.width;
      out.left = this.onGrid(box.left);
      out.width = right - out.left;
    } else if (edge.includes("e")) {
      out.width = this.onGrid(box.left + box.width) - box.left;
    }

    if (edge.includes("n")) {
      const bottom = box.top + box.height;
      out.top = this.onGrid(box.top + band) - band;
      out.height = bottom - out.top;
    } else if (edge.includes("s")) {
      out.height = this.onGrid(box.top + box.height) - box.top;
    }

    return out;
  }

  /** Puts a thing - and the handles dressing it - at a box in points, without touching the
   * document: what the eye follows during a gesture, and the instant landing after one. */
  private paintBox(element: HTMLElement, box: Box): void {
    const form = element.classList.contains("dc-form");
    if (!form) {
      element.style.left = `${box.left * PT}px`;
      element.style.top = `${box.top * PT}px`;
    }

    element.style.width = `${box.width * PT}px`;
    element.style.height = `${box.height * PT}px`;

    // The form's handles ride inside it and follow by themselves; a control's overlay is a
    // sibling and has to be carried.
    const overlay = element.nextElementSibling;
    if (!form && overlay instanceof HTMLElement && overlay.classList.contains("dc-selection")) {
      overlay.style.left = element.style.left;
      overlay.style.top = element.style.top;
      overlay.style.width = element.style.width;
      overlay.style.height = element.style.height;
    }
  }

  /** The element a name stands for: "" is the FORM's own face. */
  private elementOf(name: string): HTMLElement | null {
    return name === "" ? this.canvasScroll.querySelector<HTMLElement>(".dc-form") : this.elementFor(name);
  }

  private elementFor(name: string): HTMLElement | null {
    return this.canvasScroll.querySelector<HTMLElement>(`.dc[data-control="${CSS.escape(name)}"]`);
  }

  /**
   * The document line that declares a thing - a control by name, the FORM for "" - with its
   * geometry lifted out. Null when no line names it, or when the line is not in a shape this can
   * rewrite - a half-typed header, an attribute whose quote never closes - because a gesture that
   * cannot write the document must not pretend to have changed anything.
   *
   * THE TAG IS SCANNED, NOT MATCHED WHOLE. An attribute value can hold anything, `>` included, so
   * where the tag ends is a question only a scan that tracks quoting can answer. This read the
   * PRE-TAG dialect until 2026-08-17 - `Kind Name "Caption" at x,y size WxH` - and so matched
   * nothing at all once the documents became XAML-shaped, which silently killed every canvas
   * gesture that writes: drag, resize, nudge, align, distribute, delete, reparent, add-child and
   * the toolbox drop all run through here and all refuse when it answers null.
   */
  private headerOf(name: string): HeaderLine | null {
    const form = name === "";
    const line = form ? 1 : this.headerLineOf(name);
    if (line === 0 || line > this.model.getLineCount()) {
      return null;
    }

    const text = this.model.getLineContent(line);
    const open = text.indexOf("<");
    if (open < 0 || !/^<[A-Za-z_]\w*[\s/>]/.test(text.slice(open))) {
      return null;
    }

    // The FORM is the only element whose tag may legitimately be the document's first line
    // without naming itself, so its identity is the line rather than a Name match.
    if (form && !/^<Form[\s/>]/.test(text.slice(open))) {
      return null;
    }

    let scan = open + 1;
    let quoted = false;
    let end = -1;
    while (scan < text.length) {
      const character = text[scan];
      if (character === '"') {
        quoted = !quoted;
      } else if (character === ">" && !quoted) {
        end = scan;
        break;
      }
      scan += 1;
    }

    // An unclosed tag or an unclosed value: not a line to rewrite.
    if (end < 0 || quoted) {
      return null;
    }

    const body = text.slice(open, end);
    const close = body.trimEnd().endsWith("/") ? "/>" : ">";
    const inner = close === "/>" ? body.slice(0, body.lastIndexOf("/")) : body;

    const read = (attribute: string): number | null => {
      const found = new RegExp(`\\b${attribute}\\s*=\\s*"(-?\\d+(?:\\.\\d+)?)"`, "i").exec(inner);
      return found === null ? null : Number.parseFloat(found[1]!);
    };

    const left = read("Left");
    const top = read("Top");
    const width = read("Width");
    const height = read("Height");

    // Everything BUT the geometry, tidied of the gaps taking it out leaves. Anything after the
    // tag on the same line - a comment, a close - is kept so a rewrite does not eat it.
    const head = text.slice(0, open)
      + inner.replace(/\s*\b(?:Left|Top|Width|Height)\s*=\s*"[^"]*"/gi, "").trimEnd();

    return {
      line,
      head,
      close,
      at: left === null || top === null ? null : { left, top },
      size: width === null || height === null ? null : { width, height },
      form,
    };
  }

  /** The 1-based line that declares a control, or 0 when the document names it nowhere. */
  /** The line an element's opening tag starts on, found by its `Name="..."` - which lands right
   * even when the tag's attributes wrap over several lines, because the Name is on the first. */
  private headerLineOf(name: string): number {
    const needle = new RegExp(`<[A-Za-z_]\\w*[^>]*\\bName\\s*=\\s*"${escapeForRegExp(name)}"`);
    return this.model.getLinesContent().findIndex((text) => needle.test(text)) + 1;
  }

  /**
   * A drag driven from the debug surface: the REAL pointer sequence, on the element a mouse
   * would actually hit - elementFromPoint at a point inside the control - so the act proves
   * hit-testing, the threshold and the commit rather than the arithmetic alone. Deltas are in
   * POINTS, the designer's unit and the document's. Returns what happened, for the row.
   */
  dragControl(name: string, dx: number, dy: number, alt = false): string {
    const element = this.elementFor(name);
    if (!element) {
      return `no control named ${name} on the canvas`;
    }

    // Into view first, for the reason resizeControl reaches for a handle: a narrow tab holds a
    // form wider than the box it sits in, and a hit test aimed at a clipped control answers
    // whatever paints at those coordinates instead - a dock strip, on a 704px window. A rect is
    // reported whether or not the element is visible, so nothing about the arithmetic says so.
    //
    // CENTRED rather than merely visible, because a drag TRAVELS: a control scrolled to the edge
    // of a 48-point canvas has no room on the side the gesture is heading for, and the pointer
    // ends outside the box - where it hits nothing and no container answers, so a drag that meant
    // to reparent simply clamps. Centring gives the gesture the whole canvas either way.
    element.scrollIntoView({ block: "center", inline: "center" });

    const box = element.getBoundingClientRect();
    // The centre first, the way a hand aims; a CONTAINER's centre belongs to its children, so
    // the near corner is the fallback - inside the container's own border, above no child.
    const centre = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
    const corner = { x: box.left + 2, y: box.top + 2 };
    const hits = (spot: { x: number; y: number }): HTMLElement | null =>
      document.elementFromPoint(spot.x, spot.y)?.closest<HTMLElement>(".dc") ?? null;
    const grabbed = [centre, corner].find((spot) => hits(spot)?.dataset.control === name);
    if (!grabbed) {
      const covering = document.elementFromPoint(centre.x, centre.y);
      return `${name} is not what a press at its own centre reaches - ${covering?.className || "nothing"} is`;
    }

    // The press goes to what the hit test answered - that IS the proof - and the rest to the
    // canvas, exactly where a captured pointer's events are delivered once a real drag has
    // the mouse. Past the threshold first, then the whole way: the two moves a hand makes.
    const landed = hits(grabbed) ?? element;
    this.lastGuides = [];
    this.sendGesture(landed, grabbed, dx, dy, alt);
    return `dragged ${name}${this.guideNote()}`;
  }

  /**
   * A resize driven from the debug surface: the named thing is SELECTED first, the way a hand
   * must, and then the real pointer sequence grabs its handle - `nw`, `n`, `ne`, `e`, `se`,
   * `s`, `sw`, `w` - and pulls it by a delta in POINTS. "" is the form's own frame. The press
   * goes through the hit test at the handle, so a handle nothing can reach fails the act.
   */
  resizeControl(name: string, edge: string, dx: number, dy: number, alt = false): string {
    if (this.selectedName !== name) {
      this.select(name);
    }

    const handle = this.canvasScroll.querySelector<HTMLElement>(`.dc-handle-${CSS.escape(edge)}`);
    if (!handle) {
      return `no ${edge} handle stands on ${name === "" ? "the form" : name}`;
    }

    // Into view first, the way a hand reaches a handle past the edge of the canvas: on a narrow
    // tab the form is wider than the box that holds it, and a hit test aimed off-screen answers
    // nothing - which is a true answer to the wrong question. `nearest` alone leaves the handle
    // flush against the edge, where a 7px target is still half outside, so the scroll finishes
    // by pulling its centre a margin clear.
    handle.scrollIntoView({ block: "nearest", inline: "nearest" });
    const reach = handle.getBoundingClientRect();
    const face = this.canvasScroll.querySelector<HTMLElement>(".dc-form-client")?.getBoundingClientRect();
    if (face) {
      this.scrollToPoint(
        this.toPoints(reach.left + reach.width / 2 - face.left),
        this.toPoints(reach.top + reach.height / 2 - face.top));
    }

    const box = handle.getBoundingClientRect();
    const spot = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
    const reached = document.elementFromPoint(spot.x, spot.y)?.closest<HTMLElement>(".dc-handle");
    if (reached?.dataset.edge !== edge) {
      return `the ${edge} handle is not what a press on it reaches - ${reached?.className || "nothing"} is`;
    }

    this.lastGuides = [];
    this.sendGesture(reached, spot, dx, dy, alt);
    return `resized ${name === "" ? "the form" : name}${this.guideNote()}`;
  }

  /**
   * Brings a point of the FORM - in points from its client origin - inside the canvas's visible
   * box, with a margin so it is not flush against an edge. What a hand does by scrolling before
   * it reaches, and what any gesture aimed by coordinates has to do first on a tab too short to
   * show the whole form.
   */
  private scrollToPoint(left: number, top: number): void {
    const client = this.canvasScroll.querySelector<HTMLElement>(".dc-form-client");
    if (!client) {
      return;
    }

    const face = client.getBoundingClientRect();
    const box = this.canvasScroll.getBoundingClientRect();
    const margin = 20;
    const x = face.left + this.toPixels(left);
    const y = face.top + this.toPixels(top);

    if (y < box.top + margin) {
      this.canvasScroll.scrollTop -= box.top + margin - y;
    } else if (y > box.bottom - margin) {
      this.canvasScroll.scrollTop += y - (box.bottom - margin);
    }

    if (x < box.left + margin) {
      this.canvasScroll.scrollLeft -= box.left + margin - x;
    } else if (x > box.right - margin) {
      this.canvasScroll.scrollLeft += x - (box.right - margin);
    }
  }

  /**
   * A toolbox drag driven from the debug surface: the REAL pointer sequence, out of the palette
   * button for `kind` and onto a point on the canvas given in POINTS from the form's own client
   * origin - the same coordinates the document carries, so a row can ask for 40,120 and then
   * read `at 40,120` back. Answers the name the drop gave the new control, which is the one
   * thing the caller cannot know in advance.
   */
  addFromToolbox(kind: string, left: number, top: number): string {
    const button = this.toolbox.querySelector<HTMLElement>(`.designer-tool[data-kind="${CSS.escape(kind)}"]`);
    if (!button) {
      return `no ${kind} in the toolbox`;
    }

    const client = this.canvasScroll.querySelector<HTMLElement>(".dc-form-client");
    if (!client) {
      return "the form's client area is not drawn";
    }

    // A drop lands wherever the POINTER is, and a pointer cannot be somewhere the canvas is not
    // showing: on a short tab the form is taller than the box that holds it and the drop point
    // is off the bottom, where the hit test finds nothing at all. So the canvas is scrolled to
    // the point first, which is also what a hand does before it lets go.
    this.scrollToPoint(left, top);

    const from = button.getBoundingClientRect();
    const face = client.getBoundingClientRect();
    const to = { x: face.left + this.toPixels(left), y: face.top + this.toPixels(top) };
    const send = (target: EventTarget, type: string, x: number, y: number): void => {
      target.dispatchEvent(new PointerEvent(type, {
        bubbles: true,
        cancelable: true,
        pointerId: 1,
        isPrimary: true,
        button: 0,
        buttons: type === "pointerup" ? 0 : 1,
        clientX: x,
        clientY: y,
      }));
    };

    // The press goes to the button; the move and the drop go to the WINDOW, which is where the
    // gesture listens once it has left the palette - the same place a real pointer's events go.
    const before = new Set(this.model.getLinesContent());
    send(button, "pointerdown", from.left + from.width / 2, from.top + from.height / 2);
    send(window, "pointermove", to.x, to.y);
    send(window, "pointerup", to.x, to.y);

    const added = this.model.getLinesContent().find((line) => !before.has(line));
    const named = added ? /^\s+\S+\s+(\S+)/.exec(added)?.[1] : undefined;
    if (named) {
      return `added ${named}`;
    }

    // A refusal says WHAT the pointer found, because "nothing landed" is true of a point off the
    // form, a point over another pane, and a point the canvas could not scroll to - three
    // different faults with one sentence between them until now.
    const under = document.elementFromPoint(Math.round(to.x), Math.round(to.y));
    const room = this.canvasScroll.getBoundingClientRect();
    return `nothing landed at ${left},${top}: the pointer was at `
      + `${Math.round(to.x)},${Math.round(to.y)} over ${under?.className || under?.tagName || "nothing"}`
      + `, and the canvas is ${Math.round(room.left)},${Math.round(room.top)}`
      + ` to ${Math.round(room.right)},${Math.round(room.bottom)}`;
  }

  /**
   * A delete driven from the debug surface: the control is SELECTED the way a hand must, and
   * then the real Delete key is pressed on the canvas - the same listener, the same commit - so
   * the act cannot pass on a path the keyboard does not take.
   */
  deleteControl(name: string): string {
    if (this.selectedName !== name) {
      this.select(name);
    }

    if (this.selectedName !== name) {
      return `${name} is not on the canvas to select`;
    }

    this.canvasScroll.dispatchEvent(new KeyboardEvent("keydown", {
      key: "Delete",
      bubbles: true,
      cancelable: true,
    }));
    return this.headerLineOf(name) === 0 ? `deleted ${name}` : `${name} is still in the document`;
  }

  /**
   * A rubber band dragged over the form's own ground, for the debug surface: the real pointer
   * sequence on the real client, so the act exercises the marquee rather than the selection it
   * would have produced. Corners are POINTS from the form's client origin.
   */
  marqueeOver(left: number, top: number, right: number, bottom: number): string {
    const client = this.canvasScroll.querySelector<HTMLElement>(".dc-form-client");
    if (!client) {
      return "the canvas is not showing a form";
    }

    const room = client.getBoundingClientRect();
    const from = { x: Math.round(room.left + this.toPixels(left)), y: Math.round(room.top + this.toPixels(top)) };
    const to = { x: Math.round(room.left + this.toPixels(right)), y: Math.round(room.top + this.toPixels(bottom)) };
    const send = (target: EventTarget, type: string, x: number, y: number): void => {
      target.dispatchEvent(new PointerEvent(type, {
        bubbles: true,
        cancelable: true,
        pointerId: 1,
        isPrimary: true,
        button: 0,
        buttons: type === "pointerup" ? 0 : 1,
        clientX: x,
        clientY: y,
      }));
    };

    send(client, "pointerdown", from.x, from.y);
    send(this.canvasScroll, "pointermove", from.x + DRAG_THRESHOLD + 1, from.y + DRAG_THRESHOLD + 1);
    send(this.canvasScroll, "pointermove", to.x, to.y);
    send(this.canvasScroll, "pointerup", to.x, to.y);
    return `banded ${left},${top} to ${right},${bottom}: ${this.selection().join(", ") || "nothing"}`;
  }

  /**
   * A tab opened from the debug surface: the REAL press on the real tab, through the canvas's
   * one press path, so the act cannot pass on a route a pointer does not take.
   *
   * `which` is a page's name, a tab's label, or a 1-based position - whichever the caller has.
   * A TabStrip has only the last two, since its tabs have no names.
   */
  openTabOn(container: string, which: string): string {
    const strip = [...this.canvasScroll.querySelectorAll<HTMLElement>(".dc-page-strip")]
      .find((one) => (one.dataset.container ?? "").toLowerCase() === container.toLowerCase());
    if (!strip) {
      return `the canvas draws no tabbed container called ${container}`;
    }

    const tabs = [...strip.querySelectorAll<HTMLElement>(".dc-page-tab")];
    const at = Number(which);
    const tab = Number.isInteger(at) && at >= 1
      ? tabs[at - 1]
      : tabs.find((one) => (one.dataset.page ?? "").toLowerCase() === which.toLowerCase())
        ?? tabs.find((one) => (one.textContent ?? "").toLowerCase() === which.toLowerCase());
    if (!tab) {
      return `${container} has no tab ${which} (it has ${tabs.length})`;
    }

    const box = tab.getBoundingClientRect();
    tab.dispatchEvent(new PointerEvent("pointerdown", {
      bubbles: true,
      cancelable: true,
      pointerId: 1,
      isPrimary: true,
      button: 0,
      buttons: 1,
      clientX: Math.round(box.left + box.width / 2),
      clientY: Math.round(box.top + box.height / 2),
    }));
    return `opened ${which} on ${container}`;
  }

  /**
   * Right-clicks a tab strip, the way the menu is actually reached: a real `contextmenu` event
   * on the tab, since the menu hangs off a listener rather than off a method. The caller reads
   * the items off the menu and picks one with chooseMenuItem.
   */
  openTabMenu(container: string, which: string): string {
    const strip = [...this.canvasScroll.querySelectorAll<HTMLElement>(".dc-page-strip")]
      .find((one) => (one.dataset.container ?? "").toLowerCase() === container.toLowerCase());
    if (!strip) {
      return `the canvas draws no tabbed container called ${container}`;
    }

    const tabs = [...strip.querySelectorAll<HTMLElement>(".dc-page-tab")];
    const at = Number(which);
    const aim = which === ""
      ? strip
      : Number.isInteger(at) && at >= 1
        ? tabs[at - 1]
        : tabs.find((one) => (one.dataset.page ?? "").toLowerCase() === which.toLowerCase());
    if (!aim) {
      return `${container} has no tab ${which} (it has ${tabs.length})`;
    }

    const box = aim.getBoundingClientRect();
    const x = Math.round(box.left + box.width / 2);
    const y = Math.round(box.top + box.height / 2);

    // The PRESS as well as the menu event, because a real right-click is both and the second
    // button's press is what dismisses a menu already standing. Without it, a right-click that
    // opens NOTHING - a TabStrip's strip - leaves the last menu on screen and a caller reading
    // the items back is told about a menu it did not open.
    aim.dispatchEvent(new PointerEvent("pointerdown", {
      bubbles: true, cancelable: true, pointerId: 1, isPrimary: true, button: 2, buttons: 2,
      clientX: x, clientY: y,
    }));
    aim.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: x, clientY: y }));
    return `right-clicked ${which === "" ? `the strip of ${container}` : which}`;
  }

  /** The pointer sequence both instruments send: press where the hit test answered, one move
   * past the threshold, one move the whole way, and the release - all but the press delivered
   * to the canvas, which is where a captured pointer's events go. */
  private sendGesture(
    target: EventTarget, from: { x: number; y: number }, dx: number, dy: number, alt = false,
  ): void {
    const send = (to: EventTarget, type: string, x: number, y: number): void => {
      to.dispatchEvent(new PointerEvent(type, {
        bubbles: true,
        cancelable: true,
        pointerId: 1,
        isPrimary: true,
        button: 0,
        buttons: type === "pointerup" ? 0 : 1,
        clientX: x,
        clientY: y,
        // Held for the whole gesture, which is what a hand on Alt does; the move and the
        // drop both read it, so the preview and the commit agree about the override.
        altKey: alt,
      }));
    };

    send(target, "pointerdown", from.x, from.y);
    send(this.canvasScroll, "pointermove", from.x + DRAG_THRESHOLD + 1, from.y);
    send(this.canvasScroll, "pointermove", from.x + this.toPixels(dx), from.y + this.toPixels(dy));
    send(this.canvasScroll, "pointerup", from.x + this.toPixels(dx), from.y + this.toPixels(dy));
  }

  private dressSelection(): void {
    for (const dressed of this.canvasScroll.querySelectorAll(".dc-selection")) {
      dressed.remove();
    }
    for (const marked of this.canvasScroll.querySelectorAll(".dc-selected")) {
      marked.classList.remove("dc-selected");
    }

    // The rest of a group wears a boundary and no handles: only the anchor can be pulled, and a
    // grip on a control that will not answer it is a promise the canvas would not keep. The
    // native designer says the same thing with hollow handles rather than white ones.
    for (const extra of this.extras) {
      const element = this.elementFor(extra);
      if (!element) {
        continue;
      }

      const mark = document.createElement("div");
      mark.className = "dc-selection dc-selection-extra";
      mark.style.left = element.style.left;
      mark.style.top = element.style.top;
      mark.style.width = element.style.width;
      mark.style.height = element.style.height;
      element.insertAdjacentElement("afterend", mark);
      element.classList.add("dc-selected");
    }

    const name = this.selectedName;
    if (name === null) {
      return;
    }

    const target = name === ""
      ? this.canvasScroll.querySelector<HTMLElement>(".dc-form")
      : this.canvasScroll.querySelector<HTMLElement>(`.dc[data-control="${CSS.escape(name)}"]`);
    if (!target) {
      return;
    }

    // The handles live on an OVERLAY beside the control rather than inside it: every
    // control box clips its overflow, so handles ON the boundary would be half-eaten.
    // Pointer events pass through; M5 is where the handles learn to drag.
    const overlay = document.createElement("div");
    overlay.className = "dc-selection";

    // A PAGE wears no handles: it has neither a position nor a size of its own - the MultiPage
    // gives it both - so eight grips that can pull nothing would be a promise the canvas cannot
    // keep. Its outline lies over the page itself, the way the form's does over the form.
    const page = target.dataset.kind === "Page";
    if (!page) {
      for (const spot of ["nw", "n", "ne", "e", "se", "s", "sw", "w"]) {
        const handle = document.createElement("span");
        handle.className = `dc-handle dc-handle-${spot}`;
        // The edge rides the element: a press reads which way to pull from what it grabbed.
        handle.dataset.edge = spot;
        overlay.appendChild(handle);
      }
    }

    if (name === "" || page) {
      // Neither the form nor a page clips, so the overlay lives inside, over the whole face.
      overlay.classList.add(page ? "dc-selection-page" : "dc-selection-form");
      target.appendChild(overlay);
    } else {
      overlay.style.left = target.style.left;
      overlay.style.top = target.style.top;
      overlay.style.width = target.style.width;
      overlay.style.height = target.style.height;
      target.insertAdjacentElement("afterend", overlay);

      // A SELECTED control offers the move, and says so with the cursor (the owner's rule,
      // 2026-08-15). Only the selected one, and only a control that has a position to change -
      // the form is where the canvas puts it, and a Page fills its MultiPage, so neither
      // reaches here.
      target.classList.add("dc-selected");
    }
  }

  /** The markup caret onto the selected thing's line: the form's is line 1, a control's is
   * the line that names it. */
  /**
   * The selection's BLOCK in the document: its header line and everything indented under it -
   * a control's properties, a container's children. The form's block is the whole document.
   *
   * A control is a block in this language, not a line, so a caret on the header said "here is
   * where it starts" and left the reader to work out where it stopped (the owner, 2026-08-16:
   * "I'd like the full row or block to highlight"). Answers null when the document holds no
   * such thing, which is every moment between a canvas gesture and the projection that follows.
   */
  private blockOf(name: string): { from: number; to: number } | null {
    const lines = this.model.getLinesContent();
    if (name === "") {
      return { from: 1, to: Math.max(1, lines.length) };
    }

    const from = this.headerLineOf(name);
    if (from === 0) {
      return null;
    }

    // The element's RANGE: from its opening tag to whichever line closes it. A self-closing tag
    // is its own end, and an element that wraps across lines ends where its `>` does - both of
    // which the old indent walk could not see, because it measured spaces rather than tags.
    let depth = 0;
    let opened = false;
    for (let at = from; at <= lines.length; at += 1) {
      const text = lines[at - 1] ?? "";
      for (const tag of text.matchAll(/<(\/?)([A-Za-z_][\w]*)((?:[^>"]|"[^"]*")*)>/g)) {
        if (tag[1] === "/") {
          depth -= 1;
        } else if (!(tag[3] ?? "").trimEnd().endsWith("/")) {
          depth += 1;
          opened = true;
        } else {
          opened = true;
        }
      }

      // Opened and back to level: everything this element holds has been passed.
      if (opened && depth <= 0) {
        return { from, to: at };
      }
    }

    return { from, to: Math.max(from, lines.length) };
  }

  /** What the canvas selection lights up in the document, cleared when nothing is selected. */
  private blockMarks: string[] = [];

  /** The handler decorations, kept apart from the selection's so neither clears the other. */
  private handlerMarks: string[] = [];

  /**
   * WHICH CONTROLS HAVE CODE BEHIND THEM, shown against them in the document.
   *
   * A form is two halves and the markup only ever showed one. `OkButton_Click` exists in the
   * code-behind and nothing on the design side said so, which is the question a developer asks
   * constantly - is anything listening to this button (the owner, 2026-08-18).
   *
   * AN ANNOTATION, NOT AN ATTRIBUTE, and the distinction is the dialect's own rule: the document
   * spells what an apply can WRITE, and a handler is code. Spelling `Click="OkButton_Click"` would
   * invite an apply to create or delete a Sub, which is not what this language does. So it rides
   * as a decoration - visible, not editable, and absent from anything the parser reads.
   *
   * The names come from the outline route, which is what the tree's own unfolded list uses, so
   * there is one answer about a module's procedures rather than two that agree today.
   */
  /**
   * Re-reads the handlers. Safe to call from anywhere and as often as anything likes.
   *
   * TWO CALLERS RACED THE FIRST TIME THIS SHIPPED, and the symptom was the feature appearing for
   * a tab opened fresh and vanishing on a page reload: the constructor started a read against an
   * empty model, `update` started another against the real one, and whichever host answer came
   * back LAST won. So a read carries a ticket, and an answer that is not the newest is dropped.
   */
  refreshHandlers(): void {
    void this.paintHandlers();
  }

  /** The ticket of the newest handler read. An answer holding an older one is stale by
   * definition: something asked again after it, against a document it has not seen. */
  private handlerRead = 0;

  /** Markup line -> the handlers annotated on it, so a click can say where to go. */
  private readonly handlerLines = new Map<number, { name: string; line: number }[]>();

  /**
   * A CLICK ON THE ANNOTATION OPENS THE HANDLER, in the code half, at its Sub.
   *
   * Monaco reports a click on injected text as a content click PAST the end of the line - the
   * annotation is not in the model, so there is no column of the document under the pointer. That
   * is the whole test: a press at or beyond the last column of a line that carries an annotation
   * is a press on the annotation, and nothing else can be.
   *
   * With several handlers on one control the first is taken. Splitting the injected text into
   * separately clickable runs needs a decoration each and a column arithmetic that guesses at
   * glyph widths; going to the first of a control's handlers is one keystroke from the rest, and
   * the code half's own outline is beside it.
   */
  private onMarkupClick(event: monaco.editor.IEditorMouseEvent): void {
    const at = event.target.position;
    if (!at || event.event.rightButton) {
      return;
    }

    const found = this.handlerLines.get(at.lineNumber);
    if (!found || found.length === 0 || at.column < this.model.getLineMaxColumn(at.lineNumber)) {
      return;
    }

    this.deps.openHandler(found[0]!.line);
  }

  private async paintHandlers(): Promise<void> {
    const ticket = ++this.handlerRead;
    const procedures = await this.deps.handlers();
    if (ticket !== this.handlerRead) {
      return;
    }

    if (procedures === null) {
      // A shrug: keep whatever is already shown rather than claiming there are none.
      return;
    }

    // `OkButton_Click` is OkButton's Click. Split at the LAST underscore, because a control may
    // hold one in its own name and an event never does.
    const byControl = new Map<string, { name: string; line: number }[]>();
    for (const procedure of procedures) {
      const at = procedure.name.lastIndexOf("_");
      if (at <= 0) {
        continue;
      }

      const control = procedure.name.slice(0, at).toLowerCase();
      byControl.set(control, [
        ...(byControl.get(control) ?? []),
        { name: procedure.name, line: procedure.line },
      ]);
    }

    // Which handler sits on which markup line, so a CLICK on the annotation knows where to go.
    // Rebuilt with the decorations rather than kept in step with them, because two tables of the
    // same fact is how one of them goes stale.
    this.handlerLines.clear();

    const marks: monaco.editor.IModelDeltaDecoration[] = [];
    const lines = this.model.getLinesContent();
    for (const [at, text] of lines.entries()) {
      const name = /\bName="([^"]*)"/.exec(text)?.[1];
      if (name === undefined) {
        continue;
      }

      /*
       * THE FORM'S OWN HANDLERS ARE SPELLED `UserForm_`, whatever the form is called.
       *
       * `UserForm_Initialize` on a form named EntryForm is MSForms' convention, not a control
       * called UserForm - so a lookup by the element's Name finds nothing and the form's own line
       * was the one line in the document with no annotation, which is the line a developer most
       * expects one on (the owner, 2026-08-18). The Form element answers to both: its own name,
       * for anything written against it by name, and the literal the runtime uses.
       */
      const isForm = /^\s*<Form\b/.test(text);
      const found = [
        ...(byControl.get(name.toLowerCase()) ?? []),
        ...(isForm ? byControl.get("userform") ?? [] : []),
      ];
      if (found.length === 0) {
        continue;
      }

      this.handlerLines.set(at + 1, found);
      const column = this.model.getLineMaxColumn(at + 1);
      marks.push({
        range: new monaco.Range(at + 1, column, at + 1, column),
        options: {
          after: {
            // THE HANDLER'S OWN NAME, not just the event: `OkButton_Click` is what the developer
            // will search for, what the code half calls it, and what makes the annotation
            // self-explaining rather than a hint that needs decoding.
            //
            // LABELLED, so the yellow run is never mistaken for something the tag itself carries;
            // and NO LEADING SPACES, because they are part of the clickable run and the hover
            // underline drew itself under the gap as well as the name. The gap is a margin on the
            // class instead, which sits outside the text decoration.
            content: `${found.length > 1 ? "Events" : "Event"}: `
              + found.map((one) => one.name).join(", "),
            inlineClassName: "designer-handler-mark",
          },
          showIfCollapsed: true,
        },
      });
    }

    this.handlerMarks = this.editor.deltaDecorations(this.handlerMarks, marks);
  }

  /** @param moveCaret See `select`. The WASH always follows the selection - a highlight is a
   * picture of what is selected and is never in anybody's way - and only the caret is held back. */
  private revealInMarkup(name: string, moveCaret = true): void {
    const block = this.blockOf(name);
    if (!block) {
      return;
    }

    // The whole block, highlighted, with the caret at the end of its header so typing goes
    // where a hand would put it. The form's block is the document, and highlighting all of it
    // says nothing, so the form gets the caret and no wash.
    this.blockMarks = this.editor.deltaDecorations(this.blockMarks, name === "" ? [] : [{
      range: new monaco.Range(block.from, 1, block.to, this.model.getLineMaxColumn(block.to)),
      options: {
        isWholeLine: true,
        className: "designer-block-mark",
        overviewRuler: {
          color: "rgba(100, 160, 255, 0.5)",
          position: monaco.editor.OverviewRulerLane.Left,
        },
      },
    }]);

    if (!moveCaret) {
      return;
    }

    this.editor.setPosition({
      lineNumber: block.from,
      column: this.model.getLineMaxColumn(block.from),
    });
    this.editor.revealLineInCenterIfOutsideViewport(block.from);
  }

  /**
   * The other direction: the caret's line names a control, and the canvas selects it.
   *
   * A line INSIDE a control's block counts as that control - a property line belongs to the
   * thing it describes, and a developer clicking one means that thing (the owner, 2026-08-16:
   * "if I click on a row in the markup editor, I'd like the corresponding control in the
   * designer to focus"). The form's own line, and anything outside every block, selects the
   * form. Nothing is written: this is a selection, and the document is already what it is.
   */
  private selectFromMarkup(line: number): void {
    // THE ELEMENT THE CARET IS IN, not the line it is on. In the tagged dialect a control is a
    // RANGE - an opening tag, its children, its close - so a caret three lines into a Frame's
    // body belongs to whichever child element encloses it, and to the Frame when none does. The
    // walk keeps the same stack the parser keeps, up to the caret, and the innermost still-open
    // element is the answer.
    const text = this.model.getValueInRange({
      startLineNumber: 1,
      startColumn: 1,
      endLineNumber: line,
      endColumn: this.model.getLineMaxColumn(line),
    });

    const stack: string[] = [];
    let last = "";
    let at = 0;
    while (at < text.length) {
      const open = text.indexOf("<", at);
      if (open < 0) {
        break;
      }

      if (text.startsWith("<!--", open)) {
        const done = text.indexOf("-->", open);
        if (done < 0) {
          break;
        }

        at = done + 3;
        continue;
      }

      let cursor = open + 1;
      let quoted = false;
      let end = -1;
      while (cursor < text.length) {
        const character = text[cursor];
        if (character === '"') {
          quoted = !quoted;
        } else if (character === ">" && !quoted) {
          end = cursor;
          break;
        }
        cursor++;
      }

      const body = text.slice(open + 1, end < 0 ? text.length : end);
      const named = /\bName\s*=\s*"([^"]*)"/.exec(body)?.[1] ?? "";

      // An element that OPENS on the caret's own line is what the caret is on, self-closing or
      // not - that is the click-on-a-control case, and it beats whatever encloses it.
      const onCaretLine = text.slice(0, open).split("\n").length === line;
      if (!body.startsWith("/") && named !== "" && onCaretLine) {
        last = named;
      }

      if (end < 0) {
        break;
      }

      if (body.startsWith("/")) {
        stack.pop();
      } else if (!body.trimEnd().endsWith("/")) {
        stack.push(named);
      }

      at = end + 1;
    }

    // The innermost element still open at the caret, or the one that opened on its line. The
    // Form's own tag sits at the bottom of the stack and answers as the empty name.
    let found = last !== "" ? last : (stack[stack.length - 1] ?? "");
    if (stack.length <= 1 && last === "") {
      found = "";
    }

    // Against the PROJECTION rather than against the drawn elements: a control on a page that
    // is not open has no element and is still perfectly real, and selecting it is what opens
    // its page. Only a name the canvas does not know at all - a half-typed line, a control
    // added since the last parse - falls back to the form.
    if (found !== "" && this.kindOf(found) === "") {
      found = "";
    }

    if (this.selectedName !== found) {
      // The caret is already where the selection came from, so it stays put. See `select`.
      this.select(found, false, false);
    }
  }

  private renderCanvas(payload: FormMarkupPayload): void {
    this.shownPayload = payload;
    this.canvasScroll.replaceChildren();

    const form = document.createElement("div");
    form.className = "dc-form";
    form.style.width = `${Math.max(60, (payload.form?.width ?? 240)) * PT}px`;
    form.style.height = `${Math.max(40, (payload.form?.height ?? 180)) * PT}px`;
    // The REAL colours, through the machine's own palette - what the form surface a user
    // without xlide sees would paint. The stylesheet's grey is only the fallback.
    if (payload.form?.backColor) {
      form.style.background = payload.form.backColor;
    }
    if (payload.form?.foreColor) {
      form.style.color = payload.form.foreColor;
    }

    const titlebar = document.createElement("div");
    titlebar.className = "dc-form-title";
    const titleText = document.createElement("span");
    titleText.className = "dc-form-title-text";
    titleText.textContent = payload.form?.caption ?? this.id.module;
    const titleClose = document.createElement("span");
    titleClose.className = "dc-form-close";
    titleClose.textContent = "×";
    titlebar.append(titleText, titleClose);
    form.appendChild(titlebar);

    const client = document.createElement("div");
    client.className = "dc-form-client";
    // The REAL client area when the model says it - the Frame's own rule at form scale: side
    // borders split the width difference, and what remains of the height difference above
    // the client is the title bar. Without this the full outer rect stood in for the client,
    // and every control sat measurably off what the running form shows (the owner's
    // side-by-side, 2026-08-13). The floating-bar fallback keeps forms whose insides the
    // model will not answer rendering as before.
    const outer = payload.form;
    if (outer?.insideWidth && outer?.insideHeight && outer.width && outer.height) {
      const side = Math.max(0, (outer.width - outer.insideWidth) / 2) * PT;
      const top = Math.max(0, (outer.height - outer.insideHeight) * PT - side);
      client.style.left = `${side}px`;
      client.style.right = `${side}px`;
      client.style.bottom = `${side}px`;
      client.style.top = `${top}px`;
      titlebar.style.height = `${top}px`;
      form.classList.add("chromed");
    }
    // The form's own picture is held rather than painted here: the GRID paints on this same
    // element, and showGrid composes the two so neither wipes the other.
    this.formGround = payload.form?.picture ? pictureLayer(payload.form.picture) : null;

    form.appendChild(client);

    // Flat rows into a tree by parent NAME, the walk's own shape. A parent the rows never
    // declare (a page of a MultiPage read shallowly) lands its child at the top level rather
    // than dropping it: visible in the wrong place beats invisible.
    const rows = payload.controls ?? [];
    const byParent = new Map<string, FormMarkupControl[]>();
    const names = new Set(rows.map((row) => row.name.toLowerCase()));
    for (const row of rows) {
      const parent = row.parent && names.has(row.parent.toLowerCase()) ? row.parent.toLowerCase() : "";
      const kin = byParent.get(parent) ?? [];
      kin.push(row);
      byParent.set(parent, kin);
    }

    const renderInto = (host: HTMLElement, parentKey: string): void => {
      for (const row of byParent.get(parentKey) ?? []) {
        host.appendChild(this.renderControl(row, byParent, renderInto));
      }
    };
    renderInto(client, "");

    // The form rides a STAGE: a transform does not change layout, so without a box of the SCALED
    // size the scroll bars would still describe the form at 100% and half of a zoomed form would
    // be unreachable. The margin lives here too, for the reason it left the scroll port.
    const stage = document.createElement("div");
    stage.className = "dc-stage";
    stage.appendChild(form);
    this.canvasScroll.appendChild(stage);
    this.applyZoom();

    // A re-render replaces every element, so the selection is dressed again onto the new
    // ones; a selected control the new picture no longer holds simply loses its handles.
    this.dressSelection();

    // And the grid, for the same reason: the client it is painted on is one of those elements.
    this.showGrid();
  }

  private renderControl(
    row: FormMarkupControl,
    byParent: Map<string, FormMarkupControl[]>,
    renderInto: (host: HTMLElement, parentKey: string) => void,
  ): HTMLElement {
    const box = document.createElement("div");
    box.className = `dc dc-${row.type}`;
    box.style.left = `${(row.left ?? 0) * PT}px`;
    box.style.top = `${(row.top ?? 0) * PT}px`;
    box.style.width = `${Math.max(4, (row.width ?? 24)) * PT}px`;
    box.style.height = `${Math.max(4, (row.height ?? 12)) * PT}px`;
    box.title = `${row.name} (${row.type})`;
    box.dataset.control = row.name;
    // The kind rides the element because the gestures ask about it: a Page has no position
    // of its own to drag, and the harness reads it to aim.
    box.dataset.kind = row.type;

    // The control's own font, points scaled like every bound. A Label's BackColor is NOT
    // painted: its BackStyle defaults to transparent and the walk does not read BackStyle
    // yet, so painting would colour what the real surface leaves clear - the honest
    // approximation goes the other way.
    if (row.fontName) {
      box.style.fontFamily = `"${row.fontName}", Tahoma, sans-serif`;
    }
    if (row.fontSize) {
      box.style.fontSize = `${row.fontSize * PT}px`;
    }
    if (row.fontBold) {
      box.style.fontWeight = "bold";
    }
    if (row.fontItalic) {
      box.style.fontStyle = "italic";
    }
    if (row.backColor && row.type !== "Label") {
      box.style.background = row.backColor;
    }
    if (row.foreColor) {
      box.style.color = row.foreColor;
    }

    const caption = row.caption ?? "";

    switch (row.type) {
      case "Frame": {
        const inner = document.createElement("div");
        inner.className = "dc-frame-client";
        // The REAL client area when the model says it: side borders split the width
        // difference evenly, and what remains of the height difference above the client is
        // the caption strip - derived, not guessed, which is the parity rule.
        // The sides and the bottom are the model's - a border each way, split evenly. The TOP
        // is the caption band, which the model's InsideHeight understates; see the constant.
        if (row.insideWidth && row.insideHeight && row.width && row.height) {
          const side = Math.max(0, (row.width - row.insideWidth) / 2) * PT;
          inner.style.left = `${side}px`;
          inner.style.right = `${side}px`;
          inner.style.bottom = `${side}px`;
        }

        inner.style.top = `${FRAME_CLIENT_TOP}px`;

        /*
         * THE RULE IS DRAWN THROUGH THE CAPTION, NOT AT THE CONTROL'S TOP.
         *
         * A Frame at top 112 shows its line about four points lower and its first child about
         * nine, because the caption band belongs to the control's own rectangle. The canvas
         * drew the line at the box edge, which lines a frame up with the button beside it on
         * screen and lines them up nowhere else - the owner's side-by-side of the canvas
         * against the running form caught it (2026-08-16, "the xlide designer doesn't
         * accurately represent how the controls align with each other in actual runtime").
         */
        const rule = document.createElement("div");
        rule.className = "dc-frame-rule";
        rule.style.top = `${FRAME_RULE_TOP}px`;
        box.appendChild(rule);

        // Straddling the rule: the band is the caption's own line box, so putting the caption
        // at the top of it puts the rule through its middle.
        const legend = document.createElement("div");
        legend.className = "dc-frame-caption";
        legend.textContent = caption;
        legend.style.top = "0px";
        box.appendChild(legend);

        // A Frame's picture paints on its CLIENT, not on the whole rectangle: the caption band
        // belongs to the frame's chrome, and the runtime does not paint the picture behind it.
        if (row.picture) {
          paintPictureSurface(inner, row.picture);
        }

        box.appendChild(inner);
        renderInto(inner, row.name.toLowerCase());
        break;
      }

      case "MultiPage":
      case "TabStrip": {
        // The rule goes in FIRST so the tabs paint over it, the way a tab control's selected
        // tab breaks the line it sits on.
        const edge = document.createElement("div");
        edge.className = "dc-page-rule";
        box.appendChild(edge);

        const strip = document.createElement("div");
        strip.className = "dc-page-strip";
        // The strip carries the container it belongs to, because a tab's press and its menu
        // both ask which MultiPage they are about and a tab has no other way to say.
        strip.dataset.container = row.name;
        strip.dataset.kind = row.type;
        // A MultiPage's headers are its Page children and a TabStrip's are its Tab children -
        // both in the projection since 2026-08-16, so the canvas follows a tab TYPED into the
        // document the same way it follows a page. `row.tabs` is the older road, kept for a
        // draft whose strip has no child rows yet.
        const kin = byParent.get(row.name.toLowerCase()) ?? [];
        const pages = kin.filter((one) => one.type !== "Tab");
        const headers = kin.length > 0
          ? kin.map((one) => one.caption ?? one.name)
          : row.tabs ?? [];

        // Which one is OPEN: the developer's own choice while it still exists, and the first
        // otherwise - a page deleted or renamed away puts the picture back at the front rather
        // than leaving the container blank.
        const open = this.openTabOf(row, pages, headers.length);
        for (const [index, header] of headers.entries()) {
          const tab = document.createElement("span");
          tab.className = "dc-page-tab" + (index === open ? " current" : "");
          tab.textContent = header;
          tab.dataset.tab = String(index);
          // The tab carries its own row's NAME - a Page's for a MultiPage, a Tab's for a strip -
          // so the menu and the acts can name what they are acting on. `data-page` stays the
          // attribute for a page's body, which only a MultiPage has.
          const own = kin[index];
          if (own) {
            tab.dataset.name = own.name;
            if (row.type !== "TabStrip") {
              tab.dataset.page = own.name;
            }
          }
          strip.appendChild(tab);
        }
        box.appendChild(strip);

        // The OPEN page's content shows. A TabStrip has no page to show and never will: its
        // tabs are an index, not a container, and the runtime draws the same controls under
        // every one of them.
        const first = row.type === "TabStrip" ? undefined : pages[open];

        // The MultiPage's own client area, when the model says it: what is not client, above,
        // is the tab strip plus chrome - derived like the Frame's, and NEVER less than the strip
        // the canvas actually draws. A MultiPage's InsideHeight describes its PAGE, and the
        // difference it leaves is the borders rather than the whole tab band, so the model's
        // number put the rule through the middle of the tabs and left them hanging below the
        // line (the owner, 2026-08-16). The runtime's band measures 14 points against this 13.5,
        // which is the closer of the two by far.
        //
        // The floor is why there is no fallback here. One stood until 2026-08-16 - an honest
        // ten-point inset for a model that declines to answer - and it could never be reached:
        // it is smaller than the floor clamped over it, so the two branches had one outcome and
        // read as though they had two.
        let strung = PAGE_STRIP_HEIGHT;
        let flank = 0;
        if (row.insideWidth && row.insideHeight && row.width && row.height) {
          flank = Math.max(0, (row.width - row.insideWidth) / 2) * PT;
          strung = Math.max(PAGE_STRIP_HEIGHT, (row.height - row.insideHeight) * PT - flank);
        }

        // The rectangle starts at the BODY, for the frame's reason: the runtime draws no
        // border above or beside the tabs, and the canvas drew one all the way round - so a
        // control placed level with the tab strip looked enclosed by a box that is not there
        // (measured on the running form, 2026-08-16: nothing at all above the body's top
        // edge, which sits 14 points below the control's).
        // AT the strip rather than a pixel above it: the runtime's body edge measures 15.3pt
        // below the control's top against the 14.25pt this drew, which is the whole of the
        // MultiPage's and the TabStrip's own point of error in the parity table.
        edge.style.top = `${Math.max(0, strung)}px`;

        if (first) {
          // The body IS the page, so it wears the page's identity: a press on it selects the
          // page the way a press on the form's ground selects the form, a drop on it lands in
          // that page, and the selection has something to dress. It is a `.dc` for all three
          // of those reasons and takes no geometry from the class - a Page is placed by its
          // MultiPage, not by coordinates.
          const body = document.createElement("div");
          body.className = "dc dc-Page dc-page-body";
          body.dataset.control = first.name;
          body.dataset.kind = "Page";
          body.title = `${first.name} (Page)`;
          body.style.top = `${strung}px`;
          // ITS CHILDREN BEGIN INSIDE IT, not at its edge - see PAGE_CLIENT_INSET for the
          // measurement. A BORDER rather than padding, and the difference is the whole reason
          // this needs a note: an absolutely positioned child is placed from its containing
          // block's PADDING box, and the padding box INCLUDES the padding, so padding-top moves
          // a child not at all. The border is what the padding box starts below. Transparent,
          // because the line is already drawn by `edge` above and two would be two.
          body.style.borderTop = `${PAGE_CLIENT_INSET}px solid transparent`;
          if (flank > 0) {
            body.style.left = `${flank}px`;
            body.style.right = `${flank}px`;
            body.style.bottom = `${flank}px`;
          }

          // The open PAGE's own picture, which is the page's rather than the MultiPage's: each
          // page carries one, and the one showing is the one drawn.
          if (first.picture) {
            paintPictureSurface(body, first.picture);
          }

          box.appendChild(body);
          renderInto(body, first.name.toLowerCase());
        }
        break;
      }

      case "CheckBox":
      case "OptionButton": {
        const glyph = document.createElement("span");
        glyph.className = row.type === "CheckBox" ? "dc-glyph-box" : "dc-glyph-dot";
        box.appendChild(glyph);
        box.appendChild(document.createTextNode(caption));
        break;
      }

      case "ComboBox": {
        const arrow = document.createElement("span");
        arrow.className = "dc-drop-arrow";
        box.appendChild(arrow);
        break;
      }

      case "ScrollBar":
      case "SpinButton": {
        // The runtime draws arrow buttons at the ends; a bare rectangle read as a TextBox
        // in the owner's side-by-side. Vertical when taller than wide, the way the control
        // itself decides its axis.
        const vertical = (row.height ?? 0) >= (row.width ?? 0);
        box.classList.add(vertical ? "dc-axis-v" : "dc-axis-h");
        const startCap = document.createElement("span");
        startCap.className = "dc-arrow dc-arrow-start";
        const endCap = document.createElement("span");
        endCap.className = "dc-arrow dc-arrow-end";
        box.append(startCap, endCap);
        break;
      }

      case "Image": {
        // The picture itself since 2026-08-16, where there is one to draw - and it replaces the
        // stylesheet's crossed box, which is what "nothing to draw" looks like. Without one the
        // cross stands: the honest "an image lives here" without pretending to know it.
        if (row.picture) {
          paintPictureSurface(box, row.picture);
        }

        break;
      }

      default:
        if (caption) {
          box.appendChild(document.createTextNode(caption));
        }
        break;
    }

    // A picture that sits WITH a caption - on a button, a Label, a check box - rather than
    // being the control's whole face. The surface kinds handled their own above, because each
    // paints on a different part of itself.
    if (row.picture && !SURFACE_PICTURE_TYPES.has(row.type)) {
      dressWithPicture(box, row.picture);
    }

    // A type outside the toolbox keeps its identity visible: honest bounds, named kind.
    if (!KNOWN_TYPES.has(row.type)) {
      box.classList.add("dc-foreign");
      if (!caption) {
        box.textContent = row.type;
      }
    }

    return box;
  }

  /** Applies an orientation: the class drives the CSS (axis, order, grip), the aria and the
   * markup half's basis follow, and any dragged size is let go so the new axis starts even. */
  private setOrientation(orientation: "stacked" | "beside"): void {
    this.orientation = orientation;
    this.root.classList.toggle("stacked", orientation === "stacked");
    this.splitter.setAttribute("aria-orientation", orientation === "stacked" ? "horizontal" : "vertical");
    this.markupHalf.style.flex = "";
    this.layout();
  }

  private installSplitter(splitter: HTMLElement): void {
    let start = 0;
    let startSize = 0;

    const onMove = (event: PointerEvent) => {
      if (this.orientation === "beside") {
        const width = Math.max(160, startSize + (event.clientX - start));
        this.markupHalf.style.flex = `0 0 ${width}px`;
      } else {
        // The markup sits BELOW the divider when stacked: dragging down shrinks it.
        const height = Math.max(80, startSize - (event.clientY - start));
        this.markupHalf.style.flex = `0 0 ${height}px`;
      }
      this.layout();
    };

    splitter.addEventListener("pointerdown", (event) => {
      const box = this.markupHalf.getBoundingClientRect();
      start = this.orientation === "beside" ? event.clientX : event.clientY;
      startSize = this.orientation === "beside" ? box.width : box.height;
      splitter.setPointerCapture(event.pointerId);
      splitter.addEventListener("pointermove", onMove);
      const done = () => {
        splitter.removeEventListener("pointermove", onMove);
        splitter.removeEventListener("pointerup", done);
        splitter.removeEventListener("pointercancel", done);
      };
      splitter.addEventListener("pointerup", done);
      splitter.addEventListener("pointercancel", done);
    });

    // The switch: a double-click on the divider flips the split. No extra chrome to learn,
    // and the tooltip says so for anyone hovering the grip.
    splitter.addEventListener("dblclick", () => {
      this.setOrientation(this.orientation === "stacked" ? "beside" : "stacked");
    });

    splitter.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.setOrientation(this.orientation === "stacked" ? "beside" : "stacked");
        return;
      }

      const grow = this.orientation === "beside" ? "ArrowRight" : "ArrowUp";
      const shrink = this.orientation === "beside" ? "ArrowLeft" : "ArrowDown";
      if (event.key !== grow && event.key !== shrink) {
        return;
      }
      event.preventDefault();
      const box = this.markupHalf.getBoundingClientRect();
      const size = (this.orientation === "beside" ? box.width : box.height)
        + (event.key === grow ? 24 : -24);
      this.markupHalf.style.flex = `0 0 ${Math.max(this.orientation === "beside" ? 160 : 80, size)}px`;
      this.layout();
    });
  }

  dispose(): void {
    clearTimeout(this.lintTimer);
    this.unwatch();
    this.unwatchApplied();
    this.unwatchLint();
    this.unwatchApplySave();
    this.unwatchVocabulary();
    this.unwatchSettings();
    this.toolboxEdges.dispose();
    this.editor.dispose();
    this.model.dispose();
    this.root.remove();
  }
}

const KNOWN_TYPES = new Set([
  "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton",
  "Frame", "CommandButton", "TabStrip", "MultiPage", "Page", "ScrollBar", "SpinButton", "Image",
]);

