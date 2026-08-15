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
import type { FormMarkupApplied, FormMarkupControl, FormMarkupDraft, FormMarkupLintFinding, FormMarkupPayload } from "./bridge.js";
import type { DocumentId } from "./documents.js";

/** Points to CSS pixels at 96dpi: the designer's own unit, made visible at 100%. */
const PT = 4 / 3;

/** The markup's own language id, for the tab's document. */
const FORM_MARKUP_LANGUAGE = "xlide-form";

/**
 * The markup's grammar, once per page. STANDARD token names on purpose - keyword, type,
 * string, number, identifier - so the existing themes colour it with no theme edits; and a
 * page-side grammar is for PAINT ONLY. The language's truth is Core's parser: linting will
 * be a tolerant parse host-side (docs/userform-designer.md, the language service), never a
 * second grammar here that can drift.
 */
let markupLanguageRegistered = false;
function registerMarkupLanguage(): void {
  if (markupLanguageRegistered) {
    return;
  }
  markupLanguageRegistered = true;

  monaco.languages.register({ id: FORM_MARKUP_LANGUAGE });

  monaco.languages.setMonarchTokensProvider(FORM_MARKUP_LANGUAGE, {
    // The toolbox kinds plus the two structural words; a type outside this list still reads
    // as an identifier, which is honest - the apply treats it as foreign too.
    controlKinds: [
      "Form", "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton",
      "ToggleButton", "Frame", "CommandButton", "TabStrip", "MultiPage", "Page",
      "ScrollBar", "SpinButton", "Image",
    ],
    tokenizer: {
      root: [
        [/"(?:[^"]|"")*"/, "string"],
        [/\b(?:at|size)\b/, "keyword"],
        // The size pair is ONE value - "360x320.25" - or its x paints as an identifier.
        [/-?\d+(?:\.\d+)?x-?\d+(?:\.\d+)?/, "number"],
        [/-?\d+(?:\.\d+)?/, "number"],
        [/[A-Za-z_][\w.]*(?=\s*=)/, "attribute.name"],
        [/[A-Za-z_][\w.]*/, {
          cases: {
            "@controlKinds": "type",
            "@default": "identifier",
          },
        }],
        [/[x,=]/, "delimiter"],
      ],
    },
  });

  monaco.languages.setLanguageConfiguration(FORM_MARKUP_LANGUAGE, {
    // A container line opens a level, the way the printer indents its children.
    onEnterRules: [{
      beforeText: /^\s*(?:Frame|MultiPage|Page)\b.*$/,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    }],
    brackets: [],
    autoClosingPairs: [{ open: '"', close: '"' }],
  });
}

/** The frame's caption strip eats this much of its client area's top, approximately - an
 * honest inset rather than a measured one, until the canvas milestone measures it. */
const FRAME_INSET_TOP = 10 * PT;

/** How far a press must travel before it is a DRAG rather than a click, in CSS pixels. Below
 * it a press is a selection and nothing moves, which is what keeps a click a click. */
const DRAG_THRESHOLD = 3;

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
  head: string;
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
}

/** The smallest a gesture may make something, in points. A control can be tiny - the renderer
 * floors its own boxes at four - but a form small enough to lose its title bar is a mistake
 * rather than a design. */
const MIN_CONTROL = 4;
const MIN_FORM = 24;

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
  /** Save the workbook, the host's own File Save - the second half of the designer's
   * Ctrl+S, after a successful apply. */
  saveWorkbook(): void;
  /** Subscribe to the HOST's Ctrl+S: the accelerator never reaches the page, so the host
   * asks the tab to apply-then-save through this; returns the unwatch. */
  watchApplySave(listener: () => void): () => void;
  /** A canvas double-click: the host writes or shows the control's default event handler.
   * Null means the form itself. */
  eventStub(control: string | null): void;
  /** The selection changed: the Properties panel follows. Null is the form itself. */
  selection(control: string | null): void;
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
  private lintTimer: ReturnType<typeof setTimeout> | undefined;
  private readonly deps: DesignerViewDeps;

  /** The last APPLIED projection - the form's own truth, and the wardrobe a draft preview
   * borrows its display extras from. */
  private lastPayload: FormMarkupPayload | null = null;

  /** Whether the canvas is currently showing the DOCUMENT's draft rather than the form. */
  private draftShown = false;

  /** The selected control's name, "" for the FORM itself, null for no selection. */
  private selectedName: string | null = null;

  /** The press in flight, once it has grabbed something movable. */
  private drag: CanvasDrag | null = null;

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
    this.canvasScroll.addEventListener("dblclick", (event) => {
      const control = (event.target as HTMLElement).closest<HTMLElement>(".dc");
      if (control?.dataset.control) {
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

    this.model.onDidChangeContent(() => {
      const nowDirty = this.canonical !== null && this.model.getValue() !== this.canonical;
      if (nowDirty !== this.dirty) {
        this.dirty = nowDirty;
        deps.dirtyChanged(nowDirty);
      }

      // The squiggles follow the typing, debounced past the keystroke rate: the lint is a
      // host round trip, and a request per keystroke asks the same question mid-word.
      clearTimeout(this.lintTimer);
      this.lintTimer = setTimeout(() => deps.lint(this.model.getValue()), 350);
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
    this.unwatchApplySave = deps.watchApplySave(() => this.applyNow());
    this.request = deps.request;

    // LAST, because it lays the editor out and the editor must exist: the first version ran
    // this beside the DOM build and the constructor died reaching for an editor not yet
    // created - a designer tab that stood in the strip and mounted nothing.
    this.setOrientation("stacked");
    this.refreshCollapseTitles();
  }

  private readonly request: () => void;

  /** Set when an apply should be followed by the host's save - every Ctrl+S is, because
   * Ctrl+S means "save the workbook" everywhere else in the product and the designer must
   * not quietly mean less (the owner, 2026-08-15: "CTRL+S in designer isn't saving"). */
  private saveAfterApply = false;

  /** Ctrl+S: the document to the form, then the workbook to disk. A clean document has
   * nothing to apply and just saves. */
  applyNow(): void {
    this.errorStrip.hidden = true;
    if (!this.dirty) {
      this.deps.saveWorkbook();
      return;
    }

    this.saveAfterApply = true;
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
      this.saveAfterApply = true;
      this.deps.apply(this.model.getValue());
    });
  }

  /** The document as it stands, for the debug surface's read side. */
  markupText(): string {
    return this.model.getValue();
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

  private onApplied(outcome: FormMarkupApplied): void {
    this.pendingActOutcome?.(outcome);
    this.pendingActOutcome = null;

    const saveNext = this.saveAfterApply;
    this.saveAfterApply = false;

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
      // form the developer was just told did not take their document.
      if (saveNext) {
        this.deps.saveWorkbook();
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
    if (payload.markup === null) {
      this.notice.textContent = payload.reason ?? "the form could not be read";
      this.notice.hidden = false;
      return;
    }

    this.notice.hidden = true;

    const adopt = this.awaitingAdopt || !this.dirty;
    const first = this.canonical === null;
    this.awaitingAdopt = false;
    this.canonical = payload.markup;

    if (adopt && this.model.getValue() !== payload.markup) {
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
    }

    const nowDirty = this.model.getValue() !== this.canonical;
    if (nowDirty !== this.dirty) {
      this.dirty = nowDirty;
      this.deps.dirtyChanged(nowDirty);
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
          fontName: worn.fontName ?? null,
          fontSize: worn.fontSize ?? null,
          fontBold: worn.fontBold ?? null,
          fontItalic: worn.fontItalic ?? null,
          backColor: worn.backColor ?? null,
          foreColor: worn.foreColor ?? null,
          insideWidth: worn.insideWidth ?? null,
          insideHeight: worn.insideHeight ?? null,
          tabs: worn.tabs ?? null,
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
      }
      : appliedForm;

    return { markup: null, reason: null, form, controls };
  }

  /** The canvas as rendered, for the harness: which picture stands (draft or applied), whether
   * the document holds unapplied edits, every control's name with its placed position in
   * POINTS - the document's own unit, so a row can compare the two - what is selected, and
   * where the markup caret sits. */
  canvasSnapshot(): {
    draft: boolean;
    dirty: boolean;
    undoable: boolean;
    selected: string | null;
    markupLine: number;
    controls: { name: string; left: number; top: number; width: number; height: number }[];
  } {
    return {
      draft: this.draftShown,
      dirty: this.dirty,
      // Whether the DOCUMENT has a gesture to give back. The canvas's own Ctrl+Z is only as
      // good as this, and a row that waits for an undo to land wants to know the difference
      // between "undone and re-rendered" and "there was nothing on the stack".
      undoable: this.model.canUndo(),
      selected: this.selectedName,
      markupLine: this.editor.getPosition()?.lineNumber ?? 0,
      controls: [...this.canvasScroll.querySelectorAll<HTMLElement>(".dc")].map((el) => ({
        name: el.dataset.control ?? "",
        ...this.inPoints({
          left: Number.parseFloat(el.style.left || "0"),
          top: Number.parseFloat(el.style.top || "0"),
          width: el.offsetWidth,
          height: el.offsetHeight,
        }),
      })),
    };
  }

  /** Selects a control by name - "" is the FORM itself - dresses it in the native handles,
   * and puts the markup caret on its line: the two halves pointing at one thing. The same
   * entry a canvas click takes, so the act and the click stay one gesture. */
  select(name: string): void {
    this.selectedName = name;
    this.dressSelection();
    this.revealInMarkup(name);
    // The Properties panel follows the selection host-side - M4's bridgehead.
    this.deps.selection(name === "" ? null : name);
  }

  /** A canvas double-click, for the debug surface: the host's event-stub gesture. */
  requestEventStub(control: string | null): void {
    this.deps.eventStub(control);
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

    const control = target.closest<HTMLElement>(".dc");
    if (!control?.dataset.control) {
      if (target.closest(".dc-form")) {
        this.select("");
      }
      return;
    }

    this.select(control.dataset.control);

    // A Page is not placed by coordinates - it fills its MultiPage - so there is nothing to
    // drag it by. Everything else only ARMS here: the gesture begins past the threshold, which
    // is what keeps a plain click a plain click.
    if (control.dataset.kind === "Page") {
      return;
    }

    this.arm(control.dataset.control, null, event);
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

    this.paintBox(drag.element, this.proposal(drag, dx / PT, dy / PT));
    event.preventDefault();
  }

  /** The drop writes the document: the box the gesture ended on, in points, as one edit. */
  private onCanvasDrop(event: PointerEvent): void {
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
    const box = this.proposal(drag, (event.clientX - drag.startX) / PT, (event.clientY - drag.startY) / PT);
    if (!this.writeBox(drag.name, box, drag.edge !== null)) {
      this.paintBox(drag.element, drag.origin);
    }
  }

  /**
   * What a gesture is proposing, in points: the origin box with the pointer's travel applied -
   * moved whole for a drag, or with the grabbed edges pushed for a resize - then clamped so
   * nothing leaves its parent or collapses to nothing. Measured from the PRESS every frame, so
   * a slow drag cannot accumulate rounding, and rounded to whole points because that is what
   * the document will carry.
   */
  private proposal(drag: CanvasDrag, dx: number, dy: number): Box {
    const origin = drag.origin;
    const room = this.roomFor(drag.element);
    const floor = drag.name === "" ? MIN_FORM : MIN_CONTROL;

    if (drag.edge === null) {
      return this.clamp({ ...origin, left: origin.left + dx, top: origin.top + dy }, room, floor);
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

    return this.clamp(box, room, floor);
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
    const box = event.shiftKey
      ? { ...base, width: base.width + step.dx, height: base.height + step.dy }
      : { ...base, left: base.left + step.dx, top: base.top + step.dy };
    this.writeBox(this.selectedName, this.clamp(box, room, floor), event.shiftKey);
  }

  /**
   * Writes a box into the document: the thing's own line rewritten with the position and size
   * it now has, rounded to whole points, as ONE undoable edit, and painted at once so the
   * gesture lands before the round trip answers.
   *
   * This is the single commit every canvas gesture takes - drag, resize, nudge, and the
   * harness's acts - and it writes the DOCUMENT, never the form: the draft preview shows it
   * immediately, the dot says it has not reached the form yet, and Ctrl+S is still the only
   * apply. Whole points because a hand-placed control wants round numbers; the 6-point grid the
   * native designer snaps to is M6's. The FORM's line takes a size and never a position.
   */
  private writeBox(name: string, box: Box, sized: boolean): boolean {
    const element = this.elementOf(name);
    const header = this.headerOf(name);
    if (!element || !header) {
      return false;
    }

    const left = Math.round(box.left);
    const top = Math.round(box.top);
    const width = Math.round(box.width);
    const height = Math.round(box.height);
    // A line that spelled no size keeps spelling none unless this gesture is a resize: a move
    // must not quietly pin a size the developer left to the control.
    const spellSize = sized || header.size !== null;
    const text = header.head
      + (header.form ? "" : ` at ${left},${top}`)
      + (spellSize ? ` size ${width}x${height}` : "");

    // A stack stop BEFORE the edit and none after. Monaco appends edits to whichever element is
    // open, so the stop is what keeps one gesture from joining the gesture before it; leaving
    // the element open afterwards is equally deliberate, because the canonical print that
    // follows a save belongs to this gesture - with it riding along, one Ctrl+Z after a save
    // reaches the text from before the move rather than a step that only differs by the
    // machine's own rounding (both halves measured live, 2026-08-15).
    this.model.pushStackElement();
    this.model.pushEditOperations([], [{
      range: new monaco.Range(header.line, 1, header.line, this.model.getLineMaxColumn(header.line)),
      text,
    }], () => null);
    this.paintBox(element, { left, top, width, height });
    return true;
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
    const name = this.selectedName;
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
    return true;
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
   * The document line that declares a thing - a control by name, the FORM for "" - split at its
   * geometry. Null when no line names it, or when the line is not in a shape this can rewrite -
   * a half-typed header, a caption whose quote never closes - because a gesture that cannot
   * write the document must not pretend to have changed anything.
   *
   * The caption is skipped by scanning rather than by regex: ` at ` inside a caption is
   * ordinary text, and only a scan that knows about doubled quotes can tell the two apart.
   */
  private headerOf(name: string): HeaderLine | null {
    const form = name === "";
    const line = form ? 1 : this.headerLineOf(name);
    if (line === 0 || line > this.model.getLineCount()) {
      return null;
    }

    const text = this.model.getLineContent(line);
    const named = form
      ? /^Form\s+\S+/.exec(text)
      : new RegExp(`^\\s+\\S+\\s+${escapeForRegExp(name)}(?=\\s|$)`).exec(text);
    if (!named) {
      return null;
    }

    let at = named[0].length;
    if (/^\s+"/.test(text.slice(at))) {
      let scan = at + text.slice(at).indexOf('"') + 1;
      let closed = false;
      while (scan < text.length) {
        if (text[scan] !== '"') {
          scan += 1;
        } else if (text[scan + 1] === '"') {
          scan += 2;
        } else {
          scan += 1;
          closed = true;
          break;
        }
      }

      if (!closed) {
        return null;
      }

      at = scan;
    }

    // What follows the caption is geometry and nothing else, so a plain match is safe here.
    const tail = /^(?:\s+at\s+(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?))?(?:\s+size\s+(-?\d+(?:\.\d+)?)\s*x\s*(-?\d+(?:\.\d+)?))?\s*$/i
      .exec(text.slice(at));
    if (!tail) {
      return null;
    }

    const number = (spelled: string | undefined): number | null =>
      spelled === undefined ? null : Number.parseFloat(spelled);
    const left = number(tail[1]);
    const top = number(tail[2]);
    const width = number(tail[3]);
    const height = number(tail[4]);

    return {
      line,
      head: text.slice(0, at),
      at: left === null || top === null ? null : { left, top },
      size: width === null || height === null ? null : { width, height },
      form,
    };
  }

  /** The 1-based line that declares a control, or 0 when the document names it nowhere. */
  private headerLineOf(name: string): number {
    const needle = new RegExp(`^\\s+\\S+\\s+${escapeForRegExp(name)}(\\s|$)`);
    return this.model.getLinesContent().findIndex((text) => needle.test(text)) + 1;
  }

  /**
   * A drag driven from the debug surface: the REAL pointer sequence, on the element a mouse
   * would actually hit - elementFromPoint at a point inside the control - so the act proves
   * hit-testing, the threshold and the commit rather than the arithmetic alone. Deltas are in
   * POINTS, the designer's unit and the document's. Returns what happened, for the row.
   */
  dragControl(name: string, dx: number, dy: number): string {
    const element = this.elementFor(name);
    if (!element) {
      return `no control named ${name} on the canvas`;
    }

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
    this.sendGesture(landed, grabbed, dx, dy);
    return `dragged ${name}`;
  }

  /**
   * A resize driven from the debug surface: the named thing is SELECTED first, the way a hand
   * must, and then the real pointer sequence grabs its handle - `nw`, `n`, `ne`, `e`, `se`,
   * `s`, `sw`, `w` - and pulls it by a delta in POINTS. "" is the form's own frame. The press
   * goes through the hit test at the handle, so a handle nothing can reach fails the act.
   */
  resizeControl(name: string, edge: string, dx: number, dy: number): string {
    if (this.selectedName !== name) {
      this.select(name);
    }

    const handle = this.canvasScroll.querySelector<HTMLElement>(`.dc-handle-${CSS.escape(edge)}`);
    if (!handle) {
      return `no ${edge} handle stands on ${name === "" ? "the form" : name}`;
    }

    const box = handle.getBoundingClientRect();
    const spot = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
    const reached = document.elementFromPoint(spot.x, spot.y)?.closest<HTMLElement>(".dc-handle");
    if (reached?.dataset.edge !== edge) {
      return `the ${edge} handle is not what a press on it reaches - ${reached?.className || "nothing"} is`;
    }

    this.sendGesture(reached, spot, dx, dy);
    return `resized ${name === "" ? "the form" : name}`;
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

  /** The pointer sequence both instruments send: press where the hit test answered, one move
   * past the threshold, one move the whole way, and the release - all but the press delivered
   * to the canvas, which is where a captured pointer's events go. */
  private sendGesture(target: EventTarget, from: { x: number; y: number }, dx: number, dy: number): void {
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
      }));
    };

    send(target, "pointerdown", from.x, from.y);
    send(this.canvasScroll, "pointermove", from.x + DRAG_THRESHOLD + 1, from.y);
    send(this.canvasScroll, "pointermove", from.x + dx * PT, from.y + dy * PT);
    send(this.canvasScroll, "pointerup", from.x + dx * PT, from.y + dy * PT);
  }

  private dressSelection(): void {
    this.canvasScroll.querySelector(".dc-selection")?.remove();
    this.canvasScroll.querySelector(".dc-selected")?.classList.remove("dc-selected");

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

    for (const spot of ["nw", "n", "ne", "e", "se", "s", "sw", "w"]) {
      const handle = document.createElement("span");
      handle.className = `dc-handle dc-handle-${spot}`;
      // The edge rides the element: a press reads which way to pull from what it grabbed.
      handle.dataset.edge = spot;
      overlay.appendChild(handle);
    }

    if (name === "") {
      // The form does not clip, so its overlay lives inside it, over the whole face.
      overlay.classList.add("dc-selection-form");
      target.appendChild(overlay);
    } else {
      overlay.style.left = target.style.left;
      overlay.style.top = target.style.top;
      overlay.style.width = target.style.width;
      overlay.style.height = target.style.height;
      target.insertAdjacentElement("afterend", overlay);

      // A SELECTED control offers the move, and says so with the cursor (the owner's rule,
      // 2026-08-15). Only the selected one, and only a control that has a position to change:
      // the form is where the canvas puts it, and a Page fills its MultiPage.
      if (target.dataset.kind !== "Page") {
        target.classList.add("dc-selected");
      }
    }
  }

  /** The markup caret onto the selected thing's line: the form's is line 1, a control's is
   * the line that names it. */
  private revealInMarkup(name: string): void {
    let line = 1;
    if (name !== "") {
      line = this.headerLineOf(name);
      if (line === 0) {
        return;
      }
    }

    this.editor.setPosition({ lineNumber: line, column: this.model.getLineMaxColumn(line) });
    this.editor.revealLineInCenterIfOutsideViewport(line);
  }

  private renderCanvas(payload: FormMarkupPayload): void {
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

    this.canvasScroll.appendChild(form);

    // A re-render replaces every element, so the selection is dressed again onto the new
    // ones; a selected control the new picture no longer holds simply loses its handles.
    this.dressSelection();
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
        const legend = document.createElement("div");
        legend.className = "dc-frame-caption";
        legend.textContent = caption;
        box.appendChild(legend);
        const inner = document.createElement("div");
        inner.className = "dc-frame-client";
        // The REAL client area when the model says it: side borders split the width
        // difference evenly, and what remains of the height difference above the client is
        // the caption strip - derived, not guessed, which is the parity rule.
        if (row.insideWidth && row.insideHeight && row.width && row.height) {
          const side = Math.max(0, (row.width - row.insideWidth) / 2) * PT;
          inner.style.left = `${side}px`;
          inner.style.right = `${side}px`;
          inner.style.bottom = `${side}px`;
          inner.style.top = `${Math.max(0, (row.height - row.insideHeight) * PT - side)}px`;
        } else {
          inner.style.top = `${FRAME_INSET_TOP}px`;
        }
        box.appendChild(inner);
        renderInto(inner, row.name.toLowerCase());
        break;
      }

      case "MultiPage":
      case "TabStrip": {
        const strip = document.createElement("div");
        strip.className = "dc-page-strip";
        const pages = byParent.get(row.name.toLowerCase()) ?? [];
        // A MultiPage's headers are its Page children; a TabStrip's are its own Tabs, which
        // are not controls and ride the row instead - without them the strip drew as a bare
        // box (the owner's side-by-side, 2026-08-13).
        const headers = row.type === "TabStrip" && row.tabs?.length
          ? row.tabs
          : pages.map((page) => page.caption ?? page.name);
        for (const [index, header] of headers.entries()) {
          const tab = document.createElement("span");
          tab.className = "dc-page-tab" + (index === 0 ? " first" : "");
          tab.textContent = header;
          strip.appendChild(tab);
        }
        box.appendChild(strip);

        // The FIRST page's content shows, the way the control itself opens; the others are
        // headers only until selection lands with the canvas milestone.
        const first = pages[0];
        if (first) {
          const body = document.createElement("div");
          body.className = "dc-page-body";
          // The MultiPage's own client area, when the model says it: what is not client,
          // above, is the tab strip plus chrome - derived like the Frame's.
          if (row.insideWidth && row.insideHeight && row.width && row.height) {
            const side = Math.max(0, (row.width - row.insideWidth) / 2) * PT;
            body.style.left = `${side}px`;
            body.style.right = `${side}px`;
            body.style.bottom = `${side}px`;
            body.style.top = `${Math.max(0, (row.height - row.insideHeight) * PT - side)}px`;
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
        // A crossed box: the honest "an image lives here" without pretending to know it.
        break;
      }

      default:
        if (caption) {
          box.appendChild(document.createTextNode(caption));
        }
        break;
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
    this.editor.dispose();
    this.model.dispose();
    this.root.remove();
  }
}

const KNOWN_TYPES = new Set([
  "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton",
  "Frame", "CommandButton", "TabStrip", "MultiPage", "Page", "ScrollBar", "SpinButton", "Image",
]);
