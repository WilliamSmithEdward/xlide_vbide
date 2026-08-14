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
 * The markup half is read-only in this first landing: an editable document whose edits go
 * nowhere is a lie, and the apply loop is the next slice. The read-only banner says exactly
 * that, so nobody wonders which half is the truth.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import type { FormMarkupApplied, FormMarkupControl, FormMarkupPayload } from "./bridge.js";
import type { DocumentId } from "./documents.js";

/** Points to CSS pixels at 96dpi: the designer's own unit, made visible at 100%. */
const PT = 4 / 3;

/** The frame's caption strip eats this much of its client area's top, approximately - an
 * honest inset rather than a measured one, until the canvas milestone measures it. */
const FRAME_INSET_TOP = 10 * PT;

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
}

export class DesignerView {
  readonly root: HTMLElement;
  readonly id: DocumentId;

  private readonly markupHost: HTMLElement;
  private readonly markupHalf: HTMLElement;
  private readonly canvasScroll: HTMLElement;
  private readonly notice: HTMLElement;
  private readonly errorStrip: HTMLElement;
  private readonly editor: monaco.editor.IStandaloneCodeEditor;
  private readonly model: monaco.editor.ITextModel;
  private readonly unwatch: () => void;
  private readonly unwatchApplied: () => void;
  private readonly deps: DesignerViewDeps;

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
    splitter.setAttribute("aria-orientation", "vertical");
    splitter.setAttribute("aria-label", "Resize markup and form");
    splitter.tabIndex = 0;

    const canvasHalf = document.createElement("div");
    canvasHalf.className = "designer-canvas-half";

    this.canvasScroll = document.createElement("div");
    this.canvasScroll.className = "designer-canvas-scroll";
    canvasHalf.appendChild(this.canvasScroll);

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

    this.model = monaco.editor.createModel("", "plaintext",
      monaco.Uri.parse(`xlide-form:/${encodeURIComponent((id.project ?? "").toLowerCase())}/${encodeURIComponent(id.module)}`));
    this.editor = monaco.editor.create(this.markupHost, {
      model: this.model,
      // The view is BUILT detached - the workspace mounts it a beat later - so the editor's
      // creation-time measure is zero. The observer picks up the real box on mount, on group
      // moves, and on the splitter, the same way the group editors track their containers.
      automaticLayout: true,
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
    });

    this.installSplitter(splitter);

    this.deps = deps;
    this.unwatch = deps.watch((payload) => this.update(payload));
    this.unwatchApplied = deps.watchApplied((outcome) => this.onApplied(outcome));
    this.request = deps.request;
  }

  private readonly request: () => void;

  /** The apply itself: the whole document, as typed, to the host. */
  applyNow(): void {
    this.errorStrip.hidden = true;
    this.deps.apply(this.model.getValue());
  }

  /** For the debug surface: set the document and apply, answering the outcome - the same
   * path Ctrl+S takes, which is the point of driving it from a suite. */
  applyDocument(markup: string): Promise<FormMarkupApplied> {
    this.model.setValue(markup);
    return new Promise((settle) => {
      this.pendingActOutcome = settle;
      this.applyNow();
    });
  }

  /** The document as it stands, for the debug surface's read side. */
  markupText(): string {
    return this.model.getValue();
  }

  private pendingActOutcome: ((outcome: FormMarkupApplied) => void) | null = null;

  private onApplied(outcome: FormMarkupApplied): void {
    this.pendingActOutcome?.(outcome);
    this.pendingActOutcome = null;

    if (outcome.ok) {
      this.errorStrip.hidden = true;
      // The fresh projection that follows is this apply's canonical print; adopt it even
      // though the dot just cleared.
      this.awaitingAdopt = true;
      if (this.dirty) {
        this.dirty = false;
        this.deps.dirtyChanged(false);
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
    this.awaitingAdopt = false;
    this.canonical = payload.markup;

    if (adopt && this.model.getValue() !== payload.markup) {
      const state = this.editor.saveViewState();
      // A single undoable edit rather than setValue: the developer's undo stack survives the
      // canonical print landing after their own apply.
      this.model.pushEditOperations([], [{
        range: this.model.getFullModelRange(),
        text: payload.markup,
      }], () => null);
      if (state) {
        this.editor.restoreViewState(state);
      }
    }

    const nowDirty = this.model.getValue() !== this.canonical;
    if (nowDirty !== this.dirty) {
      this.dirty = nowDirty;
      this.deps.dirtyChanged(nowDirty);
    }

    this.renderCanvas(payload);
  }

  private renderCanvas(payload: FormMarkupPayload): void {
    this.canvasScroll.replaceChildren();

    const form = document.createElement("div");
    form.className = "dc-form";
    form.style.width = `${Math.max(60, (payload.form?.width ?? 240)) * PT}px`;
    form.style.height = `${Math.max(40, (payload.form?.height ?? 180)) * PT}px`;

    const titlebar = document.createElement("div");
    titlebar.className = "dc-form-title";
    titlebar.textContent = payload.form?.caption ?? this.id.module;
    form.appendChild(titlebar);

    const client = document.createElement("div");
    client.className = "dc-form-client";
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

    const caption = row.caption ?? "";

    switch (row.type) {
      case "Frame": {
        const legend = document.createElement("div");
        legend.className = "dc-frame-caption";
        legend.textContent = caption;
        box.appendChild(legend);
        const inner = document.createElement("div");
        inner.className = "dc-frame-client";
        inner.style.top = `${FRAME_INSET_TOP}px`;
        box.appendChild(inner);
        renderInto(inner, row.name.toLowerCase());
        break;
      }

      case "MultiPage":
      case "TabStrip": {
        const strip = document.createElement("div");
        strip.className = "dc-page-strip";
        const pages = byParent.get(row.name.toLowerCase()) ?? [];
        for (const [index, page] of pages.entries()) {
          const tab = document.createElement("span");
          tab.className = "dc-page-tab" + (index === 0 ? " first" : "");
          tab.textContent = page.caption ?? page.name;
          strip.appendChild(tab);
        }
        box.appendChild(strip);

        // The FIRST page's content shows, the way the control itself opens; the others are
        // headers only until selection lands with the canvas milestone.
        const first = pages[0];
        if (first) {
          const body = document.createElement("div");
          body.className = "dc-page-body";
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

  private installSplitter(splitter: HTMLElement): void {
    let startX = 0;
    let startWidth = 0;

    const onMove = (event: PointerEvent) => {
      const width = Math.max(160, startWidth + (event.clientX - startX));
      this.markupHalf.style.flex = `0 0 ${width}px`;
      this.layout();
    };

    splitter.addEventListener("pointerdown", (event) => {
      startX = event.clientX;
      startWidth = this.markupHalf.getBoundingClientRect().width;
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

    splitter.addEventListener("keydown", (event) => {
      if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") {
        return;
      }
      event.preventDefault();
      const width = this.markupHalf.getBoundingClientRect().width
        + (event.key === "ArrowRight" ? 24 : -24);
      this.markupHalf.style.flex = `0 0 ${Math.max(160, width)}px`;
      this.layout();
    });
  }

  dispose(): void {
    this.unwatch();
    this.unwatchApplied();
    this.editor.dispose();
    this.model.dispose();
    this.root.remove();
  }
}

const KNOWN_TYPES = new Set([
  "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton",
  "Frame", "CommandButton", "TabStrip", "MultiPage", "Page", "ScrollBar", "SpinButton", "Image",
]);
