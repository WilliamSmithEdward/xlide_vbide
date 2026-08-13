/*
 * The open documents, as live Monaco models - one per (workbook, module), for as long as the
 * module's pane is open (decision 12).
 *
 * Identity is the pair, never the bare name: two workbooks holding a Module1 are two documents,
 * and the model URI carries both parts so they cannot collide. WHICH documents are open is the
 * host's truth, published with the tabs; this store follows that list. Which model an editor
 * shows is the workspace's business, not this store's.
 *
 * Undo stacks, markers, decorations, and view state all hang off the model, so keeping models
 * alive is what makes switching tabs free and background modules first-class: their squiggles
 * update, their bookmarks stay, their undo history survives.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import { VBA_LANGUAGE_ID } from "./vba.js";

/** One open document's identity: the module, and its workbook display name when known. */
export interface DocumentId {
  module: string;
  project: string | null;
}

/** The identity two documents are the same by. Case-insensitive, the way the host compares. */
export function docKeyOf(module: string, project: string | null | undefined): string {
  return `${(project ?? "").toLowerCase()}\0${module.toLowerCase()}`;
}

/** The model URI for a document. Both parts encoded, so names with slashes cannot forge paths. */
export function docUriOf(module: string, project: string | null | undefined): monaco.Uri {
  return monaco.Uri.parse(
    `xlide:/${encodeURIComponent((project ?? "").toLowerCase())}/${encodeURIComponent(module)}`);
}

interface Entry {
  id: DocumentId;
  model: monaco.editor.ITextModel;
}


export class DocumentStore {
  private readonly entries = new Map<string, Entry>();

  /** Called with each newly created model, so listeners attach exactly once per model. */
  onModelCreated: ((id: DocumentId, model: monaco.editor.ITextModel) => void) | null = null;

  /** Called just before a model is disposed, so holders can let go of it first. */
  onModelClosing: ((id: DocumentId, model: monaco.editor.ITextModel) => void) | null = null;

  /**
   * Opens a document, idempotently: a model that already exists adopts the text in place when
   * it differs (the host re-opens everything after a page reload, and re-sends a clean
   * document whose module changed underneath), and a new one is created with it.
   *
   * The adopt returns true when an in-place edit was applied, so the caller can wrap it in
   * its echo suppression.
   */
  open(module: string, project: string | null, text: string, adopt: (model: monaco.editor.ITextModel, text: string) => void): monaco.editor.ITextModel {
    const key = docKeyOf(module, project);
    const existing = this.entries.get(key);
    if (existing) {
      if (existing.model.getValue() !== text) {
        adopt(existing.model, text);
      }
      return existing.model;
    }

    const model = monaco.editor.createModel(text, VBA_LANGUAGE_ID, docUriOf(module, project));
    const entry: Entry = { id: { module, project }, model };
    this.entries.set(key, entry);
    this.onModelCreated?.(entry.id, model);
    return model;
  }

  /** The model for a document, or null when it is not open. */
  get(module: string, project: string | null | undefined): monaco.editor.ITextModel | null {
    return this.entries.get(docKeyOf(module, project))?.model ?? null;
  }

  /** The identity behind a model, or null for a model this store does not own. */
  idOf(model: monaco.editor.ITextModel): DocumentId | null {
    for (const entry of this.entries.values()) {
      if (entry.model === model) {
        return entry.id;
      }
    }
    return null;
  }

  /** Every open document, in opening order. */
  all(): DocumentId[] {
    return [...this.entries.values()].map((entry) => entry.id);
  }

  /**
   * Disposes every document that is not in the open list. The host publishes which modules are
   * open with the tabs; models follow that truth, so a closed pane's model goes here and its
   * undo history honestly dies with it.
   */
  closeMissing(open: DocumentId[]): DocumentId[] {
    const keep = new Set(open.map((id) => docKeyOf(id.module, id.project)));
    const closed: DocumentId[] = [];

    for (const [key, entry] of [...this.entries]) {
      if (!keep.has(key)) {
        closed.push(entry.id);
        this.entries.delete(key);
        this.onModelClosing?.(entry.id, entry.model);
        entry.model.dispose();
      }
    }

    return closed;
  }

  /** Disposes everything: the empty workspace. */
  clear(): void {
    for (const entry of [...this.entries.values()]) {
      this.entries.delete(docKeyOf(entry.id.module, entry.id.project));
      this.onModelClosing?.(entry.id, entry.model);
      entry.model.dispose();
    }
  }
}
