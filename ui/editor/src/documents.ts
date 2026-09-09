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

/** One open document's identity: the module, its workbook display name when known, and the
 * FACE when it is not the code pane - a form's designer tab is the same module worn a second
 * way, and the two are two tabs. `design` is the designer; `history:<short>` is the module's
 * text at a commit, read-only, one tab per commit opened. */
export interface DocumentId {
  module: string;
  project: string | null;
  face?: string;
}

/** The face prefix of a past-version tab; the short hash follows it. */
export const HISTORY_FACE_PREFIX = "history:";

/** Whether an id is a form's designer tab. */
export function isDesignFace(id: { face?: string | null | undefined }): boolean {
  return id.face === "design";
}

/** Whether an id is a module's past-version tab. */
export function isHistoryFace(id: { face?: string | null | undefined }): boolean {
  return typeof id.face === "string" && id.face.startsWith(HISTORY_FACE_PREFIX);
}

/** The face a past-version tab wears for a commit. */
export function historyFaceOf(short: string): string {
  return `${HISTORY_FACE_PREFIX}${short}`;
}

/** The short hash a history face names, or "" for any other face. */
export function historyShortOf(face: string | null | undefined): string {
  return typeof face === "string" && face.startsWith(HISTORY_FACE_PREFIX)
    ? face.slice(HISTORY_FACE_PREFIX.length)
    : "";
}

/**
 * The face a host or a tab names, or undefined for the code pane.
 *
 * ONLY THE FACES THIS PAGE KNOWS. The host publishes null or "code" for a mirrored pane, and a
 * dataset attribute reads back "" when absent; both are the code identity. A face this page has
 * never heard of would otherwise become a tab nothing can show and nothing can close.
 */
export function knownFace(face: string | null | undefined): string | undefined {
  const id = { face };
  return isDesignFace(id) || isHistoryFace(id) ? face ?? undefined : undefined;
}

/** The identity two documents are the same by. Case-insensitive, the way the host compares. */
export function docKeyOf(module: string, project: string | null | undefined, face?: string | null): string {
  return `${(project ?? "").toLowerCase()}\0${module.toLowerCase()}${face ? `\0${face}` : ""}`;
}

/**
 * Whether a caller's spelling of a workbook names the one a tab is holding.
 *
 * TWO SPELLINGS, BOTH OURS. `projects()` answers a `projectId` - the workbook's full path - and
 * a `project` - the name the tree draws. The host's routes take either, deliberately: its own
 * comment says an identity is accepted "because this product hands them out and then would not"
 * take them back. The PAGE's actions compared the caller's string to a tab's display name and
 * nothing else, so the field literally named `projectId` was refused by `activate` while its
 * refusal listed the module as open - which reads as the surface contradicting itself, and is
 * how it was found (2026-08-22).
 *
 * A tab carries the display name, so a path matches when its file name does. An unsaved
 * workbook has no path and falls through to the exact compare, which is the right answer for it.
 */
export function namesTheSameWorkbook(held: string | null | undefined, asked: string): boolean {
  const mine = (held ?? "").toLowerCase();
  const theirs = asked.toLowerCase();
  if (mine === theirs) {
    return true;
  }

  return fileNameOf(theirs) === mine || fileNameOf(mine) === theirs;
}

const fileNameOf = (path: string): string =>
  path.slice(Math.max(path.lastIndexOf("\\"), path.lastIndexOf("/")) + 1);

/** The model URI for a document. Both parts encoded, so names with slashes cannot forge paths. */
export function docUriOf(module: string, project: string | null | undefined): monaco.Uri {
  return monaco.Uri.parse(
    `xlide:/${encodeURIComponent((project ?? "").toLowerCase())}/${encodeURIComponent(module)}`);
}

/** The scheme a past-version tab's model wears, so nothing mistakes it for a live document. */
export const HISTORY_URI_SCHEME = "xlide-history";

/**
 * The model URI for a module's text at a commit. A DIFFERENT scheme from the live documents',
 * and a path whose LAST segment is the hash rather than the module: the change highlighter
 * finds a module's model by the last path segment, and this store answers null for a model it
 * does not own, so a history model is painted by nobody and announced as nobody's.
 */
export function historyUriOf(module: string, project: string | null | undefined, short: string): monaco.Uri {
  return monaco.Uri.parse(
    `${HISTORY_URI_SCHEME}:/${encodeURIComponent((project ?? "").toLowerCase())}`
    + `/${encodeURIComponent(module)}/${encodeURIComponent(short)}`);
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
