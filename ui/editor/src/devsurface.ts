/*
 * The page's own answer to "what does the surface look like" and "do this to it".
 *
 * Everything here exists because the alternative was a bespoke DOM script per question, written
 * fresh each time by whoever was debugging. That is the most expensive habit in this repo's short
 * history: three of the findings in lessons-2026-08-07.md are probes that measured themselves, and
 * every one of them was a hand-rolled querySelectorAll. Two more went the other way and drove the
 * page by synthesising events that its handlers do not listen for, then reported a working feature
 * broken. The tab close box is the standing example: it arms at pointerdown and fires at pointerup,
 * so `element.click()` on it does nothing at all, silently.
 *
 * So the page reports its own state, from the fields the render reads rather than from the rendered
 * rows, and the page performs its own actions, through the same methods a click reaches. A probe
 * that asks `xlideUi.state()` cannot be looking at a stale render and calling it the state, and one
 * that calls `xlideUi.act("closeActive")` cannot miss because it guessed the wrong event name.
 *
 * Reached from the xlide api's `ui` and `act` routes, and from a devtools console. Shipped in every
 * build: it is a few hundred bytes of read-only reporting over objects the page already holds, and
 * a door that is only there in Debug is a door nobody trusts in the build that matters.
 */

import type * as monaco from "monaco-editor";
// The runtime object, not only its types: getModelMarkers is what draws the squiggles, and a
// type-only import cannot call it.
import * as monacoApi from "monaco-editor/editor/editor.api.js";

import { openTabOrderDialog } from "./taborderdialog.js";
import type { EditorBridge } from "./bridge.js";
import { colourPickerState, pickColour } from "./colourpicker.js";
import type { Explorer, ExplorerSnapshot } from "./explorer.js";
import type { Workspace, WorkspaceSnapshot } from "./workspace.js";
import { currentSettings } from "./settings.js";
import { changesPaneProbe } from "./changespane.js";
import { agentDialogProbe } from "./agentdialog.js";
import { syncDialogProbe } from "./syncdialog.js";

/** What is standing in front of the page, since a modal swallows every key sent at the surface. */
interface DialogSnapshot {
  id: string;
  title: string;
}

export interface UiSnapshot {
  workspace: WorkspaceSnapshot;
  explorer: ExplorerSnapshot;
  /** Panes and whether each is showing. */
  panes: { name: string; title: string; open: boolean }[];
  /** Page-side dialogs: settings, help, sponsor, references, object browser. */
  dialogs: DialogSnapshot[];
  /** What has not arrived yet, so a blank surface can be told from a slow one. */
  waiting: { documents: string[]; hostActive: unknown; stoppedRows: number };
  /**
   * The editor the keyboard would reach, and what it holds.
   *
   * `host` is a DIFFERENT editor whenever the active group is not the one showing the
   * host-active module, which a split makes ordinary. Anything driven through the host lands
   * there, so anything measuring the host has to read there: a caret trip that set the host's
   * caret and read the active group's waited four seconds for two editors that were never
   * going to agree (2026-08-07). Null when no group is showing the host-active module.
   */
  focus: {
    model: string | null;
    hasFocus: boolean;
    line: number;
    column: number;
    host: { model: string; line: number; column: number } | null;
  };
  settings: Record<string, unknown>;
  /** Whether the empty view is up, which is a different thing from having no tabs. */
  emptyViewShown: boolean;
  /** The Properties panel: the component it is showing, its kind, and every row it draws. */
  properties: {
    component: string;
    kind: string;
    rows: { name: string; value: string; writable: boolean; boolean: boolean; options?: string[] | null; swatch?: string | null; picture?: boolean; previewBytes?: number }[];
  };
  /**
   * What the status line is saying right now, empty when it is saying nothing.
   *
   * The status bar was in no snapshot at all, so a whole class of behaviour - an action declined
   * with an explanation - could be driven and could not be observed. The condition that forced it
   * is the engine's cold start, which now holds a notice for the seconds before language features
   * work, and a held notice that fails to clear is exactly the defect nobody would see in a test.
   */
  statusNotice: string;
  /**
   * The status bar's other two readouts, read from the elements the render writes: the caret
   * position ("Ln 4, Col 7") and the module name. They complete what statusNotice started - the
   * caret readout is the one place the developer SEES where a Run would land, so a readout that
   * drifts from the native caret is the misdirection debugger-features exists to catch, and
   * until these were here it could drive the caret and never ask what the bar claimed.
   */
  statusPosition: string;
  statusModule: string;
  /**
   * The sync dialog while it is open, null otherwise: direction, folder, mode, busy, the
   * status sentence, and every row with its tick. Before this the dialog's rows could only be
   * read by querySelector from a harness file, which is a test of the selector as much as of
   * the dialog - and the one thing that catches the api and the dialog disagreeing about an
   * import is comparing exactly this against the sync route's plan.
   */
  sync: {
    direction: string;
    folder: string;
    mode: string;
    busy: boolean;
    status: string;
    rows: { file: string; status: string; detail: string; ticked: boolean; actionable: boolean }[];
  } | null;
  /**
   * The change log as the Changes pane is drawing it: which project, which files it can be
   * pointed at, the newest round marked reviewed, and the rounds with one row per module. Null
   * before the pane has been built.
   */
  /**
   * The agent card while it stands, null otherwise: whether the api door is open, and the
   * instruction text as it is - which is what a developer would paste, read from the box they
   * would paste it out of rather than rebuilt here from the same inputs.
   */
  agent: {
    open: boolean;
    api: boolean;
    busy: boolean;
    text: string;
    copied: boolean;
  } | null;
  changes: {
    project: string;
    files: string[];
    acceptedAt: number;
    busy: boolean;
    rounds: {
      round: number;
      by: string;
      label: string | null;
      open: boolean;
      accepted: boolean;
      modules: { module: string; added: number; removed: number; held: boolean }[];
    }[];
    showing: string | null;
    full: boolean;
  } | null;
  /** Main-thread stalls over 50ms, worst first. What the surface felt like, in numbers. */
  longTasks: LongTask[];
  /** How many models and documents are alive, since a leak shows here first. */
  census: { models: number; documents: number };
  /** The find/replace widget: open, its query and scope, and how many matches it holds. */
  search: {
    open: boolean;
    query: string;
    replacement: string;
    scope: string;
    matchCase: boolean;
    wholeWord: boolean;
    matches: number;
    current: number;
    replaceShown: boolean;
  };
  /** The bookmark lines of the model on screen, read from its live decorations. */
  bookmarks: number[];
  /**
   * What is at a position, when `ui` was asked with `line`/`column` or `word`.
   *
   * The COLOUR is read off the rendered span, not derived from a token type and a theme map:
   * what a developer means by "is this word the wrong colour" is the pixel, and every step
   * between the tokeniser and the pixel is a step that can be wrong. The squiggles are monaco's
   * markers at that position, which is exactly what draws the underline.
   */
  at: AtPosition | null;
}

export interface AtPosition {
  line: number;
  column: number;
  word: string | null;
  /** The monaco token class on the rendered span, e.g. "mtk12". Null when nothing is rendered. */
  tokenClass: string | null;
  /** The computed colour of that span: what is actually on screen. */
  colour: string | null;
  /** Bold, italic and underline as rendered, since a theme can carry those too. */
  style: { fontWeight: string; fontStyle: string; textDecoration: string } | null;
  /** Every marker covering the position: this is the squiggle. */
  squiggles: {
    severity: string;
    message: string;
    code: string | null;
    owner: string;
    startColumn: number;
    endColumn: number;
  }[];
}

export interface DevSurfaceParts {
  workspace: Workspace;
  explorer: Explorer;
  bridge: EditorBridge;
  /** The designer tabs, for driving the markup-apply path a Ctrl+S takes. */
  designer: {
    viewFor(module: string, project: string | null): {
      applyDocument(markup: string): Promise<{ ok: boolean; added: string[]; removed: string[]; set: number; refused?: string | null }>;
      markupText(): string;
      markupCaret(line: number): string;
      setDocument(markup: string): void;
      lintMarkers(): { line: number; message: string; severity: string }[];
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
      };
      /** Whether the PROJECTION holds a control of this name, which is a different question
       * from whether the canvas is drawing one: a control on a closed page is both real and
       * undrawn, and selecting it is what opens the page. */
      knows(name: string): boolean;
      /** Whether a Properties row about this target and property is one this document can
       * carry: the form or a control it holds, and an attribute the dialect can spell. */
      spells(target: string, property: string): boolean;
      completions(line: number, column: number): {
        label: string; detail: string | null; documentation: string | null; insert: string;
        replaces: { from: number; to: number };
      }[];
      hover(line: number, column: number): string[];
      headerHint(line: number, column: number): { label: string; active: number; parameter: string } | null;
      select(name: string, extend?: boolean): void;
      marqueeOver(left: number, top: number, right: number, bottom: number): string;
      arrange(how: string): string;
      zorder(names: string[], front: boolean): string;
      showTabOrder(): string;
      setZoom(what: number | "fit"): string;
      zoomPercent(): number;
      requestEventStub(control: string | null): void;
      dragControl(name: string, dx: number, dy: number, alt?: boolean): string;
      resizeControl(name: string, edge: string, dx: number, dy: number, alt?: boolean): string;
      deleteControl(name: string): string;
      addFromToolbox(kind: string, left: number, top: number): string;
      openTabOn(container: string, which: string): string;
      openTabMenu(container: string, which: string): string;
      format(how: string): Promise<string>;
      copySelection(): string;
      cutSelection(): string;
      pasteClipboard(): string;
      duplicateSelection(): string;
    } | null;
    /** Applies and saves every dirty designer document; sync's export flush rides this. */
    saveDirty(): Promise<{ saved: string[]; refused: string[] }>;
  };
  search: {
    state(): UiSnapshot["search"];
    find(query: string, options?: { scope?: string; matchCase?: boolean; wholeWord?: boolean }): void;
    open(options?: { scope?: string; withReplace?: boolean }): void;
    close(): void;
    /**
     * The panel's buttons. `find` types a query and raises `input`, whose only handler searches
     * when the scope is "module" - so under any other scope it types and searches nothing, which
     * is what a person sees too before they press Find All. These are those presses.
     */
    runFindAll(): void;
    setReplacement(text: string): void;
    runReplaceAll(): void;
    runReplaceCurrent(): void;
    goToNextMatch(): void;
    goToPreviousMatch(): void;
  };
  bookmarks: { marksOn(model: monaco.editor.ITextModel): number[] };
  panes: {
    list(): { name: string; title: string; open: boolean; permanent: boolean }[];
    /** Moves a pane to a dock side, through the method a real drop calls. */
    moveTo(name: string, side: "left" | "right" | "top" | "bottom"): boolean;
  };
  /**
   * The registered language providers themselves, so the api asks what the editor asks.
   *
   * Not the bridge requests underneath them: every provider refuses to answer for anything but
   * the host-active module, and that refusal is the difference between a feature that works and
   * one that is silent on screen.
   */
  providers: {
    hover: monaco.languages.HoverProvider;
    completion: monaco.languages.CompletionItemProvider;
    signature: monaco.languages.SignatureHelpProvider;
    codeAction: monaco.languages.CodeActionProvider;
    definition: monaco.languages.DefinitionProvider;
    rename: monaco.languages.RenameProvider;
  };
  /**
   * Every use of the symbol at a position, as the references DIALOG is given them.
   *
   * Not a monaco provider: this product deliberately registers none, because Go to References
   * and its peek were gated off in favour of its own list. So the honest mirror is the function
   * the menu entry calls, which is this one.
   */
  referencesAt(position: { line: number; column: number }): Promise<{ word: string; found: unknown[] } | null>;
  /**
   * Opens the references DIALOG at a position - the same one Shift+F12 opens, through the same
   * lookup referencesAt uses. Answers whether a dialog was opened. The dialog's rendering is the
   * feature (it draws a module with no tab open), so a data-only lookup does not exercise it.
   */
  openReferences(position: { line: number; column: number }): Promise<boolean>;
  openSettings(): void;
  openSponsors(): void;
  /** What the status line is showing, read from the element the render writes. */
  statusNotice(): string;
  /** The status bar's caret readout ("Ln 4, Col 7"), read the same way. */
  statusPosition(): string;
  /** The status bar's module name, read the same way. */
  statusModule(): string;
  /** Presses a toolbar command by id, through the button's own click. */
  pressToolbar(id: string): { did: boolean; detail: string };
  /** Every command the strip is currently drawing, and whether each is pressable. */
  toolbarCommands(): { id: string; label: string; disabled: boolean }[];
  /**
   * The Properties panel: what it is showing, and the one way to change a value.
   *
   * It writes real component state through the object model's own setter - renaming a module is
   * the "(Name)" row - and until 2026-08-11 nothing in the api could drive it, read it, or even
   * name it. It was the only user-visible surface with no presence in either direction.
   */
  properties(): { component: string; kind: string; round: number; rows: { name: string; value: string; writable: boolean; boolean: boolean; options?: string[] | null; swatch?: string | null; picture?: boolean; preview?: string }[] };
  editProperty(name: string, value: string): boolean;
  /**
   * Changes one setting THROUGH THE PAGE, the way the dialog's own controls do.
   *
   * The `settings` route on the host is a different path and always was a safer one: it reads the
   * stored settings and replaces one field. The page posts what it believes the whole settings
   * object to be, which is where a field can go missing - and one did, silently resetting the
   * chosen sync planner on every unrelated change. Nothing could reach this path, so nothing saw
   * it. The dialog's handlers call exactly this.
   */
  changeSetting(key: string, value: unknown): boolean;
}

/**
 * Whatever is standing in front of the page, asked of the DOM rather than of a list.
 *
 * The first version of this was a hand-kept map of dialog element ids, and it was wrong on the
 * day it shipped: it did not name `close-confirm-backdrop`, which is the ONE dialog most likely
 * to be up when something looks stuck. A Save / Don't Save / Cancel box sat on screen through an
 * entire invariant sweep while this reported no dialogs at all, and every close in that sweep
 * failed for a reason nothing could see (2026-08-07).
 *
 * Every dialog this page builds already marks itself `aria-modal="true"`, because that is what
 * makes it a dialog to a screen reader. So a new dialog is reported the moment it exists, with no
 * list to remember to update. The name is its accessible label, its title text, or its id, in that
 * order of usefulness.
 */
/** A token that never cancels: these calls are asked for deliberately and answered in full. */
const NO_CANCEL = {
  isCancellationRequested: false,
  onCancellationRequested: () => ({ dispose: () => { } }),
} as never;

/** The trigger context monaco fills in. Nothing in these providers reads it. */
const EMPTY_CONTEXT = {} as never;

function dialogsUp(): DialogSnapshot[] {
  return [...document.querySelectorAll('[aria-modal="true"]')].map((element) => {
    // The card carries the modal marking; the backdrop is what holds it and what a close removes.
    const root = element.closest("[id]") ?? element;
    const label = element.getAttribute("aria-label")
      ?? element.querySelector("[id$='-title']")?.textContent
      ?? root.id;

    return { id: root.id || element.id, title: (label ?? "").trim() };
  });
}

/**
 * Long tasks, which is what jank IS.
 *
 * A frame is 16ms. Anything holding the main thread longer than 50ms is a stretch during which
 * the surface answered no key, painted nothing, and scrolled nowhere, and no counter on the host
 * side can see it: the host thread was fine the whole time.
 *
 * Two entry types, because they answer different halves. `longtask` says a stall happened and how
 * long it was. `long-animation-frame` names the SCRIPT, with its source and the function that
 * invoked it, which is the difference between "something is slow" and a file to open.
 *
 * A CAVEAT worth more than the instrument: **a stall provoked through the `eval` route is not a
 * page task and neither type sees it.** Testing this by holding the main thread from a debug-api
 * script reported nothing at all, which reads exactly like a broken observer, and would have read
 * exactly like "no jank" had it been believed (2026-08-07). Provoke from inside the page - a
 * setTimeout, a real interaction - or provoke nothing and read what the session collected.
 */
interface LongTask {
  startedAt: number;
  durationMs: number;
  /** For a long-animation-frame: the script and the call that ran it. Empty for a bare longtask. */
  attribution: string;
}

const longTasks: LongTask[] = [];

function watchLongTasks(): void {
  // Not every engine ships either type, and a PerformanceObserver constructed with an
  // unsupported one THROWS rather than observing nothing. Nothing here may take the page down.
  const supported = (PerformanceObserver.supportedEntryTypes ?? [])
    .filter((type) => type === "longtask" || type === "long-animation-frame");

  if (supported.length === 0) {
    return;
  }

  // ONE OBSERVER PER TYPE, deliberately. `buffered: true` replays what was recorded before the
  // observer existed, and it is only accepted with the single-`type` form: an `entryTypes` list
  // takes both types and silently gives up the history, which is exactly where the long ones are.
  // Page start-up is the worst stretch this product has, and it happens before any of this runs.
  const install = (type: string) => {
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) {
        // A long-animation-frame and a longtask report the SAME stall, so the frame's richer
        // entry replaces the bare one rather than doubling it.
        const frame = entry as PerformanceEntry & {
          scripts?: { name?: string; invoker?: string; sourceURL?: string; duration?: number }[];
        };

        const worst = (frame.scripts ?? [])
          .slice()
          .sort((a, b) => (b.duration ?? 0) - (a.duration ?? 0))[0];

        const attribution = worst
          ? `${worst.invoker ?? worst.name ?? "?"} ${worst.sourceURL ?? ""}`.trim()
          : "";

        const startedAt = Math.round(entry.startTime);
        const already = longTasks.findIndex((one) => Math.abs(one.startedAt - startedAt) <= 2);
        const row = { startedAt, durationMs: Math.round(entry.duration), attribution };

        if (already >= 0) {
          // Keep whichever knows more. The bare longtask usually lands first.
          if (attribution !== "" || (longTasks[already]?.durationMs ?? 0) < row.durationMs) {
            longTasks[already] = row;
          }
        } else {
          longTasks.push(row);
        }
      }

      // Worst first, capped, so one bad minute cannot push out the record holder.
      longTasks.sort((a, b) => b.durationMs - a.durationMs);
      longTasks.length = Math.min(longTasks.length, 24);
    }).observe({ type, buffered: true });
  };

  for (const type of supported) {
    try {
      install(type);
    } catch {
      // One type failing must not cost the other. An empty list is honest: nothing observed.
    }
  }
}

/**
 * A boolean argument, from a caller that may have sent one or may have sent a query value.
 *
 * The `act` route carries arguments as query values, so `open=false` arrives as the STRING
 * "false", and `Boolean("false")` is true. That collapsed expand and collapse into one action
 * the first time it was driven through the door (2026-08-07), which is the whole reason the
 * coercion lives here rather than being written out per argument.
 */
function flag(value: unknown, whenMissing: boolean): boolean {
  if (value === undefined || value === null || value === "") {
    return whenMissing;
  }
  if (typeof value === "string") {
    return !["false", "0", "no", "off"].includes(value.toLowerCase());
  }
  return Boolean(value);
}

/**
 * The result of an action: what it did, in words a failing test can print.
 *
 * `did` is false for "the page declined", which is a real answer and not an error: closing a tab
 * that is not open, unfolding a workbook that is not there. A script that treats those as throws
 * ends up wrapped in try/catch and stops distinguishing them from a broken door.
 */
export interface ActResult {
  did: boolean;
  detail: string;
  /** What the action found, when it is a question as much as an act. */
  data?: unknown;
}

/** An action's answer. A promise for anything whose outcome crosses to the host and back. */
type ActAnswer = ActResult | Promise<ActResult>;

export function installDevSurface(parts: DevSurfaceParts): void {
  const { workspace, explorer, bridge, designer } = parts;

  /**
   * A position argument as an offset into the active model.
   *
   * Takes `line`/`column`, or `word` to find the first occurrence of an identifier - which is
   * how a person describes where they are looking ("hover over Recalculate") without counting
   * columns. Null when nothing is open.
   */
  /** A position as monaco wants it, with the model it belongs to. Null when nothing is open. */
  const positionFrom = (args: Record<string, unknown>):
    { model: monaco.editor.ITextModel; position: monaco.Position } | null => {
    const model = workspace.activeEditor().getModel();
    const offset = offsetFrom(args);
    if (!model || offset === null) { return null; }

    const at = model.getPositionAt(offset);
    return { model, position: new monacoApi.Position(at.lineNumber, at.column) };
  };

  const offsetFrom = (args: Record<string, unknown>): number | null => {
    const model = workspace.activeEditor().getModel();
    if (!model) { return null; }

    if (args.word !== undefined) {
      // CASE-INSENSITIVE, because VBA is: the language does not distinguish `total` from `Total`,
      // and the host RECASES identifiers on write to match their declaration - writing
      // `total = 1` through the object model comes back `Total = 1`. A case-sensitive lookup
      // therefore fails to find a word the caller can see on screen, and answers "no such word"
      // about a word that is right there (2026-08-08).
      const wanted = String(args.word).toLowerCase();
      const at = model.getValue().toLowerCase().indexOf(wanted);
      if (at < 0) { return null; }

      // The middle of the word, so a provider keyed on "inside an identifier" is satisfied.
      return at + Math.floor(wanted.length / 2);
    }

    const line = Number(args.line ?? 0);
    const column = Number(args.column ?? 1);
    if (!Number.isFinite(line) || line < 1) { return null; }

    return model.getOffsetAt({ lineNumber: line, column });
  };

  // Installed at wiring time, not at first ask: `buffered: true` recovers what the observer
  // would have seen slightly earlier, but nothing recovers a stall that happened before the
  // page had an observer at all, and start-up is where the long ones live.
  watchLongTasks();

  const editorFocus = (): UiSnapshot["focus"] => {
    const editor = workspace.activeEditor();
    const position = editor.getPosition();
    const model = editor.getModel() as monaco.editor.ITextModel | null;

    const hostModel = bridge.hostActiveModel();
    const hostEditor = hostModel ? workspace.editorShowing(hostModel) : null;
    const hostPosition = hostEditor?.getPosition();

    return {
      model: model ? model.uri.toString() : null,
      hasFocus: editor.hasTextFocus(),
      line: position?.lineNumber ?? 0,
      column: position?.column ?? 0,
      host: hostModel && hostPosition
        ? { model: hostModel.uri.toString(), line: hostPosition.lineNumber, column: hostPosition.column }
        : null,
    };
  };

  /**
   * Colour and squiggles at a position.
   *
   * The colour comes off the DOM: monaco renders each token as a span with a class the theme
   * colours, so the computed style of the span covering the column is what the developer is
   * looking at. Deriving it from the token type instead would agree with the theme and disagree
   * with the screen whenever the interesting bug is between them.
   *
   * The line has to be RENDERED to be read. Monaco only builds spans for visible lines, so a
   * position outside the viewport answers a null colour rather than a wrong one.
   */
  const describeAt = (where: { lineNumber: number; column: number }): AtPosition => {
    const editor = workspace.activeEditor();
    const model = editor.getModel();
    const word = model?.getWordAtPosition(where) ?? null;

    const markers = model
      ? monacoApi.editor.getModelMarkers({ resource: model.uri })
          .filter((marker) =>
            marker.startLineNumber <= where.lineNumber && marker.endLineNumber >= where.lineNumber
            && marker.startColumn <= where.column && marker.endColumn >= where.column)
      : [];

    const severityName: Record<number, string> = { 1: "hint", 2: "info", 4: "warning", 8: "error" };

    // The rendered span under the column, found by walking the line's text nodes and counting
    // characters: monaco splits a line into spans per token, and the column tells which one.
    let tokenClass: string | null = null;
    let colour: string | null = null;
    let style: AtPosition["style"] = null;

    const lineNode = editor.getDomNode()?.querySelectorAll(".view-line")[
      where.lineNumber - (editor.getVisibleRanges()[0]?.startLineNumber ?? 1)
    ];

    if (lineNode) {
      let seen = 0;
      for (const span of [...lineNode.querySelectorAll("span span")] as HTMLElement[]) {
        const length = (span.textContent ?? "").length;
        if (where.column - 1 < seen + length) {
          const computed = getComputedStyle(span);
          tokenClass = span.className;
          colour = computed.color;
          style = {
            fontWeight: computed.fontWeight,
            fontStyle: computed.fontStyle,
            textDecoration: computed.textDecorationLine,
          };
          break;
        }
        seen += length;
      }
    }

    return {
      line: where.lineNumber,
      column: where.column,
      word: word?.word ?? null,
      tokenClass,
      colour,
      style,
      squiggles: markers.map((marker) => ({
        severity: severityName[marker.severity] ?? String(marker.severity),
        message: marker.message,
        code: typeof marker.code === "string" ? marker.code : (marker.code?.value ?? null),
        owner: marker.owner ?? "",
        startColumn: marker.startColumn,
        endColumn: marker.endColumn,
      })),
    };
  };

  /**
   * The Properties rows with any picture's BYTES taken out and its size left behind.
   *
   * A thumbnail's data URI is a whole bitmap in base64 - kilobytes for an icon, megabytes for a
   * photograph - and a snapshot is read on every wait in every suite. What a caller can actually
   * check is that a row HAS a picture and that it changed, which `previewBytes` answers in eight
   * characters. The pixels are proved where pixels belong: against the running form's photograph.
   */
  const withoutPictureBytes = (shown: ReturnType<typeof parts.properties>): ReturnType<typeof parts.properties> => ({
    ...shown,
    // `preview` arrives NULL rather than absent from the host's record, so the test is for
    // either: reading `.length` off the null took the whole snapshot down with it, which is a
    // failure every suite sees at once because every wait reads a snapshot.
    rows: shown.rows.map(({ preview, ...row }) => (
      preview ? { ...row, previewBytes: preview.length } : row)),
  });

  const state = (at?: { lineNumber: number; column: number }): UiSnapshot => ({
    workspace: workspace.snapshot(),
    explorer: explorer.treeState(),
    panes: parts.panes.list().map(({ name, title, open }) => ({ name, title, open })),
    dialogs: dialogsUp(),
    waiting: bridge.waitingOn(),
    focus: editorFocus(),
    settings: { ...currentSettings() } as unknown as Record<string, unknown>,
    emptyViewShown: workspace.emptyViewShown(),
    properties: withoutPictureBytes(parts.properties()),
    statusNotice: parts.statusNotice(),
    statusPosition: parts.statusPosition(),
    statusModule: parts.statusModule(),
    sync: syncDialogProbe()?.state() ?? null,
    changes: changesPaneProbe()?.state() ?? null,
    agent: agentDialogProbe()?.state() ?? null,
    longTasks: [...longTasks],
    census: bridge.modelCensus(),
    search: parts.search.state(),
    bookmarks: (() => {
      const model = workspace.activeEditor().getModel();
      return model ? parts.bookmarks.marksOn(model) : [];
    })(),
    at: at ? describeAt(at) : null,
  });

  /**
   * A keystroke as the page receives one.
   *
   * The chords this product owns are bound on `document` with `capture: true` rather than through
   * monaco's keybinding service, so they are reachable this way and only this way: the shim's
   * `type` route goes through the HOST's keyboard pipeline, which is a different path with a
   * different set of interceptions. Ctrl+W is the case that matters, because it is intercepted on
   * both sides and the two halves have disagreed before.
   */
  const sendKey = (args: Record<string, unknown>): ActResult => {
    const code = String(args.code ?? "");
    if (!code) {
      return { did: false, detail: "no code given; pass code: \"KeyW\"" };
    }

    const target = args.target === "document"
      ? document
      : (document.activeElement ?? document.body);

    const event = new KeyboardEvent("keydown", {
      code,
      key: String(args.key ?? code.replace(/^Key/, "").toLowerCase()),
      ctrlKey: flag(args.ctrl, false),
      shiftKey: flag(args.shift, false),
      altKey: flag(args.alt, false),
      bubbles: true,
      cancelable: true,
    });

    const delivered = target.dispatchEvent(event);
    return {
      did: true,
      detail: `${code} sent to ${target === document ? "document" : "the focused element"}`
        + (delivered ? "" : ", and something handled it"),
    };
  };

  const actions: Record<string, (args: Record<string, unknown>) => ActAnswer> = {
    /*
     * The designer tab's own apply - the path Ctrl+S takes, through the view's document and
     * the host round trip, so a suite proves the TAB works and not merely the route the tab
     * shares a service with. Answers the outcome the strip under the document shows.
     */
    designerSaveDirty: async () => {
      // Sync's export flush: what lands in the folder is what is on screen. No arguments -
      // the dirty set is the page's own knowledge, and a caller cannot know it better.
      const flushed = await designer.saveDirty();
      return {
        did: flushed.refused.length === 0,
        detail: flushed.refused.length > 0
          ? `refused: ${flushed.refused.join("; ")}`
          : flushed.saved.length === 0
            ? "nothing was dirty"
            : `saved ${flushed.saved.join(", ")}`,
        data: flushed,
      };
    },

    designerApply: async (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const markup = typeof args.markup === "string" ? args.markup : null;
      if (!module || markup === null) {
        return { did: false, detail: "designerApply takes module and markup" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const outcome = await view.applyDocument(markup);
      return {
        did: outcome.ok,
        detail: outcome.ok
          ? `+${outcome.added.length} -${outcome.removed.length} set ${outcome.set}`
          : outcome.refused ?? "refused",
        data: outcome,
      };
    },

    /** Sets the designer tab's document WITHOUT applying - the typing path, for driving
     * the squiggles. */
    designerSetMarkup: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const markup = typeof args.markup === "string" ? args.markup : null;
      if (!module || markup === null) {
        return { did: false, detail: "designerSetMarkup takes module and markup" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      view.setDocument(markup);
      return { did: true, detail: `${markup.length} char(s) set, not applied` };
    },

    /*
     * The markup language service, asked where a developer would ask it. Each of these places
     * the CARET at the spot first, exactly as the gesture that raises the widget does, and then
     * reads the real provider - there is no second copy of the answer for a probe to read.
     */

    /** What the completion widget offers at line/column, in `data`. */
    designerComplete: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerComplete takes module, line and column" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const items = view.completions(Number(args.line ?? 1), Number(args.column ?? 1));
      return {
        did: true,
        detail: items.length === 0
          ? "nothing is offered here"
          : `${items.length} suggestion(s): ${items.slice(0, 6).map((item) => item.label).join(", ")}`,
        data: items,
      };
    },

    /** What a hover at line/column says, block by block, in `data`. */
    designerHover: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerHover takes module, line and column" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const blocks = view.hover(Number(args.line ?? 1), Number(args.column ?? 1));
      return {
        did: blocks.length > 0,
        detail: blocks.length === 0 ? "nothing to say here" : `${blocks.length} block(s)`,
        data: blocks,
      };
    },

    /** The header hint at line/column: the grammar and the clause it points at, in `data`. */
    designerHint: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerHint takes module, line and column" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const hint = view.headerHint(Number(args.line ?? 1), Number(args.column ?? 1));
      return {
        did: hint !== null,
        detail: hint === null ? "no header hint here" : `${hint.label} - on ${hint.parameter}`,
        data: hint,
      };
    },

    /** The markup document's current squiggles, in `data`. */
    designerLint: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerLint takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      return view
        ? { did: true, detail: `${view.lintMarkers().length} finding(s)`, data: view.lintMarkers() }
        : { did: false, detail: `no designer tab is open for ${module}` };
    },

    /** The canvas as rendered, in `data`: whether the picture is the DRAFT the document
     * describes or the applied projection, and every control's name and placed position. */
    designerCanvas: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerCanvas takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const snapshot = view.canvasSnapshot();
      return {
        did: true,
        detail: `${snapshot.controls.length} control(s), ${snapshot.draft ? "draft" : "applied"}`,
        data: snapshot,
      };
    },

    /** Selects on the canvas by name - the click's own entry; "" or omitted selects the
     * FORM. The markup caret follows; read the result through designerCanvas. */
    designerSelect: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerSelect takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const name = typeof args.control === "string" ? args.control : "";
      // Against the PROJECTION rather than against the drawing. An act that names a control the
      // form does not have answers false rather than setting a selection nothing can see - a row
      // that went on to open the tab-order dialog for a control deleted three sections earlier
      // was told the selection had taken. But a control on a page that is NOT OPEN is perfectly
      // real, and selecting it is what opens that page: asking the canvas conflated the two and
      // refused the second (2026-08-16).
      if (!view.knows(name)) {
        return { did: false, detail: `${name} is not on this form to select` };
      }

      // `extend` is Ctrl+click: it adds the control to the selection, or takes it back out, and
      // leaves the anchor where it was. Read the whole group back through designerCanvas.
      const extend = args.extend === true || args.extend === "1" || args.extend === "true";
      view.select(name, extend);
      return {
        did: true,
        detail: extend
          ? `the selection is ${view.canvasSnapshot().group.join(", ") || "empty"}`
          : name === "" ? "the form is selected" : `${name} is selected`,
      };
    },

    /**
     * The rubber band over the form's own ground: a real pointer drag from one corner to the
     * other, in POINTS from the form's client origin, selecting every control of that ground it
     * TOUCHES. The read side is designerCanvas's `group`, anchor first.
     */
    designerMarquee: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerMarquee takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const corners = ["left", "top", "right", "bottom"].map((one) => Number(args[one] ?? Number.NaN));
      if (corners.some((one) => !Number.isFinite(one))) {
        return { did: false, detail: "designerMarquee takes left, top, right and bottom in points" };
      }

      const [left, top, right, bottom] = corners as [number, number, number, number];
      const outcome = view.marqueeOver(left, top, right, bottom);
      return { did: outcome.startsWith("banded"), detail: outcome };
    },

    /**
     * Lines the selection up on its ANCHOR, sizes it to the anchor, or spreads it evenly - the
     * native Format menu's arrange group, which lives on this canvas's own context menu because
     * the product has no menu bar and the editor's Format menu would act on the native
     * designer's selection rather than ours.
     *
     * `how` is left, centreX, right, top, centreY, bottom, width, height, across or down. It
     * writes the DOCUMENT as one undoable edit, like every other canvas gesture.
     */
    designerArrange: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const how = typeof args.how === "string" ? args.how : null;
      if (!module || !how) {
        return { did: false, detail: "designerArrange takes module and how" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const outcome = view.arrange(how);
      return { did: outcome.startsWith(how), detail: outcome };
    },

    /**
     * The Format gestures that act on each control ALONE rather than lining them up with each
     * other: `act("designerFormat", { module, how: "centreX"|"centreY"|"fit"|"grid" })`.
     *
     * Separate from designerArrange because these are meaningful for one control - arrange
     * refuses a selection smaller than two, and refusing "Size to Fit" on one button would be
     * refusing the case it is for.
     */
    designerFormat: async (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const how = typeof args.how === "string" ? args.how : null;
      if (!module || !how) {
        return { did: false, detail: "designerFormat takes module and how" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const outcome = await view.format(how);
      return { did: outcome.startsWith(how), detail: outcome };
    },

    /**
     * The Tab Order dialog: `open` shows it for the container the selection sits in, `move` sends
     * one control a place up or down, and with neither it answers the order the list is showing.
     *
     * The rows are the container's controls in TAB order, which is not the order the canvas draws
     * them in - the walk reads a container's collection in creation order. A move is a TabIndex
     * write through the host's own SetControlProperty, and MSForms renumbers the rest, so read the
     * result back off `designer` rather than trusting the list.
     */
    designerTabOrder: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerTabOrder takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      if (args.open === true || args.open === "1" || args.open === "true") {
        const outcome = view.showTabOrder();
        return { did: outcome.startsWith("tab order"), detail: outcome };
      }

      const open = openTabOrderDialog();
      if (!open) {
        return { did: false, detail: "no tab-order dialog is open" };
      }

      // Closed by its own Close button rather than by the handle, for the mirror rule - and
      // because a modal left standing is a backdrop over the canvas that fails every row after
      // it, which is how this verb came to exist.
      if (args.close === true || args.close === "1" || args.close === "true") {
        const button = document.querySelector<HTMLElement>("#taborder-card .modal-button.primary");
        button?.click();
        return { did: openTabOrderDialog() === null, detail: "the tab-order dialog is closed" };
      }

      const control = typeof args.control === "string" ? args.control : null;
      if (control !== null) {
        const by = String(args.move ?? "up") === "down" ? 1 : -1;
        const moved = open.move(control, by);
        return {
          did: moved,
          detail: moved
            ? `${control} moved ${by === 1 ? "down" : "up"}: ${open.order().join(", ")}`
            : `${control} cannot move ${by === 1 ? "down" : "up"}`,
        };
      }

      return { did: true, detail: open.order().join(", "), data: open.order() };
    },

    /**
     * How large the canvas draws the form: `to` is a percentage or `fit`, and with neither it
     * answers the zoom that is showing.
     *
     * Everything a gesture does stays in POINTS whatever the zoom, which is the thing worth
     * pinning: `designerDrag` by 12 points moves a control 12 points at 200% as it does at 100%,
     * because only the screen boundary knows about the scale.
     */
    designerZoom: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerZoom takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      if (args.to === undefined) {
        return { did: true, detail: `zoom ${view.zoomPercent()}%`, data: view.zoomPercent() };
      }

      const asked = String(args.to);
      const outcome = asked.toLowerCase() === "fit"
        ? view.setZoom("fit")
        : view.setZoom(Number(asked) / 100);
      return { did: outcome.startsWith("zoom"), detail: outcome };
    },

    /**
     * Bring to Front / Send to Back on the whole selection - `front=1` for the front.
     *
     * The one canvas gesture that writes the MODEL rather than the document, so there is nothing
     * in `designerMarkup` to read back and nothing on the canvas to see: MSForms' Controls
     * collection is not in z-order and does not move when ZOrder is called (measured). The proof
     * is the RUNNING form - launch it and photograph it, the way designer-parity.mjs does.
     */
    designerZOrder: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerZOrder takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const front = args.front === true || args.front === "1" || args.front === "true";
      const control = typeof args.control === "string" && args.control.length > 0 ? args.control : null;
      const outcome = control === null
        ? view.zorder(view.canvasSnapshot().group.filter((one) => one !== ""), front)
        : view.zorder([control], front);
      return { did: !outcome.startsWith("nothing"), detail: outcome };
    },

    /**
     * Opens a tab of a MultiPage or a TabStrip on the canvas - the real press on the real tab.
     * `tab` is a page's name, the tab's label, or a 1-based position.
     *
     * A MultiPage draws that page's controls and selects the PAGE, as the native designer does.
     * A TabStrip only marks the tab: its tabs are an index rather than containers, and the
     * runtime draws the same controls under every one. Read it back through designerCanvas,
     * whose `containers` say which tab each strip is showing.
     */
    designerOpenTab: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const container = typeof args.container === "string" ? args.container : null;
      const tab = args.tab === undefined ? null : String(args.tab);
      if (!module || !container || tab === null) {
        return { did: false, detail: "designerOpenTab takes module, container and tab" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const outcome = view.openTabOn(container, tab);
      return { did: outcome.startsWith("opened"), detail: outcome };
    },

    /**
     * Right-clicks a MultiPage's tab strip and reports the menu it produced - New Page and
     * Delete Page, the native designer's own pair. `tab` names the tab to right-click, or is
     * omitted for the strip itself; a right-click on a tab opens that page first, so the item
     * chosen acts on the page named. Pick an item with chooseMenuItem.
     *
     * A TabStrip opens NO menu, deliberately: its tabs are not in the markup dialect, so there
     * is no line to add or take away and every item would be a claim this cannot keep.
     */
    designerTabMenu: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const container = typeof args.container === "string" ? args.container : null;
      if (!module || !container) {
        return { did: false, detail: "designerTabMenu takes module and container" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const outcome = view.openTabMenu(container, args.tab === undefined ? "" : String(args.tab));
      if (!outcome.startsWith("right-clicked")) {
        return { did: false, detail: outcome };
      }

      // Disabled items are SAID to be disabled: this menu greys Delete Page for a MultiPage
      // with no pages, and a row that could not tell the two apart would pass either way.
      const labels = [...document.querySelectorAll<HTMLElement>(".menu-dropdown .menu-item")]
        .map((one) => (one.textContent ?? "").trim() + (one.classList.contains("disabled") ? " (disabled)" : ""));
      return labels.length > 0
        ? { did: true, detail: labels.join(" | ") }
        : { did: false, detail: `${outcome} and no menu opened` };
    },

    /** Drags a control on the canvas by a delta in POINTS - the document's own unit - through
     * the REAL pointer sequence, on the element the hit test answers. The drop rewrites the
     * control's line in the document; read the result through designerCanvas or
     * designerMarkup. */
    designerDrag: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const control = typeof args.control === "string" ? args.control : null;
      if (!module || !control) {
        return { did: false, detail: "designerDrag takes module and control" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const dx = Number(args.dx ?? 0);
      const dy = Number(args.dy ?? 0);
      if (!Number.isFinite(dx) || !Number.isFinite(dy)) {
        return { did: false, detail: "dx and dy are deltas in points" };
      }

      // `alt=1` holds the Alt key for the whole gesture, which overrides every snap.
      const held = args.alt === true || args.alt === "1" || args.alt === "true";
      const outcome = view.dragControl(control, dx, dy, held);
      return {
        did: outcome.startsWith("dragged"),
        detail: `${outcome} by ${dx},${dy}pt${held ? " with alt held" : ""}`,
      };
    },

    /** Resizes by a HANDLE, the way a hand does: the thing is selected, its `edge` handle -
     * nw, n, ne, e, se, s, sw, w - is pressed through the hit test and pulled by a delta in
     * POINTS. `control` omitted means the FORM's own frame. The drop rewrites that line's
     * geometry in the document, so the read side is designerCanvas or designerMarkup. */
    designerResize: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const edge = typeof args.edge === "string" ? args.edge : null;
      if (!module || !edge) {
        return { did: false, detail: "designerResize takes module and edge (nw, n, ne, e, se, s, sw, w)" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const dx = Number(args.dx ?? 0);
      const dy = Number(args.dy ?? 0);
      if (!Number.isFinite(dx) || !Number.isFinite(dy)) {
        return { did: false, detail: "dx and dy are deltas in points" };
      }

      const heldOnEdge = args.alt === true || args.alt === "1" || args.alt === "true";
      const outcome = view.resizeControl(
        typeof args.control === "string" ? args.control : "", edge, dx, dy, heldOnEdge);
      return {
        did: outcome.startsWith("resized"),
        detail: `${outcome} by ${dx},${dy}pt on ${edge}${heldOnEdge ? " with alt held" : ""}`,
      };
    },

    /** Drags a KIND out of the xlide toolbox and drops it on the form: the real pointer
     * sequence, from the palette button to a point given in POINTS from the form's client
     * origin. The drop writes a new line into the DOCUMENT - named the way the native toolbox
     * names one - and `detail` answers that name; the form gains the control on the tab's
     * Ctrl+S, through the same add the designer route makes. */
    designerToolbox: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const kind = typeof args.kind === "string" ? args.kind : null;
      if (!module || !kind) {
        return { did: false, detail: "designerToolbox takes module and kind" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const left = Number(args.left ?? 0);
      const top = Number(args.top ?? 0);
      if (!Number.isFinite(left) || !Number.isFinite(top)) {
        return { did: false, detail: "left and top are points from the form's client origin" };
      }

      const outcome = view.addFromToolbox(kind, left, top);
      return { did: outcome.startsWith("added"), detail: `${outcome} at ${left},${top}` };
    },

    /** Deletes a control from the canvas: selects it, then presses the real Delete key on the
     * canvas. The control's line - and everything indented under it, its properties and a
     * container's children - leaves the DOCUMENT as one undoable edit; the form itself keeps
     * the control until the tab's Ctrl+S carries the removal through the apply. */
    /**
     * The canvas clipboard: `act("designerClipboard", { module, how: "copy" | "cut" |
     * "paste" | "duplicate" })`.
     *
     * All four write the DOCUMENT and none of them touch the form until Ctrl+S, so read the
     * result through `designerMarkup` rather than through `designer`. The detail names what
     * landed, which is what a row wants to assert against - a paste allocates free names and the
     * caller cannot know them in advance.
     */
    designerClipboard: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const how = String(args.how ?? "").toLowerCase();
      if (!module || !["copy", "cut", "paste", "duplicate"].includes(how)) {
        return { did: false, detail: 'designerClipboard takes module and how: copy|cut|paste|duplicate' };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const said = how === "copy"
        ? view.copySelection()
        : how === "cut"
          ? view.cutSelection()
          : how === "paste" ? view.pasteClipboard() : view.duplicateSelection();

      return { did: /^(copied|cut|pasted)/.test(said), detail: said };
    },

    designerDelete: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const control = typeof args.control === "string" ? args.control : null;
      if (!module || !control) {
        return { did: false, detail: "designerDelete takes module and control" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const outcome = view.deleteControl(control);
      return { did: outcome.startsWith("deleted"), detail: outcome };
    },

    /** The canvas double-click: asks the host for the default event handler - written when
     * absent, shown either way. Omit `control` for the form's own. */
    designerEventStub: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerEventStub takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const control = typeof args.control === "string" && args.control.length > 0 ? args.control : null;
      view.requestEventStub(control);
      return { did: true, detail: `the ${control ?? "form"} stub was requested` };
    },

    /** The designer tab's document as it stands, in `data`. */
    designerMarkup: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      if (!module) {
        return { did: false, detail: "designerMarkup takes module" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      return view
        ? { did: true, detail: `${view.markupText().length} char(s)`, data: view.markupText() }
        : { did: false, detail: `no designer tab is open for ${module}` };
    },

    /** Puts the markup caret on a LINE, which is the drive side of the selection link that
     * runs the other way: a caret inside a control's block selects that control on the canvas,
     * the way clicking the line does. Goes through the editor's own position, so the cursor
     * listener answers it as it answers a click rather than being bypassed. */
    designerCaret: (args) => {
      const module = typeof args.module === "string" ? args.module : null;
      const line = Number(args.line);
      if (!module || !Number.isFinite(line)) {
        return { did: false, detail: "designerCaret takes module and line" };
      }

      const view = designer.viewFor(module, typeof args.project === "string" ? args.project : null);
      if (!view) {
        return { did: false, detail: `no designer tab is open for ${module}` };
      }

      const outcome = view.markupCaret(line);
      return { did: outcome.startsWith("caret"), detail: outcome };
    },

    /*
     * Closing is a REQUEST, and saying otherwise is how this route lied for an afternoon.
     *
     * The host answers a close on a module with unsaved changes by asking the page to confirm,
     * and the tab stays until somebody answers. The first version reported `did: true, closed
     * Watcher` five times running while the tab never moved and a Save / Don't Save / Cancel box
     * stood on screen the whole time. A feature that reports its own success is reporting its
     * intent (lessons-2026-08-07, finding 14) and this was a fresh instance of it, in the very
     * route built to stop probes fooling themselves.
     */
    closeActive: () => {
      const before = workspace.activeDocument();
      if (!before) {
        return { did: false, detail: "nothing is active" };
      }

      workspace.closeActive();

      // WAITED FOR, because a synchronous answer here can only report the request. The close
      // crosses to the host and comes back either as the tab leaving or as a confirm box
      // standing, and neither has happened by the time this line runs. The `act` route awaits a
      // promise, so the honest answer is reachable; the first version returned immediately and
      // said `closed Watcher` five times over a tab that never moved.
      const held = () => workspace.snapshot().groups
        .some((group) => group.tabs.some((tab) =>
          tab.module === before.module && tab.project === before.project));

      return (async () => {
        const deadline = Date.now() + 2000;
        let gone = 0;

        while (Date.now() < deadline) {
          await new Promise((wake) => setTimeout(wake, 50));

          // The DIALOG is checked first, and the tab's absence has to hold for two polls.
          // The two race: a close removes the tab locally before the host's confirm request
          // arrives, so a loop that took the tab's absence as the answer reported `closed`
          // with a Save / Don't Save / Cancel box standing on screen (2026-08-07).
          const standing = dialogsUp();
          if (standing.length > 0) {
            return {
              did: false,
              detail: `close of ${before.module} is waiting on: ${standing.map((one) => one.title).join(", ")}`,
            };
          }

          gone = held() ? 0 : gone + 1;
          if (gone >= 2) {
            return { did: true, detail: `closed ${before.module}` };
          }
        }

        return { did: false, detail: `close of ${before.module} was asked for and nothing came back` };
      })();
    },

    /**
     * Answers the unsaved-changes box the way a person would.
     *
     * Without this the only way past it was a synthesised click on a button found by its text,
     * which is the guessing this module exists to end. `answer` is save, discard, or cancel.
     */
    answerCloseConfirm: (args) => {
      const answer = String(args.answer ?? "discard").toLowerCase();
      const wanted = { save: "Save", discard: "Don't Save", cancel: "Cancel" }[answer];
      if (!wanted) {
        return { did: false, detail: `answer must be save, discard or cancel; got ${answer}` };
      }

      const button = [...document.querySelectorAll("#close-confirm-buttons button")]
        .find((one) => (one.textContent ?? "").trim() === wanted) as HTMLButtonElement | undefined;

      if (!button) {
        return { did: false, detail: "no unsaved-changes box is up" };
      }

      button.click();
      return { did: true, detail: `answered ${wanted}` };
    },

    /**
     * Shows a module, THROUGH THE PATH A TAB CLICK TAKES, and reports whether it landed.
     *
     * This used to call `workspace.reveal` and answer `did: true` regardless. Reveal shows a
     * document the page already holds and tells the host nothing, and the page holds a document
     * only for modules that have been activated, not for every tab in the strip. So against any
     * of the other open tabs it moved nothing and said it had: seven tabs open, one document
     * held, and `activate` cheerfully reporting "revealed Consumer" while the page, the surface
     * and the native pane all stayed on Helpers (2026-08-07).
     *
     * A tab click goes through `selectTab`, which shows it page-locally AND asks the host to
     * activate the native pane behind it. That is the state a developer's click leaves, so it is
     * the state this must leave too.
     *
     * The name is resolved against the open tabs the way every other act resolves one:
     * case-insensitively, project optional. `pickTab` takes an exact identity - as it should,
     * a tab IS its identity - but until 2026-08-13 this act passed the caller's bare name
     * straight through, so `activate` with no project missed every tab whose project was set
     * and answered did:false about a tab that was on screen. A tab click needs no project, so
     * the act must not either. A name two workbooks both hold falls to the shown project's
     * tab; with no tab there it is refused by name rather than resolved to whichever workbook
     * answers first, which is the same refusal the routes give an unmatched workbook.
     */
    activate: (args) => {
      const module = String(args.module ?? "");
      const project = args.project === undefined || args.project === null
        ? null
        : String(args.project);
      if (!module) {
        return { did: false, detail: "no module given" };
      }

      const wantedModule = module.toLowerCase();
      const wantedProject = project === null ? null : project.toLowerCase();
      // The FACE is part of the identity: a form's code tab and designer tab share a name,
      // and without this an activate could pick either. No face means the code tab.
      const wantedFace = args.face === "design" ? "design" : undefined;
      const open = workspace.snapshot().groups.flatMap((group) => group.tabs);
      const matches = open.filter((tab) => tab.module.toLowerCase() === wantedModule
        && (wantedProject === null || (tab.project ?? "").toLowerCase() === wantedProject)
        && (tab.face ?? undefined) === wantedFace);

      if (matches.length === 0) {
        const strip = open.map((tab) => tab.label).join(", ");
        return {
          did: false,
          detail: `no open tab answers to ${module}${project === null ? "" : ` in ${project}`}; open: ${strip.length > 0 ? strip : "nothing"}`,
        };
      }

      let picked = matches[0]!;
      if (matches.length > 1) {
        const shownProject = (workspace.activeDocument()?.project ?? "").toLowerCase();
        const inShown = matches.find((tab) => (tab.project ?? "").toLowerCase() === shownProject);
        if (!inShown) {
          return {
            did: false,
            detail: `${matches.length} workbooks hold an open ${module} and none is the shown project's; pass project`,
          };
        }
        picked = inShown;
      }

      workspace.pickTab({
        module: picked.module,
        project: picked.project,
        ...(picked.face === "design" ? { face: "design" as const } : {}),
      });

      /*
       * AWAITED, because the outcome is not synchronous the first time.
       *
       * A module whose text the page has never held is shown only once the host answers with the
       * document, so a check taken immediately reads the module that was there before. Measured:
       * the first visit to each of three modules reported did=false while all three had in fact
       * arrived a moment later, and the second visit reported did=true because by then the page
       * already held it (2026-08-07).
       *
       * The same shape `closeActive` and `format` use: report what actually happened, having
       * waited for it to happen. The whole identity is compared, not the name: with the same
       * name open from two workbooks, a name-only check would call the wrong workbook's tab
       * success.
       */
      const landed = () => {
        const active = workspace.activeDocument();
        return active !== null
          && active.module.toLowerCase() === picked.module.toLowerCase()
          && (active.project ?? "").toLowerCase() === (picked.project ?? "").toLowerCase()
          && (active.face ?? undefined) === (picked.face ?? undefined);
      };

      return (async () => {
        const deadline = Date.now() + 4000;
        while (Date.now() < deadline) {
          if (landed()) {
            return { did: true, detail: `showing ${picked.module}` };
          }
          await new Promise((resolve) => setTimeout(resolve, 60));
        }

        const active = workspace.activeDocument();
        return { did: false, detail: `asked for ${module}; the page is showing ${active?.module ?? "nothing"}` };
      })();
    },

    cycleTab: (args) => {
      const delta = Number(args.delta ?? 1);
      const landed = workspace.cycleTab(delta);
      return landed
        ? { did: true, detail: `moved to ${landed}` }
        : { did: false, detail: "nothing to cycle to" };
    },

    split: (args) => {
      const direction = args.direction === "down" ? "down" : "right";
      workspace.splitActive(direction);
      return { did: true, detail: `split ${direction}, ${workspace.groupCount()} groups now` };
    },

    // The tree, through the same methods its rows reach, so a script exercises the code a click
    // exercises rather than a parallel path that can drift away from it.
    expandWorkbook: (args) => {
      const workbook = String(args.workbook ?? "");
      const open = flag(args.open, true);
      return explorer.setWorkbookExpanded(workbook, open)
        ? { did: true, detail: `${workbook} ${open ? "expanded" : "collapsed"}` }
        : { did: false, detail: `no workbook named ${workbook}` };
    },

    unfoldModule: (args) => {
      const module = String(args.module ?? "");
      if (!module) {
        return { did: false, detail: "no module given" };
      }
      explorer.unfold(module, args.workbook === undefined ? undefined : String(args.workbook));
      return { did: true, detail: `toggled ${module}` };
    },

    /**
     * Right-clicks a row of the tree and reports the menu it produced.
     *
     * Through the DOM, and deliberately: the menu hangs off a `contextmenu` listener on the tree,
     * so there is no method behind it to call. A real event on the real row is the only thing that
     * exercises what a right-click exercises - which row was marked, which workbook the menu was
     * told about, which items the component's class earns.
     *
     * `module` names a component row; `workbook` alone names a workbook row, and narrows a
     * component row when the same name lives in two open books. Matching is case-insensitive,
     * because the editor unifies identifier case and the name a caller has may not be the spelling
     * the tree is showing.
     */
    treeMenu: (args) => {
      const module = args.module === undefined ? "" : String(args.module);
      const workbook = args.workbook === undefined ? "" : String(args.workbook);
      const same = (a: string | undefined, b: string): boolean =>
        (a ?? "").toLowerCase() === b.toLowerCase();

      if (!module && !workbook) {
        return { did: false, detail: "name a module or a workbook" };
      }

      const row = module
        ? [...document.querySelectorAll<HTMLElement>("[data-component]")].find((one) =>
          same(one.dataset.component, module) && (!workbook || same(one.dataset.workbook, workbook)))
        : [...document.querySelectorAll<HTMLElement>("[data-project]")].find((one) =>
          same(one.dataset.project, workbook));

      if (!row) {
        return {
          did: false,
          detail: module
            ? `the tree has no row for ${module}${workbook ? ` in ${workbook}` : ""}`
              + " (an unexpanded workbook has no rows)"
            : `the tree has no row for the workbook ${workbook}`,
        };
      }

      const box = row.getBoundingClientRect();
      row.dispatchEvent(new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        clientX: Math.round(box.left + 8),
        clientY: Math.round(box.top + box.height / 2),
      }));

      const labels = [...document.querySelectorAll(".menu-dropdown .menu-item")]
        .map((one) => (one.textContent ?? "").trim());

      return labels.length > 0
        ? { did: true, detail: labels.join(" | ") }
        : { did: false, detail: `right-clicking ${module || workbook} opened no menu` };
    },

    /**
     * Renames a module and everything that names it, the way the tree's Rename does.
     *
     * THE ONE STATE-CHANGING PRODUCT ACTION THAT NOTHING COULD DRIVE. It spans modules, it has an
     * engine behind it, and until now the only way to reach it was a person typing into
     * `window.prompt`, which no probe can answer. `component?action=rename` is not the same
     * operation: that is the fixture primitive, the bare `VBComponent.Name` setter, and it leaves
     * every reference to the old name pointing at a module that no longer answers to it.
     *
     * The prompt is the only thing skipped, because the prompt is where the human's answer comes
     * from. Everything after it is the code the menu item runs.
     *
     * Answers what the host answered: how many mentions were replaced, in how many modules, or the
     * refusal in the words the rename box would have shown.
     */
    renameModule: async (args) => {
      const module = String(args.module ?? "");
      const newName = String(args.newName ?? "");
      if (!module || !newName) {
        return { did: false, detail: "renameModule needs module and newName" };
      }

      const workbook = args.workbook === undefined || args.workbook === null
        ? null
        : String(args.workbook);

      const answer = await bridge.requestModuleRename(module, workbook, newName);
      if (answer.refused) {
        return { did: false, detail: answer.refused };
      }

      return {
        did: true,
        detail: `renamed ${module} to ${newName}: ${answer.replaced} mention(s) in `
          + `${answer.modules.length} module(s) [${answer.modules.join(", ")}]`,
      };
    },

    /**
     * Presses a workbook row's plus and reports the menu it opened.
     *
     * The plus is hidden until the row is hovered, and CSS :hover cannot be provoked from script,
     * so a probe reaching for it by pointer has to move a real mouse. It is a real button and the
     * tree listens for a real click, so this presses it directly: the reveal is a presentation
     * rule and pressing a control that is styled invisible still runs everything pressing it runs.
     * Whether it becomes VISIBLE on hover is a separate question and belongs to a probe that can
     * move a pointer.
     */
    treeAdd: (args) => {
      const workbook = String(args.workbook ?? "");
      if (!workbook) {
        return { did: false, detail: "no workbook given" };
      }

      const plus = [...document.querySelectorAll<HTMLElement>("[data-add-project]")].find((one) =>
        (one.dataset.addProject ?? "").toLowerCase() === workbook.toLowerCase());

      if (!plus) {
        return { did: false, detail: `the tree has no row for the workbook ${workbook}` };
      }

      plus.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));

      const labels = [...document.querySelectorAll(".menu-dropdown .menu-item")]
        .map((one) => (one.textContent ?? "").trim());

      return labels.length > 0
        ? { did: true, detail: labels.join(" | ") }
        : { did: false, detail: `the plus on ${workbook} opened no menu` };
    },

    /**
     * Opens the surface's own menu - the wrench at the head of the toolbar - and reports it.
     *
     * Its items come from the HOST, so unlike every other menu on this surface it is not there the
     * moment the button is pressed. Awaited rather than slept on, and reported empty rather than
     * reported as nothing when the host does not answer, because "the menu is empty" and "the menu
     * never arrived" are different failures.
     */
    menuBar: async () => {
      const button = document.querySelector<HTMLElement>("#menubar .menu-top");
      if (!button) {
        return { did: false, detail: "the surface has no menu" };
      }

      const options = { bubbles: true, cancelable: true };
      button.dispatchEvent(new PointerEvent("pointerdown", options));
      button.dispatchEvent(new PointerEvent("pointerup", options));

      const items = () => [...document.querySelectorAll(".menu-dropdown .menu-item")]
        .map((one) => (one.textContent ?? "").trim());

      for (let waited = 0; items().length === 0 && waited < 3000; waited += 25) {
        await new Promise((settle) => { setTimeout(settle, 25); });
      }

      const labels = items();
      return labels.length > 0
        ? { did: true, detail: labels.join(" | ") }
        : { did: false, detail: "the menu opened and the host sent no items in 3s" };
    },

    /**
     * Chooses an item of the open context menu by its label.
     *
     * On `pointerup`, which is what the menu listens for. A synthesised `click` lands on nothing
     * and reports having worked, the same trap as every other control in this product that arms
     * on a pointer event rather than on click.
     */
    chooseMenuItem: (args) => {
      const label = String(args.label ?? "");
      if (!label) {
        return { did: false, detail: "no label given" };
      }

      const rows = [...document.querySelectorAll<HTMLElement>(".menu-dropdown .menu-item")];
      if (rows.length === 0) {
        return { did: false, detail: "no context menu is open" };
      }

      // Prefix rather than equality: a destructive item wears a trailing ellipsis to say it will
      // ask first, and a caller should not have to spell the character to reach it.
      const wanted = rows.find((one) => {
        const text = (one.textContent ?? "").trim();
        return text === label || text.replace(/\.+$/, "") === label.replace(/\.+$/, "");
      });

      if (!wanted) {
        return {
          did: false,
          detail: `no item named ${label}; the menu offers `
            + rows.map((one) => (one.textContent ?? "").trim()).join(" | "),
        };
      }

      if (wanted.classList.contains("disabled")) {
        return { did: false, detail: `${label} is disabled` };
      }

      wanted.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, cancelable: true }));
      return { did: true, detail: `chose ${(wanted.textContent ?? "").trim()}` };
    },

    /**
     * Answers the box that asks before a component is removed. `answer` is remove or cancel.
     *
     * Its own action rather than a generic button-clicker, so a script cannot delete a module by
     * naming a button that happened to move. Cancel is what an unrecognised answer gets.
     */
    answerRemoveConfirm: (args) => {
      const answer = String(args.answer ?? "cancel").toLowerCase();
      const wanted = { remove: "Remove", cancel: "Cancel" }[answer];
      if (!wanted) {
        return { did: false, detail: `answer must be remove or cancel; got ${answer}` };
      }

      const button = [...document.querySelectorAll("#remove-confirm-buttons button")]
        .find((one) => (one.textContent ?? "").trim() === wanted) as HTMLButtonElement | undefined;

      if (!button) {
        return { did: false, detail: "no remove-component box is up" };
      }

      button.click();
      return { did: true, detail: `answered ${wanted}` };
    },

    /**
     * Opens the settings dialog, or with `key` and `value`, changes one setting through the page.
     *
     * Bare, it is the wrench's Settings item. With arguments it is what a control in that dialog
     * does when the developer touches it, which is a genuinely different path from the `settings`
     * ROUTE: the route replaces one field of the stored settings on the host side, the page posts
     * the whole object as it believes it to be. Only the second one can drop a field, and for
     * months it dropped syncEngine on every change.
     */
    settings: (args) => {
      const key = args.key === undefined ? null : String(args.key);
      if (key === null) {
        parts.openSettings();
        return { did: true, detail: "settings dialog opened" };
      }

      const raw = String(args.value ?? "");
      const value = raw === "true" ? true : raw === "false" ? false : Number.isNaN(Number(raw)) ? raw : Number(raw);

      return parts.changeSetting(key, value)
        ? { did: true, detail: `posted ${key}=${String(value)} the way the dialog does` }
        : { did: false, detail: `${key} is not a setting; try ${Object.keys(currentSettings()).join(", ")}` };
    },

    sponsors: () => {
      parts.openSponsors();
      return { did: true, detail: "sponsor dialog opened" };
    },

    /**
     * Presses a toolbar command by id: `act("toolbar", {command: "openSync"})`.
     *
     * The strip is where thirty commands live and nothing could reach any of them. `press` is a
     * keyboard key and the shim's `command` route is the NATIVE editor's command by name, so the
     * only way to the Object Browser, the sync dialog, the Panes menu or Help was to find the
     * element by its data-command attribute and click it - which three harness files did, and
     * which keeps passing after the button is renamed, disabled, or left out of the build.
     *
     * With no argument it answers the strip's contents instead, which is the other half of the
     * same question: a command that is not there cannot be pressed, and menu-bar.mjs was reading
     * the DOM to find that out.
     */
    toolbar: (args) => {
      const wanted = args.command === undefined ? null : String(args.command);
      if (wanted === null) {
        const shown = parts.toolbarCommands();
        return {
          did: true,
          detail: `${shown.length} command(s) on the strip`,
          data: { commands: shown },
        };
      }

      return parts.pressToolbar(wanted);
    },

    /**
     * Drives the Changes pane through its own controls: `{press}` clicks a named button (refresh,
     * snapshot, accept, or rail for the full-size card's own show-and-hide), `{round, module}`
     * opens one module's comparison exactly as clicking its row does, `{file}` points the pane at
     * another open file, and `{expand}` shows the comparison full size or closes it. `ui.changes`
     * is the read side, and carries `full`, `fullChoices`, `railUp` and `railWidth`.
     *
     * `{round, module, in: "full"}` clicks the row in the FULL-SIZE card's rail instead of the
     * pane's list. Two controls onto the same comparison, so a check that drove one has not said
     * anything about the other.
     *
     * There is no revert here because there is none in the pane: the log shows, and writing is
     * done through `module`, where it lands in the log like any other write.
     */
    /**
     * Drives the agent card: `{press}` clicks a named control (toggle, copy, close). `ui.agent`
     * is the read side and carries the instruction text as it stands.
     *
     * THE TOGGLE IS A REAL SWITCH, not a rehearsal - pressing it opens or shuts the api door for
     * this session and writes the choice to the settings file, exactly as a developer's click
     * does. A check that drives it is changing the machine's state, which is the point: there is
     * no other way to prove the door actually moves.
     */
    agentCard: (args) => {
      const card = agentDialogProbe();
      if (!card) {
        return { did: false, detail: "the agent card is not open; use command openAgent first" };
      }

      if (args.press !== undefined) {
        const control = String(args.press);
        return card.press(control)
          ? { did: true, detail: `${control} pressed` }
          : { did: false, detail: `no control named ${control}; use toggle, copy or close` };
      }

      return { did: false, detail: "nothing asked; pass press=toggle|copy|close" };
    },

    changesPane: (args) => {
      const pane = changesPaneProbe();
      if (!pane) {
        return { did: false, detail: "the Changes pane has not been built yet" };
      }

      if (args.press !== undefined) {
        const control = String(args.press);
        return pane.press(control)
          ? { did: true, detail: `${control} pressed` }
          : { did: false, detail: `no control named ${control}; use refresh, snapshot, accept or rail` };
      }

      if (args.module !== undefined) {
        const module = String(args.module);
        const round = Number(args.round ?? 0);
        const where = String(args.in ?? "pane") === "full" ? "full" : "pane";
        return pane.show(round, module, where)
          ? { did: true, detail: `showing ${module} from round ${round}, from the ${where}` }
          : { did: false, detail: `the ${where} has no row for ${module} in round ${round}` };
      }

      if (args.file !== undefined) {
        const file = String(args.file);
        return pane.chooseFile(file)
          ? { did: true, detail: `pointed at ${file} through the select's own change` }
          : { did: false, detail: `${file} is not one of the open files the pane was given` };
      }

      if (args.expand !== undefined) {
        const open = args.expand !== false && args.expand !== "false";
        return pane.expand(open)
          ? { did: true, detail: open ? "the comparison is up full size" : "the comparison was closed" }
          : { did: false, detail: open ? "there is no comparison on screen to show" : "nothing was up" };
      }

      return {
        did: false,
        detail: "nothing asked; pass press, file, expand, or module (with round, and in=full for the card's rail)",
      };
    },

    /**
     * Drives the OPEN sync dialog through its own controls: `{press}` clicks a named button
     * (apply, close, all, none, export, import), `{tick, on}` a row's checkbox by file name,
     * `{folder}` types the path through the input's own change event, and `{project}` points it
     * at another open project through the select's own change event. `ui.sync` is the read side,
     * and it carries `project` and `projects` so a driver can see the choice as well as make it.
     * Before this pair, driving the dialog meant querySelector against its private DOM from a
     * harness file, which tests the selector as much as the dialog.
     */
    syncDialog: (args) => {
      const dialog = syncDialogProbe();
      if (!dialog) {
        return { did: false, detail: 'the sync dialog is not open; act("toolbar", {command: "openSync"}) summons it' };
      }

      if (args.press !== undefined) {
        const control = String(args.press);
        return dialog.press(control)
          ? { did: true, detail: `${control} pressed` }
          : { did: false, detail: `no control named ${control}; use apply, close, all, none, export or import` };
      }

      if (args.tick !== undefined) {
        const file = String(args.tick);
        const on = args.on !== false && args.on !== "false";
        return dialog.tick(file, on)
          ? { did: true, detail: `${file} ${on ? "ticked" : "unticked"}` }
          : { did: false, detail: `no row named ${file}` };
      }

      if (args.folder !== undefined) {
        dialog.setFolder(String(args.folder));
        return { did: true, detail: "folder set through the input's own change" };
      }

      if (args.project !== undefined) {
        const name = String(args.project);
        return dialog.chooseProject(name)
          ? { did: true, detail: `pointed at ${name} through the select's own change` }
          : { did: false, detail: `${name} is not one of the open projects the dialog was given` };
      }

      return { did: false, detail: "nothing asked; pass press, tick (with on), folder or project" };
    },

    /**
     * Edits a row of the Properties panel, the way touching its control does.
     *
     * **THIS CHANGES REAL COMPONENT STATE.** The panel is a view over the object model's own
     * property bag, and the host applies the write immediately: `(Name)` renames the component,
     * and a worksheet's rows are the worksheet's. There is no undo behind it.
     *
     * The panel shows whichever component is selected, so `act("selectComponent")` or a click in
     * the tree comes first; `ui.properties` says which one is up.
     */
    editProperty: (args) => {
      const name = String(args.name ?? "").trim();
      const value = args.value === undefined ? null : String(args.value);

      if (!name || value === null) {
        return { did: false, detail: "editProperty needs name and value" };
      }

      const shown = parts.properties();
      if (!shown.component) {
        return { did: false, detail: "no component is selected, so the panel is showing nothing" };
      }

      if (!parts.editProperty(name, value)) {
        const writable = shown.rows.filter((row) => row.writable).map((row) => row.name);
        return {
          did: false,
          detail: `${shown.component} has no writable property named ${name}; it offers `
            + (writable.length > 0 ? writable.join(", ") : "none"),
        };
      }

      /*
       * A ROW ON A DESIGNER TAB IS ANSWERED FROM THE DOCUMENT, because no host answer is coming.
       *
       * Those edits stopped going to the host on 2026-08-17 (task #68): the panel now writes the
       * tab's document and Ctrl+S applies it, which is what gives them an undo step. This waited
       * for the host's republish regardless, so it sat out its two seconds and then called every
       * one of them refused - a verdict about a round trip that no longer happens. The document
       * IS the outcome here, so the document is what gets read back.
       */
      const active = workspace.activeDocument();
      const owning = active?.face === "design" ? designer.viewFor(active.module, active.project ?? null) : null;
      if (owning?.spells(shown.component, name)) {
        const spelled = new RegExp(`\\b${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*=\\s*"([^"]*)"`, "i");
        const held = spelled.exec(owning.markupText())?.[1] ?? null;
        return {
          did: true,
          detail: held === null
            ? `${shown.component}.${name} is at its default, so the document leaves it unspoken`
            : `${shown.component}.${name} is ${JSON.stringify(held)} in the document; Ctrl+S applies it`,
        };
      }

      // AND THE HOST'S ANSWER, not the page's request.
      //
      // The controls post the edit and do not wait; the host applies it, and on a refusal - an
      // illegal identifier, a property the object model will not take - it says so and republishes
      // the row with the value it still holds. So the panel is right a moment later and an act
      // that returned here would have reported every refused write as a success. It did, for the
      // first twenty minutes this existed, and the suite caught it: "set to 'not a legal name'"
      // for a name VBA cannot hold.
      // WAITING FOR THE HOST'S REPUBLISH, not for the row to hold what we just put in it. The
      // control sets the row as it posts, the way a responsive control does, so reading the value
      // back proves only that we asked. `round` changes when the host answers - applied or
      // refused, it republishes either way - and only then is the value worth comparing.
      const asked = shown.round;

      // A REFUSAL IS THE HOST SAYING SO, not the row failing to echo the request back.
      //
      // This compared the republished value against the text that was sent, which held only
      // while the panel showed values in exactly the spelling a caller typed. It does not any
      // more: a property whose type library names its values answers `fmCycleAllForms` to a
      // written `0`, and the honest write read as a refusal for two seconds and then lied
      // (2026-08-15). The host complains through the status line when it will not take a
      // value, so the notice is the signal - and the detail reports what the row HOLDS, which
      // is the answer the caller actually wanted.
      const saidBefore = parts.statusNotice();

      const settled = (): Promise<ActResult> => new Promise((answer) => {
        const deadline = Date.now() + 2000;

        const look = (): void => {
          const now = parts.properties();
          const row = now.rows.find((one) => one.name.toLowerCase() === name.toLowerCase());
          const said = parts.statusNotice();

          if (now.round !== asked && (said === saidBefore || !said)) {
            answer({
              did: true,
              detail: `${now.component}.${row?.name ?? name} is ${JSON.stringify(row?.value ?? value)}`,
            });
            return;
          }

          if (Date.now() >= deadline) {
            const said = parts.statusNotice();
            answer({
              did: false,
              detail: `${shown.component}.${name} was refused; it holds `
                + `${JSON.stringify(row?.value ?? null)}`
                + (said ? `. The surface said: ${said}` : ""),
            });
            return;
          }

          setTimeout(look, 60);
        };

        look();
      });

      return settled();
    },

    /**
     * The colour picker, through the swatch and the swatch's own click. `property` names the
     * row, and `choose` - a `#rrggbb` from the palette or a system colour's NAME - presses that
     * choice; without it the picker is left standing, which is how a probe reads what it offers.
     *
     * The picker is a page widget over a host value, so this act stops at the page: the write it
     * causes is an ordinary property edit, and `editProperty`'s own wait is what proves the host
     * took it.
     */
    colourPicker: (args) => {
      const property = typeof args.property === "string" ? args.property : null;
      const choose = typeof args.choose === "string" ? args.choose : null;

      if (property) {
        const swatch = document.querySelector<HTMLElement>(
          `.prop-row[data-property="${CSS.escape(property)}"] .prop-swatch`);
        if (!swatch) {
          const shown = parts.properties();
          const colours = shown.rows.filter((row) => row.swatch).map((row) => row.name);
          return {
            did: false,
            detail: `no colour row named ${property}; the panel offers `
              + (colours.length > 0 ? colours.join(", ") : "none"),
          };
        }

        swatch.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, cancelable: true }));
      }

      const state = colourPickerState();
      if (!state) {
        return { did: false, detail: "no color picker is open" };
      }

      if (choose === null) {
        return {
          did: true,
          detail: `${state.property} is open on ${state.value}: `
            + `${state.palette} palette colour(s), ${state.system} system colour(s)`,
          data: state,
        };
      }

      const picked = pickColour(choose);
      return picked === null
        ? { did: false, detail: `the picker offers no colour spelled ${choose}`, data: state }
        : { did: true, detail: `picked ${picked} for ${state.property}`, data: { ...state, picked } };
    },

    // Every page dialog closes on Escape, and none of them expose a handle. This is the one
    // action that has to go through the DOM, and it goes through the key the user would press.
    closeDialogs: () => {
      const up = dialogsUp();
      if (up.length === 0) {
        return { did: false, detail: "no dialog is up" };
      }
      for (const _ of up) {
        document.dispatchEvent(new KeyboardEvent("keydown", {
          key: "Escape", code: "Escape", bubbles: true, cancelable: true,
        }));
      }
      return { did: true, detail: `escaped ${up.map((one) => one.title).join(", ")}` };
    },

    key: sendKey,

    focusEditor: () => {
      workspace.activeEditor().focus();
      return { did: true, detail: "the active editor has focus" };
    },

    /**
     * Find and replace, driven the way the box is.
     *
     * `find` types the query and fires the input event the widget listens for, rather than
     * setting the value and hoping: an input whose value is assigned in script raises no event
     * and the widget would never search. That is the same class of miss as clicking a control
     * that arms on pointerdown.
     */
    search: (args) => {
      const shut = flag(args.close, false);
      if (shut) {
        parts.search.close();
        return { did: true, detail: "the find box is closed" };
      }

      const query = args.query === undefined ? null : String(args.query);
      if (query === null) {
        parts.search.open(args.scope ? { scope: String(args.scope) } : undefined);
        return { did: true, detail: "the find box is open" };
      }

      parts.search.find(query, {
        ...(args.scope ? { scope: String(args.scope) } : {}),
        matchCase: flag(args.matchCase, false),
        wholeWord: flag(args.wholeWord, false),
      });

      // TYPING IS NOT SEARCHING outside module scope, and this used to stop here and answer as
      // though it were. `find` raises an `input` event; the only handler for it searches when the
      // scope is "module" and does nothing otherwise. So a project search typed a query, searched
      // nothing, and answered did:true - and the field a probe would then check, `matches`, is
      // structurally 0 for every non-module scope, so the two agreed and the feature looked fine.
      //
      // `run` is the press that follows the typing, and it is the same method the button's own
      // click listener calls. Without it, Replace All - a text rewrite across every module of a
      // project, and the most destructive thing on this surface - could not be triggered at all.
      if (args.replacement !== undefined) {
        parts.search.setReplacement(String(args.replacement));
      }

      const run = args.run === undefined ? null : String(args.run).toLowerCase();
      if (run === null) {
        return { did: true, detail: `typed ${JSON.stringify(query)}; nothing was run` };
      }

      const press: Record<string, () => void> = {
        find: () => parts.search.runFindAll(),
        findall: () => parts.search.runFindAll(),
        next: () => parts.search.goToNextMatch(),
        previous: () => parts.search.goToPreviousMatch(),
        replace: () => parts.search.runReplaceCurrent(),
        replaceall: () => parts.search.runReplaceAll(),
      };

      const pressed = press[run];
      if (!pressed) {
        return {
          did: false,
          detail: `run must be findAll, next, previous, replace or replaceAll; got ${run}`,
        };
      }

      pressed();
      return { did: true, detail: `${run} for ${JSON.stringify(query)}` };
    },

    /*
     * THE LANGUAGE FEATURES, asked the way the EDITOR asks them.
     *
     * Each calls the very provider object monaco calls, with the arguments monaco passes. That
     * is deliberate and it is the whole point: every one of these providers begins by refusing
     * to answer for anything but the host-active module, and an api that skipped that gate would
     * answer for a module the developer's editor stays silent on.
     *
     * The first version DID skip it, calling the bridge request underneath, and it reported
     * hover healthy through an entire session in which hover was dead on screen because the host
     * believed no module was active (2026-08-08). The coverage agreed with the code and
     * disagreed with the product.
     *
     * An api action must report what the same action through the UI would; where it cannot, it
     * says so.
     */

    hover: async (args) => {
      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      const found = await parts.providers.hover.provideHover(
        where.model, where.position, NO_CANCEL, EMPTY_CONTEXT);
      return {
        did: Boolean(found),
        detail: found ? "hover answered" : "the editor would show nothing here",
        data: found ?? null,
      };
    },

    completions: async (args) => {
      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      const answer = await parts.providers.completion.provideCompletionItems(
        where.model,
        where.position,
        { triggerKind: 0, triggerCharacter: args.trigger === undefined ? undefined : String(args.trigger) } as never,
        NO_CANCEL);

      const items = answer?.suggestions ?? [];
      return {
        did: items.length > 0,
        detail: `${items.length} completion(s)`,
        data: items.map((one) => ({ label: one.label, kind: one.kind, detail: one.detail })),
      };
    },

    signature: async (args) => {
      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      const answer = await parts.providers.signature.provideSignatureHelp(
        where.model, where.position, NO_CANCEL,
        { triggerKind: 1, isRetrigger: false } as never);

      return {
        did: Boolean(answer),
        detail: answer ? "signature help answered" : "the editor would show nothing here",
        data: answer?.value ?? null,
      };
    },

    quickFixes: async (args) => {
      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      // The whole word, so a fix attached to an identifier is offered from anywhere inside it,
      // which is what the lightbulb does for a caret sitting in the middle of one.
      const word = where.model.getWordAtPosition(where.position);
      const range = word
        ? new monacoApi.Range(
          where.position.lineNumber, word.startColumn, where.position.lineNumber, word.endColumn)
        : new monacoApi.Range(
          where.position.lineNumber, where.position.column,
          where.position.lineNumber, where.position.column);

      const answer = await parts.providers.codeAction.provideCodeActions(
        where.model, range, { trigger: 1, only: undefined } as never, NO_CANCEL);

      const actions = answer?.actions ?? [];
      return {
        did: actions.length > 0,
        detail: `${actions.length} quick fix(es)`,
        data: actions.map((one) => ({ title: one.title, kind: one.kind, isPreferred: one.isPreferred })),
      };
    },

    /**
     * A language feature timed INSIDE the page, which is the only place the number is honest.
     *
     * Everything else measures across the door, and the door collects a promise by polling - so
     * an async route carries a floor of tens of milliseconds whatever the feature cost. That
     * floor is most of every figure taken from outside, and it hid a whole scaling curve behind
     * a flat line until it was noticed (2026-08-08).
     *
     * This runs the provider n times here and reports the distribution, so one door round trip
     * covers every sample and the door's cost is amortised to nothing. What comes back is what
     * the developer waits for: the page's own work plus the shim and the analyzer behind it,
     * with nothing of the harness in it.
     */
    timeFeature: async (args) => {
      const which = String(args.what ?? "hover");
      const runs = Math.max(1, Math.min(50, Number(args.n ?? 10)));
      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      const call = {
        hover: () => parts.providers.hover.provideHover(where.model, where.position, NO_CANCEL, EMPTY_CONTEXT),
        completions: () => parts.providers.completion.provideCompletionItems(
          where.model, where.position, { triggerKind: 0 } as never, NO_CANCEL),
        definition: () => parts.providers.definition.provideDefinition(where.model, where.position, NO_CANCEL),
        signature: () => parts.providers.signature.provideSignatureHelp(
          where.model, where.position, NO_CANCEL, { triggerKind: 1, isRetrigger: false } as never),
      }[which];

      if (!call) {
        return { did: false, detail: `what must be hover, completions, definition or signature; got ${which}` };
      }

      const samples: number[] = [];
      let answered = 0;

      for (let run = 0; run < runs; run++) {
        const began = performance.now();
        const answer = await call();
        samples.push(Math.round((performance.now() - began) * 100) / 100);
        if (answer) { answered++; }
      }

      const ordered = [...samples].sort((a, b) => a - b);
      return {
        did: true,
        detail: `${which} x${runs}, ${answered} answered`,
        data: {
          what: which,
          runs,
          answered,
          minMs: ordered[0],
          medianMs: ordered[ordered.length >> 1],
          p95Ms: ordered[Math.min(ordered.length - 1, Math.floor(ordered.length * 0.95))],
          maxMs: ordered[ordered.length - 1],
          samplesMs: samples,
        },
      };
    },

    /**
     * Find All References, as the dialog lists them.
     *
     * The dialog is opened by the same lookup, so this cannot show a different set from what the
     * developer sees - which is the point of there being one function rather than two.
     */
    references: async (args) => {
      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      const at = { line: where.position.lineNumber, column: where.position.column };

      // open=1 LEAVES THE DIALOG STANDING, which is what Shift+F12 does and what the data-only
      // default does not: the dialog's rendering is the feature (it draws a module with no tab
      // open, which monaco's peek cannot), so a lookup that returns rows and shows nothing
      // exercises everything except the thing the feature exists for. Both forms share one
      // referencesAt lookup, so the dialog cannot list a set the data form would not.
      if (flag(args.open, false)) {
        const opened = await parts.openReferences(at);
        return opened
          ? { did: true, detail: `the references dialog is open at ${at.line}:${at.column}` }
          : { did: false, detail: "the editor would not answer here; the host-active module is another" };
      }

      const answer = await parts.referencesAt(at);
      if (!answer) {
        return { did: false, detail: "the editor would not answer here; the host-active module is another" };
      }

      return {
        did: answer.found.length > 0,
        detail: `${answer.found.length} reference(s) to ${answer.word}`,
        data: answer.found,
      };
    },

    /**
     * Go to Definition, through the provider F12 goes through.
     *
     * Answers the locations the editor would navigate to, so a probe can assert WHERE without
     * moving the caret - which matters because moving it is what cancels a symbol-navigation
     * request in the first place (lessons, finding 10).
     */
    definition: async (args) => {
      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      const found = await parts.providers.definition.provideDefinition(
        where.model, where.position, NO_CANCEL);

      const locations = Array.isArray(found) ? found : (found ? [found] : []);
      return {
        did: locations.length > 0,
        detail: `${locations.length} definition(s)`,
        data: locations.map((one) => ({
          uri: String((one as monaco.languages.Location).uri ?? ""),
          range: (one as monaco.languages.Location).range,
        })),
      };
    },

    /**
     * RENAME, the flagship, and the one feature here that rewrites the developer's code.
     *
     * It had no api at all until now: a change that edits every module using a symbol, across a
     * workbook, with nothing able to drive it but a hand on F2. `newName` is required, and the
     * answer carries the provider's own refusal when it declines - which is what the rename box
     * shows the developer, word for word.
     *
     * This DOES change state. It is the same call F2 makes, so what it leaves behind is what the
     * developer would have: the host applies the edits and the surface republishes. Undo it with
     * `undoRename`, which is the same path the editor's own Undo Rename takes.
     */
    rename: async (args) => {
      const newName = String(args.newName ?? "");
      if (!newName) { return { did: false, detail: "newName is required" }; }

      const where = positionFrom(args);
      if (!where) { return { did: false, detail: "nothing open, or no such word" }; }

      const answer = await parts.providers.rename.provideRenameEdits(
        where.model, where.position, newName, NO_CANCEL);

      const refused = (answer as { rejectReason?: string } | null)?.rejectReason;
      return {
        did: !refused,
        detail: refused ?? `renamed to ${newName}`,
        data: { refused: refused ?? null },
      };
    },

    /**
     * Undo and redo, the two most state-affecting keys a developer presses and the two that had
     * no way of being driven.
     *
     * `undoRename` existed, and only undoes a rename. Everything else a probe did to a module
     * could be made and never taken back the way a person takes it back, so the one operation
     * most likely to leave the workbook, the surface and the analyzer holding three different
     * texts was the one operation no check ever performed.
     *
     * Triggered by name rather than looked up: undo and redo are built into the editor rather
     * than registered as actions, so `getAction("undo")` finds nothing. This is the same path the
     * toolbar button and Ctrl+Z take.
     */
    undo: (args) => {
      const editor = workspace.activeEditor();
      const id = args.redo === undefined ? "undo" : flag(args.redo, false) ? "redo" : "undo";
      const times = Math.max(1, Math.min(64, Number(args.times ?? 1) || 1));

      editor.focus();
      for (let step = 0; step < times; step += 1) {
        editor.trigger("xlide", id, null);
      }

      const at = editor.getPosition();
      return {
        did: true,
        detail: `${id} ×${times}, caret at ${at?.lineNumber}:${at?.column}`,
        data: { line: at?.lineNumber ?? null, column: at?.column ?? null },
      };
    },

    /**
     * Docks a pane on a side, THROUGH THE METHOD A DROP CALLS.
     *
     * The docking gestures had no coverage of any kind: `layout()` reported the arrangement and
     * `resetLayout()` restored it, and nothing in between could move a thing, so the surface was
     * drivable only by hand. Synthesising the drag was never the answer: the drop arms on
     * pointerdown and completes on pointerup against a compass that follows the pointer, so a
     * synthetic sequence would test the synthesiser.
     *
     * Answers the arrangement afterwards, so a caller can see where the pane landed rather than
     * having to ask separately.
     */
    dock: (args) => {
      const pane = String(args.pane ?? args.name ?? "");
      const side = String(args.side ?? args.to ?? "");

      if (!pane) {
        return { did: false, detail: "no pane given; pass pane: \"Immediate\"" };
      }

      if (side !== "left" && side !== "right" && side !== "top" && side !== "bottom") {
        return { did: false, detail: `side must be left, right, top or bottom; got ${JSON.stringify(side)}` };
      }

      const moved = parts.panes.moveTo(pane, side);
      return moved
        ? { did: true, detail: `${pane} docked ${side}`, data: { pane, side } }
        : { did: false, detail: `no pane named ${pane}` };
    },

    /**
     * Backspace, through the command the key is bound to.
     *
     * The shim's `type` route cannot express one: it drives `trigger("keyboard", "type")`, which
     * only ever inserts. Backspace is `deleteLeft`, and it is worth being able to send because in
     * a line's leading whitespace it does not delete one character. With `useTabStops` on it
     * takes back a whole indent level, which is the behaviour that makes indenting with spaces
     * feel like indenting with tabs, and the only reason removing the tabs setting was tolerable.
     *
     * `times` presses it more than once, because the interesting cases are the second and third
     * press: level two to level one, level one to the margin, and then the margin holding.
     */
    /**
     * Presses a key, as the keyboard would.
     *
     * THE GAP THIS CLOSES. `type` inserts a string, and inserting one is not the same as typing
     * it: Monaco applies its enter rules to a newline TYPED as a single character and not to one
     * that arrives inside a longer string. So every behaviour that hangs off Enter - auto-indent,
     * smart Enter's block layout, comment continuation, comment-spacing mirroring - could not be
     * driven from here at all, and the settings that govern them had no live test. A fix to the
     * indentation rules on 2026-08-09 had to ship reasoned rather than measured for exactly this.
     *
     * `backspace` below is the same idea, added when it was needed and never generalised. It stays
     * because probes name it, and it now goes through here.
     *
     * NOT `key`, WHICH IS A DIFFERENT THING and the reason this is named apart. `key` dispatches a
     * synthetic KeyboardEvent at the document to exercise the chords this product binds there,
     * Ctrl+W above all. Monaco does not act on synthesised events, so `key` cannot type, and this
     * cannot test a chord. Two capabilities, two names, rather than one name that does half of
     * each depending on the argument.
     *
     *   act("press", { key: "Enter" })
     *   act("press", { key: "Tab", times: 2 })
     *
     * Answers the caret and the line it ended on, which is what a caller is checking.
     */
    press: (args) => {
      const editor = workspace.activeEditor();
      const times = Math.max(1, Math.min(64, Number(args.times ?? 1) || 1));
      const wanted = String(args.key ?? "");

      // A newline TYPED, one character at a time, which is what makes the enter rules run. The
      // others are editor commands, because that is how the keybindings reach them.
      const press: Record<string, () => void> = {
        Enter: () => editor.trigger("keyboard", "type", { text: "\n" }),
        Tab: () => editor.trigger("keyboard", "tab", null),
        Backspace: () => editor.trigger("keyboard", "deleteLeft", null),
        Delete: () => editor.trigger("keyboard", "deleteRight", null),
        Escape: () => editor.trigger("keyboard", "cancelSelection", null),
      };

      const stroke = press[wanted];
      if (stroke === undefined) {
        return {
          did: false,
          detail: `no key '${wanted}'; this drives ${Object.keys(press).join(", ")}`,
          data: null,
        };
      }

      editor.focus();
      for (let at = 0; at < times; at += 1) { stroke(); }

      const at = editor.getPosition();
      return {
        did: true,
        detail: `${wanted} ×${times}, caret at ${at?.lineNumber}:${at?.column}`,
        data: {
          line: at?.lineNumber ?? null,
          column: at?.column ?? null,
          text: at ? editor.getModel()?.getLineContent(at.lineNumber) ?? null : null,
        },
      };
    },

    /**
     * Types text into the page's editor, without the host seeing it first.
     *
     * `type` on the shim goes through the HOST's keyboard pipeline, and the host normalises what
     * it takes: it respells keywords canonically as they arrive. So text written or typed through
     * the door is already canonical by the time the page has it, and any page behaviour that acts
     * on non-canonical text cannot be reached from outside at all. `formatCanonicalKeywords` is
     * exactly that: a setting whose whole job is respelling keywords, which could not be observed
     * because nothing could hand the page a keyword spelled the other way (2026-08-09).
     *
     * The caret goes where the text ends, as typing does. Multi-line text is inserted rather than
     * typed line by line, so this does NOT run the enter rules: `press` is for that, and the two
     * are separate for the same reason `press` and `key` are.
     */
    insert: (args) => {
      const editor = workspace.activeEditor();
      const text = String(args.text ?? "");
      if (text.length === 0) {
        return { did: false, detail: "no text given; pass text: \"...\"", data: null };
      }

      editor.focus();
      editor.trigger("keyboard", "type", { text });

      const at = editor.getPosition();
      return {
        did: true,
        detail: `inserted ${text.length} character(s), caret at ${at?.lineNumber}:${at?.column}`,
        data: {
          line: at?.lineNumber ?? null,
          column: at?.column ?? null,
          text: at ? editor.getModel()?.getLineContent(at.lineNumber) ?? null : null,
        },
      };
    },

    /**
     * Backspace, THE KEY - not the editor's deleteLeft command underneath it.
     *
     * Smart Backspace is bound to the key with a when-clause and no command id, which is the
     * only way to rebind a key here without taking it away from every other editor on the page.
     * Nothing can invoke it by name, so driving `deleteLeft` reached the stock command and the
     * product's own rules - the continued comment's marker, a blank line's indent - were
     * untestable from outside, while the checks that used this action read as if they covered
     * them (found 2026-08-21, adding the blank-line rule).
     *
     * A synthesised keydown on the editor's textarea DOES reach the keybinding service, which
     * is where a keybinding lives; it is only the typing path that ignores synthetic events,
     * because that arrives as input rather than as a key. When a rule declines, its handler
     * falls through to deleteLeft, so this stays exactly as useful for the ordinary cases.
     */
    backspace: (args) => {
      const editor = workspace.activeEditor();
      const times = Math.max(1, Math.min(64, Number(args.times ?? 1) || 1));

      editor.focus();
      const area = editor.getDomNode()?.querySelector("textarea");
      for (let press = 0; press < times; press += 1) {
        if (area) {
          area.dispatchEvent(new KeyboardEvent("keydown", {
            key: "Backspace", code: "Backspace", keyCode: 8, which: 8,
            bubbles: true, cancelable: true,
          }));
        } else {
          editor.trigger("keyboard", "deleteLeft", null);
        }
      }

      const at = editor.getPosition();
      return {
        did: true,
        detail: `Backspace ×${times}, caret at ${at?.lineNumber}:${at?.column}`,
        data: {
          line: at?.lineNumber ?? null,
          column: at?.column ?? null,
          text: at ? editor.getModel()?.getLineContent(at.lineNumber) ?? null : null,
        },
      };
    },

    /**
     * Format Module, or Format Selection: the editor's own formatting actions.
     *
     * The third consumer of the indent settings, after typing and smart Enter, and the one where
     * a disagreement is most visible: formatting rewrites the whole module at once.
     */
    format: (args) => {
      const whole = !flag(args.selection, false);
      const id = whole ? "editor.action.formatDocument" : "editor.action.formatSelection";
      const found = workspace.activeEditor().getAction(id);

      if (!found) {
        return { did: false, detail: `${id} is not registered on this editor` };
      }

      // Awaited: formatting replaces the whole model, and a caller reading the text straight
      // after an un-awaited run reads what was there before.
      return found.run().then(() => ({ did: true, detail: `ran ${id}` }));
    },

    /**
     * A bookmark on the caret's line, or a hop between them: the editor's own actions.
     *
     * The argument is `which`, NOT `do`. The route uses `do` to pick the action, so an action
     * that also took `do` had its own value overwrite the selector: `act("bookmark", {do:
     * "toggle"})` became `do=bookmark&do=toggle` and the door answered "unknown action toggle".
     * The client refuses a `do` argument now, so the collision cannot be made silently again.
     */
    bookmark: (args) => {
      const which = String(args.which ?? "toggle").toLowerCase();
      const action = {
        toggle: "xlide.bookmark.toggle",
        next: "xlide.bookmark.next",
        previous: "xlide.bookmark.previous",
        clear: "xlide.bookmark.clearAll",
      }[which];

      if (!action) {
        return { did: false, detail: `which must be toggle, next, previous or clear; got ${which}` };
      }

      const editor = workspace.activeEditor();
      const found = editor.getAction(action);
      if (!found) {
        return { did: false, detail: `${action} is not registered on this editor` };
      }

      void found.run();
      return { did: true, detail: `ran ${action}` };
    },

    /**
     * Any editor action, by the id it is registered under.
     *
     * The actions this product adds are the run() bodies of `editor.addAction` calls, and several
     * of them are reachable ONLY from a context menu or a key: Undo Rename is one, and the door
     * covered it by calling the shim's undo directly, which proves the operation is reversible and
     * not that the menu item works. Anything registered can be driven from here, so a feature
     * whose only surface is an action does not need a route of its own to be tested.
     *
     * Awaited when the action answers a promise, for the reason `format` gives: a caller reading
     * the text straight after an un-awaited run reads what was there before.
     *
     * Not every command is an action. Undo and redo are registered as commands, so `getAction`
     * finds nothing for them - see the note above `undo` for why, and use that action instead.
     */
    editorAction: (args) => {
      const id = String(args.id ?? "").trim();
      if (!id) {
        return { did: false, detail: "editorAction needs an id" };
      }

      const found = workspace.activeEditor().getAction(id);
      if (!found) {
        return { did: false, detail: `${id} is not registered on this editor` };
      }

      return Promise.resolve(found.run()).then(() => ({ did: true, detail: `ran ${id}` }));
    },
  };

  const act = (name: string, args?: Record<string, unknown>): ActAnswer => {
    const action = actions[name];
    if (!action) {
      return { did: false, detail: `unknown action ${name}; try ${Object.keys(actions).sort().join(", ")}` };
    }

    const named = (error: unknown): ActResult => ({
      did: false,
      detail: `${name} threw: ${error instanceof Error ? error.message : String(error)}`,
    });

    try {
      const answer = action(args ?? {});

      // A rejected promise walks straight past a try/catch, and an action is allowed to be
      // async now. Without this the door would carry a bare rejection with no stack and no
      // action name, which is the half worth having.
      return answer instanceof Promise ? answer.catch(named) : answer;
    } catch (error) {
      // Reported rather than thrown: the door carries a script error as a bare message with no
      // stack, and "which action failed" is the half that goes missing.
      return named(error);
    }
  };

  (globalThis as { xlideUi?: unknown }).xlideUi = {
    /**
     * `state()` for the whole surface; `state(line, column)` or `state(null, null, word)` to
     * also answer what is AT a position: its colour as painted and the squiggles covering it.
     */
    state: (line?: number, column?: number, word?: string) => {
      if (word !== undefined && word !== null && word !== "") {
        const offset = offsetFrom({ word });
        const model = workspace.activeEditor().getModel();
        const position = offset !== null && model ? model.getPositionAt(offset) : null;
        return state(position ? { lineNumber: position.lineNumber, column: position.column } : undefined);
      }

      return state(
        typeof line === "number" && line >= 1
          ? { lineNumber: line, column: typeof column === "number" && column >= 1 ? column : 1 }
          : undefined);
    },
    act,
    actions: () => Object.keys(actions).sort(),
  };
}
