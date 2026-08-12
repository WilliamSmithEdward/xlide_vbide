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
 * Reached from the debug api's `ui` and `act` routes, and from a devtools console. Shipped in every
 * build: it is a few hundred bytes of read-only reporting over objects the page already holds, and
 * a door that is only there in Debug is a door nobody trusts in the build that matters.
 */

import type * as monaco from "monaco-editor";
// The runtime object, not only its types: getModelMarkers is what draws the squiggles, and a
// type-only import cannot call it.
import * as monacoApi from "monaco-editor/editor/editor.api.js";

import type { EditorBridge } from "./bridge.js";
import type { Explorer, ExplorerSnapshot } from "./explorer.js";
import type { Workspace, WorkspaceSnapshot } from "./workspace.js";
import { currentSettings } from "./settings.js";

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
    rows: { name: string; value: string; writable: boolean; boolean: boolean }[];
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
  openSettings(): void;
  openSponsors(): void;
  /** What the status line is showing, read from the element the render writes. */
  statusNotice(): string;
  /**
   * The Properties panel: what it is showing, and the one way to change a value.
   *
   * It writes real component state through the object model's own setter - renaming a module is
   * the "(Name)" row - and until 2026-08-11 nothing in the api could drive it, read it, or even
   * name it. It was the only user-visible surface with no presence in either direction.
   */
  properties(): { component: string; kind: string; round: number; rows: { name: string; value: string; writable: boolean; boolean: boolean }[] };
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
  const { workspace, explorer, bridge } = parts;

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

  const state = (at?: { lineNumber: number; column: number }): UiSnapshot => ({
    workspace: workspace.snapshot(),
    explorer: explorer.treeState(),
    panes: parts.panes.list().map(({ name, title, open }) => ({ name, title, open })),
    dialogs: dialogsUp(),
    waiting: bridge.waitingOn(),
    focus: editorFocus(),
    settings: { ...currentSettings() } as unknown as Record<string, unknown>,
    emptyViewShown: workspace.emptyViewShown(),
    properties: parts.properties(),
    statusNotice: parts.statusNotice(),
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
     */
    activate: (args) => {
      const module = String(args.module ?? "");
      const project = args.project === undefined || args.project === null
        ? null
        : String(args.project);
      if (!module) {
        return { did: false, detail: "no module given" };
      }

      workspace.pickTab({ module, project });

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
       * waited for it to happen.
       */
      const showing = () => workspace.activeEditor().getModel()?.uri.path.split("/").pop() ?? null;
      const wanted = module.toLowerCase();

      return (async () => {
        const deadline = Date.now() + 4000;
        while (Date.now() < deadline) {
          if ((showing() ?? "").toLowerCase() === wanted) {
            return { did: true, detail: `showing ${module}` };
          }
          await new Promise((resolve) => setTimeout(resolve, 60));
        }

        return { did: false, detail: `asked for ${module}; the page is showing ${showing() ?? "nothing"}` };
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

      const settled = (): Promise<ActResult> => new Promise((answer) => {
        const deadline = Date.now() + 2000;

        const look = (): void => {
          const now = parts.properties();
          const row = now.rows.find((one) => one.name.toLowerCase() === name.toLowerCase());

          if (now.round !== asked && row?.value === value) {
            answer({ did: true, detail: `${now.component}.${row.name} is ${JSON.stringify(value)}` });
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

      const answer = await parts.referencesAt({
        line: where.position.lineNumber,
        column: where.position.column,
      });

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

    backspace: (args) => {
      const editor = workspace.activeEditor();
      const times = Math.max(1, Math.min(64, Number(args.times ?? 1) || 1));

      editor.focus();
      for (let press = 0; press < times; press += 1) {
        editor.trigger("keyboard", "deleteLeft", null);
      }

      const at = editor.getPosition();
      return {
        did: true,
        detail: `deleteLeft ×${times}, caret at ${at?.lineNumber}:${at?.column}`,
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
