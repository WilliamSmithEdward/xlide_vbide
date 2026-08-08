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
  };
  bookmarks: { marksOn(model: monaco.editor.ITextModel): number[] };
  panes: { list(): { name: string; title: string; open: boolean; permanent: boolean }[] };
  openSettings(): void;
  openSponsors(): void;
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
 * exactly like "no jank" had it been believed (2026-08-07). Provoke from inside the page — a
 * setTimeout, a real interaction — or provoke nothing and read what the session collected.
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
   * Takes `line`/`column`, or `word` to find the first occurrence of an identifier — which is
   * how a person describes where they are looking ("hover over Recalculate") without counting
   * columns. Null when nothing is open.
   */
  const offsetFrom = (args: Record<string, unknown>): number | null => {
    const model = workspace.activeEditor().getModel();
    if (!model) { return null; }

    if (args.word !== undefined) {
      const wanted = String(args.word);
      const at = model.getValue().indexOf(wanted);
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

    activate: (args) => {
      const module = String(args.module ?? "");
      const project = args.project === undefined || args.project === null
        ? null
        : String(args.project);
      if (!module) {
        return { did: false, detail: "no module given" };
      }
      workspace.reveal({ module, project });
      return { did: true, detail: `revealed ${module}` };
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

    settings: () => {
      parts.openSettings();
      return { did: true, detail: "settings dialog opened" };
    },

    sponsors: () => {
      parts.openSponsors();
      return { did: true, detail: "sponsor dialog opened" };
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
      return { did: true, detail: `searching for ${JSON.stringify(query)}` };
    },

    /*
     * THE LANGUAGE FEATURES, asked at a position.
     *
     * Each goes through the bridge request its monaco provider goes through, so the whole chain
     * is exercised — page to shim to engine and back — without depending on a widget being on
     * screen. Reading the rendered hover widget instead is the trap this repo already paid for:
     * monaco reuses an open widget, so a probe that triggers and scrapes gets the previous
     * answer in every state, including the ones where the feature is switched off (finding 1).
     *
     * What is on SCREEN is a different question, and `ui`'s `at` field answers that one.
     */

    hover: async (args) => {
      const offset = offsetFrom(args);
      if (offset === null) { return { did: false, detail: "no module is active, or no such word" }; }

      const found = await bridge.requestHover(offset);
      return { did: found !== null, detail: found ? "hover answered" : "nothing to say here", data: found };
    },

    completions: async (args) => {
      const offset = offsetFrom(args);
      if (offset === null) { return { did: false, detail: "no module is active, or no such word" }; }

      const items = await bridge.requestCompletions(offset);
      return {
        did: items.length > 0,
        detail: `${items.length} completion(s)`,
        data: items.map((one) => ({ label: one.label, kind: one.kind, detail: one.detail })),
      };
    },

    signature: async (args) => {
      const offset = offsetFrom(args);
      if (offset === null) { return { did: false, detail: "no module is active, or no such word" }; }

      const found = await bridge.requestSignatureHelp(offset);
      return { did: found !== null, detail: found ? "signature help answered" : "no signature here", data: found };
    },

    quickFixes: async (args) => {
      const offset = offsetFrom(args);
      if (offset === null) { return { did: false, detail: "no module is active, or no such word" }; }

      // The whole word under the position, so a fix attached to an identifier is found from
      // anywhere inside it rather than only from its first character.
      const model = workspace.activeEditor().getModel();
      const position = model?.getPositionAt(offset);
      const word = model && position ? model.getWordAtPosition(position) : null;
      const start = model && position && word
        ? model.getOffsetAt({ lineNumber: position.lineNumber, column: word.startColumn })
        : offset;
      const end = model && position && word
        ? model.getOffsetAt({ lineNumber: position.lineNumber, column: word.endColumn })
        : offset;

      const actions = await bridge.requestCodeActions(start, end);
      return {
        did: actions.length > 0,
        detail: `${actions.length} quick fix(es)`,
        data: actions.map((one) => ({ title: one.title, isPreferred: one.isPreferred })),
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
