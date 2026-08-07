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
}

export interface DevSurfaceParts {
  workspace: Workspace;
  explorer: Explorer;
  bridge: EditorBridge;
  panes: { list(): { name: string; title: string; open: boolean; permanent: boolean }[] };
  openSettings(): void;
  openSponsors(): void;
}

/**
 * Every page-side dialog by the id of its root element, which is how each one already tests for
 * its own presence. Names, not selectors, so a probe asserts on "settings" rather than on markup.
 */
const DIALOG_IDS: Record<string, string> = {
  "settings-backdrop": "settings",
  "help-backdrop": "help",
  "sponsor-backdrop": "sponsors",
  "references-backdrop": "references",
  "objbrowser-card": "object browser",
  "panes-menu": "panes",
};

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

function dialogsUp(): DialogSnapshot[] {
  return Object.entries(DIALOG_IDS)
    .filter(([id]) => document.getElementById(id) !== null)
    .map(([id, title]) => ({ id, title }));
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
}

export function installDevSurface(parts: DevSurfaceParts): void {
  const { workspace, explorer, bridge } = parts;

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

  const state = (): UiSnapshot => ({
    workspace: workspace.snapshot(),
    explorer: explorer.treeState(),
    panes: parts.panes.list().map(({ name, title, open }) => ({ name, title, open })),
    dialogs: dialogsUp(),
    waiting: bridge.waitingOn(),
    focus: editorFocus(),
    settings: { ...currentSettings() } as unknown as Record<string, unknown>,
    emptyViewShown: workspace.emptyViewShown(),
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

  const actions: Record<string, (args: Record<string, unknown>) => ActResult> = {
    closeActive: () => {
      const before = workspace.activeDocument();
      if (!before) {
        return { did: false, detail: "nothing is active" };
      }
      workspace.closeActive();
      return { did: true, detail: `closed ${before.module}` };
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
  };

  const act = (name: string, args?: Record<string, unknown>): ActResult => {
    const action = actions[name];
    if (!action) {
      return { did: false, detail: `unknown action ${name}; try ${Object.keys(actions).sort().join(", ")}` };
    }

    try {
      return action(args ?? {});
    } catch (error) {
      // Reported rather than thrown: the door carries a script error as a bare message with no
      // stack, and "which action failed" is the half that goes missing.
      return { did: false, detail: `${name} threw: ${error instanceof Error ? error.message : String(error)}` };
    }
  };

  (globalThis as { xlideUi?: unknown }).xlideUi = { state, act, actions: () => Object.keys(actions).sort() };
}
