/*
 * Vim mode (#173): modal editing over the code editors, off until the developer asks for it.
 *
 * The keymap is monaco-vim's, which is CodeMirror's Vim emulation behind an adapter that speaks
 * to Monaco. This module is everything around it that is this product's to decide: which editors
 * carry it, which keys it may have, what `:w` and `u` mean here, and where the mode is shown.
 *
 * WHICH KEYS ARE VIM'S is vimkeys.ts: plain keys, shifted keys and Escape, and of the Ctrl chords
 * only the four the editor has no use for.
 *
 * ONE VIM, MANY EDITORS. The keymap, the registers and the ex command table are the library's
 * module state, shared by every editor on the page, which is what a yank in one group and a put
 * in the other needs. The mode is per editor, and the status bar shows the mode of the editor
 * that last had the caret.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import { VimMode } from "monaco-vim";
import { routeVimKey } from "./vimkeys.js";

type Editor = monaco.editor.IStandaloneCodeEditor;

/** What this module asks of the rest of the page. */
export interface VimModeHooks {
  /** `:w`. The host's Save: the project is what gets written, not the module. */
  save(): void;
  /**
   * `:q` closes the active tab through the gate its own X uses, which asks about unsaved
   * changes. `:wq` and `:x` pass `save`: the answer that question's Save button gives.
   */
  closeTab(save: boolean): void;
  /** `u`. The surface's Undo, which knows a rename is undone across modules. */
  undo(editor: Editor): void;
  /** True while text nobody typed is going into a model: a host edit, a closer, a recase. */
  automated(editor: Editor): boolean;
  /** The editor of the active group, shown until one has been focused. */
  activeEditor(): Editor;
}

/** What the xlide api's `ui` route reports: the switch, and the editor on screen's mode. */
export interface VimModeState {
  enabled: boolean;
  /** "normal", "insert", "visual", "visual line", "visual block" or "replace"; "" when off. */
  mode: string;
  /** The keys of a command still being typed, such as `2d` before the motion. */
  keys: string;
  /** What the status bar's mode label reads. */
  label: string;
  /** The text in the `:` or `/` line while one is open, or null. */
  prompt: string | null;
  /** The last message Vim put in the status bar, while it is still showing. */
  notice: string;
}

/** The options Vim passes with a prompt, as monaco-vim's own status bar reads them. */
interface PromptOptions {
  selectValueOnOpen?: boolean;
  value?: string;
  onKeyUp?: (event: KeyboardEvent, value: string, close: () => void) => void;
  onKeyDown?: (event: KeyboardEvent, value: string, close: () => void) => boolean | void;
  onKeyInput?: (event: Event, value: string, close: () => void) => void;
  onBlur?: (event: FocusEvent, close: () => void) => void;
  closeOnBlur?: boolean;
  closeOnEnter?: boolean;
}

/** The part of an adapter's status bar the keymap calls. */
interface VimStatusSurface {
  setSec(
    content: Node | string | null | undefined,
    callback?: ((value: string) => void) | null,
    options?: PromptOptions,
  ): () => void;
  showNotification(content: Node | string): void;
}

/** One editor's adapter, as far as this module uses it. The library's own types are `any`. */
interface VimAdapter {
  editor: Editor;
  state: {
    vim?: {
      insertMode?: boolean;
      inputState?: {
        keyBuffer?: string | string[];
        prefixRepeat?: string[];
        motionRepeat?: string[];
        operator?: string | null;
        registerName?: string | null;
      };
    } | null;
  };
  attach(): void;
  dispose(): void;
  dispatch(signal: string, ...args: unknown[]): void;
  on(event: "vim-mode-change", handler: (event: { mode: string; subMode?: string }) => void): void;
  on(event: "vim-keypress", handler: (key: string) => void): void;
  on(event: "vim-command-done", handler: () => void): void;
  setStatusBar(bar: VimStatusSurface): void;
}

interface VimKeyMap {
  call(key: string, cm: VimAdapter | undefined): unknown;
}

interface VimLibrary {
  new (editor: Editor): VimAdapter;
  keyMap: Record<string, VimKeyMap | undefined>;
  commands: Record<string, ((cm: VimAdapter) => void) | undefined>;
  Vim: { defineEx(name: string, prefix: string, run: (cm: VimAdapter) => void): void };
}

const Vim = VimMode as unknown as VimLibrary;

interface Session {
  adapter: VimAdapter;
  mode: string;
  subMode: string;
  keys: string;
}

/** How long a message from Vim stays in the status bar, the library's own figure. */
const NOTICE_MS = 5000;

/**
 * The key event being handled, kept for one question the keymap's own key name cannot answer:
 * whether AltGr produced a character. Set by a listener each editor registers before its
 * adapter exists, so it runs first.
 */
let keyBeingHandled: KeyboardEvent | null = null;

/** The key policy (vimkeys.ts), asked with what only the live editor knows. */
function routeKey(key: string, cm: VimAdapter | undefined): string | null {
  const event = keyBeingHandled;
  return routeVimKey(key, {
    insertMode: cm?.state.vim?.insertMode === true,
    altGraphCharacter: event?.getModifierState("AltGraph") ? event.key : null,
  });
}

let libraryPrepared = false;

/**
 * Once per page: the key policy over the library's three keymaps, and the commands that mean
 * something of this product's. The library holds all of it as module state, so doing it per
 * editor would wrap the wrappers.
 */
function prepareLibrary(hooks: VimModeHooks): void {
  if (libraryPrepared) {
    return;
  }
  libraryPrepared = true;

  for (const name of ["vim", "vim-insert", "vim-replace"]) {
    const map = Vim.keyMap[name];
    if (!map) {
      continue;
    }

    const lookUp = map.call;
    map.call = function (this: unknown, key: string, cm: VimAdapter | undefined): unknown {
      const routed = routeKey(key, cm);
      return routed === null ? undefined : lookUp.call(this, routed, cm);
    };
  }

  // `:w` runs the keymap's save command when one is defined, and nothing when none is.
  Vim.commands.save = () => hooks.save();
  Vim.commands.undo = (cm) => hooks.undo(cm.editor);

  // `:wq` is one request, not a save and then a close: the save is the host's and takes a
  // moment, and a close sent after it asked "save your changes?" of a developer who just had.
  Vim.Vim.defineEx("quit", "q", () => hooks.closeTab(false));
  Vim.Vim.defineEx("wq", "wq", () => hooks.closeTab(true));
  Vim.Vim.defineEx("xit", "x", () => hooks.closeTab(true));
}

/**
 * Whether the keymap is part-way through a command: a count typed, an operator waiting for its
 * motion, a register named, or the first key of a two-key command.
 */
function commandPending(adapter: VimAdapter): boolean {
  const input = adapter.state.vim?.inputState;
  if (!input) {
    return false;
  }

  return (input.keyBuffer?.length ?? 0) > 0
    || (input.prefixRepeat?.length ?? 0) > 0
    || (input.motionRepeat?.length ?? 0) > 0
    || Boolean(input.operator)
    || Boolean(input.registerName);
}

function labelOf(session: Session): string {
  if (session.mode === "visual" && session.subMode === "linewise") {
    return "-- VISUAL LINE --";
  }

  if (session.mode === "visual" && session.subMode === "blockwise") {
    return "-- VISUAL BLOCK --";
  }

  return `-- ${session.mode.toUpperCase()} --`;
}

function modeOf(session: Session): string {
  if (session.mode === "visual" && session.subMode === "linewise") {
    return "visual line";
  }

  if (session.mode === "visual" && session.subMode === "blockwise") {
    return "visual block";
  }

  return session.mode;
}

export class VimModeController {
  /** Every code editor, with its session while Vim mode is on. */
  private readonly editors = new Map<Editor, Session | null>();
  private enabled = false;

  /** The editor that last had the caret: the one whose mode the status bar shows. */
  private shown: Editor | null = null;

  private readonly modeLabel: HTMLElement;
  private readonly promptSlot: HTMLElement;
  private readonly notice: HTMLElement;
  private readonly keysLabel: HTMLElement;

  private prompt: { owner: Editor; input: HTMLInputElement | null; listeners: AbortController } | null = null;
  private noticeTimer: ReturnType<typeof setTimeout> | undefined;

  /**
   * `root` is the status bar's Vim region, empty and hidden in the markup. It is filled here
   * and shown only while the mode is on.
   */
  constructor(private readonly root: HTMLElement, private readonly hooks: VimModeHooks) {
    const part = (id: string): HTMLElement => {
      const element = document.createElement("span");
      element.id = id;
      root.appendChild(element);
      return element;
    };

    this.modeLabel = part("status-vim-mode");
    this.promptSlot = part("status-vim-prompt");
    this.notice = part("status-vim-notice");
    this.keysLabel = part("status-vim-keys");

    // Said aloud when it changes: the mode is the one thing here a screen reader cannot infer.
    this.modeLabel.setAttribute("role", "status");
    this.notice.setAttribute("role", "status");
    this.keysLabel.setAttribute("aria-hidden", "true");
    root.hidden = true;
  }

  /** Takes a code editor on, for as long as it lives. It gets Vim whenever the mode is on. */
  attach(editor: Editor): void {
    if (this.editors.has(editor)) {
      return;
    }
    this.editors.set(editor, null);

    // Before the adapter's own listener, which is registered when the mode goes on: see
    // keyBeingHandled.
    editor.onKeyDown((event) => { keyBeingHandled = event.browserEvent; });
    editor.onDidFocusEditorText(() => {
      this.shown = editor;
      this.paint();
    });
    editor.onDidDispose(() => {
      this.stop(editor);
      this.editors.delete(editor);
      if (this.shown === editor) {
        this.shown = null;
      }
      this.paint();
    });

    if (this.enabled) {
      this.start(editor);
      this.paint();
    }
  }

  /** The developer's switch. Off puts every editor back exactly as it was, caret included. */
  setEnabled(enabled: boolean): void {
    if (enabled === this.enabled) {
      return;
    }
    this.enabled = enabled;

    for (const editor of [...this.editors.keys()]) {
      if (enabled) {
        this.start(editor);
      } else {
        this.stop(editor);
      }
    }
    this.paint();
  }

  /** What the status bar says, as data, for the editor it is speaking about. */
  state(): VimModeState {
    const session = this.shownSession();
    return {
      enabled: this.enabled,
      mode: session ? modeOf(session) : "",
      keys: session?.keys ?? "",
      label: this.modeLabel.textContent ?? "",
      prompt: this.prompt ? this.prompt.input?.value ?? this.promptSlot.textContent ?? "" : null,
      notice: this.notice.textContent ?? "",
    };
  }

  private shownSession(): Session | null {
    const editor = this.shown ?? this.hooks.activeEditor();
    return this.editors.get(editor) ?? null;
  }

  private start(editor: Editor): void {
    if (this.editors.get(editor)) {
      return;
    }
    prepareLibrary(this.hooks);

    const adapter = new Vim(editor);
    const session: Session = { adapter, mode: "normal", subMode: "", keys: "" };
    this.editors.set(editor, session);

    /*
     * TEXT NOBODY TYPED IS NOT TYPING. The keymap records what goes into the model during insert
     * mode so that `.` can type it again, and it cannot tell the developer's keys from the edits
     * this page makes around them: the `End If` Smart Enter writes, the respelling of `dim` to
     * `Dim` a moment after the line is left alone. Recorded, a repeat would type both the
     * original and its correction. So those changes are kept from the keymap, and a repeat types
     * what the developer typed, which the automation then treats as it treated the first.
     */
    const dispatch = adapter.dispatch.bind(adapter);
    adapter.dispatch = (signal, ...args) => {
      if (signal === "change" && this.hooks.automated(editor)) {
        return;
      }
      dispatch(signal, ...args);
    };

    adapter.on("vim-mode-change", (event) => {
      const leftInsert = session.mode === "insert" || session.mode === "replace";
      session.mode = event.mode;
      session.subMode = event.subMode ?? "";
      session.keys = "";

      // Escape was Vim's, so the popups that would have closed on it never heard it.
      if (leftInsert && event.mode === "normal") {
        editor.trigger("vim", "hideSuggestWidget", null);
        editor.trigger("vim", "closeParameterHints", null);
      }
      this.paint();
    });
    adapter.on("vim-keypress", (key) => {
      session.keys = key === ":" ? "" : session.keys + key;
      this.paint();

      // The keymap announces a key before it acts on it, and says "done" for a finished command
      // only: not for the Escape that leaves insert mode, nor the q that ends a recording. So
      // once the key has been acted on, the bar keeps it only if a command is still open.
      queueMicrotask(() => {
        if (session.keys !== "" && !commandPending(adapter)) {
          session.keys = "";
          this.paint();
        }
      });
    });
    adapter.on("vim-command-done", () => {
      session.keys = "";
      this.paint();
    });

    adapter.setStatusBar({
      setSec: (content, callback, options) => this.openPrompt(editor, content, callback ?? null, options ?? {}),
      showNotification: (content) => this.showNotice(content),
    });
    adapter.attach();
  }

  private stop(editor: Editor): void {
    const session = this.editors.get(editor);
    if (!session) {
      return;
    }

    if (this.prompt?.owner === editor) {
      this.closePrompt(false);
    }
    // Detaches the keymap, which puts the caret's shape and blink back as it found them.
    session.adapter.dispose();
    this.editors.set(editor, null);
  }

  private paint(): void {
    const session = this.enabled ? this.shownSession() : null;
    this.root.hidden = !this.enabled;
    const label = session ? labelOf(session) : "";
    // Only on a change: an unchanged status region rewritten is announced again.
    if (this.modeLabel.textContent !== label) {
      this.modeLabel.textContent = label;
    }
    this.keysLabel.textContent = session?.keys ?? "";

    if (!this.enabled) {
      this.notice.textContent = "";
    }
  }

  /**
   * Vim's command line: the `:` and `/` prompts, and the standing "(recording)" note of a macro.
   * The keymap hands over the nodes, an input among them when there is something to type, and
   * gets back the function that closes it. This is monaco-vim's own status bar contract, kept
   * key for key; what differs is that a prompt closed by a click elsewhere does not pull the
   * focus back from where the developer just put it.
   */
  private openPrompt(
    owner: Editor,
    content: Node | string | null | undefined,
    callback: ((value: string) => void) | null,
    options: PromptOptions,
  ): () => void {
    const close = (): void => this.closePrompt(true);
    this.notice.textContent = "";
    if (content === undefined) {
      return close;
    }

    this.closePrompt(false);
    if (content === null || content === "") {
      return close;
    }

    this.promptSlot.append(content);
    const input = this.promptSlot.querySelector("input");
    const listeners = new AbortController();
    this.prompt = { owner, input, listeners };
    if (!input) {
      return close;
    }

    input.setAttribute("aria-label", "Vim command line");
    if (options.value) {
      input.value = options.value;
    }
    input.focus();
    if (options.selectValueOnOpen) {
      input.select();
    }

    const signal = listeners.signal;
    input.addEventListener("keyup", (event) => options.onKeyUp?.(event, input.value, close), { signal });
    input.addEventListener("input", (event) => options.onKeyInput?.(event, input.value, close), { signal });
    input.addEventListener("blur", (event) => {
      const leave = (): void => this.closePrompt(false);
      options.onBlur?.(event, leave);
      if (options.closeOnBlur) {
        leave();
      }
    }, { signal });
    input.addEventListener("keydown", (event) => {
      if (options.onKeyDown?.(event, input.value, close)) {
        return;
      }

      const value = input.value;
      const enter = event.key === "Enter";
      if (event.key === "Escape" || (enter && options.closeOnEnter !== false)) {
        event.stopPropagation();
        close();
      }

      if (enter && callback) {
        event.stopPropagation();
        event.preventDefault();
        callback(value);
      }
    }, { signal });

    return close;
  }

  private closePrompt(returnFocus: boolean): void {
    const open = this.prompt;
    if (!open) {
      return;
    }

    this.prompt = null;
    open.listeners.abort();
    this.promptSlot.replaceChildren();
    if (returnFocus) {
      open.owner.focus();
    }
  }

  private showNotice(content: Node | string): void {
    if (this.noticeTimer !== undefined) {
      clearTimeout(this.noticeTimer);
    }

    // The words only. The keymap dresses its messages in inline colours chosen for another page.
    this.notice.textContent = typeof content === "string" ? content : content.textContent ?? "";
    this.noticeTimer = setTimeout(() => {
      this.noticeTimer = undefined;
      this.notice.textContent = "";
    }, NOTICE_MS);
  }
}
