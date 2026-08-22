/*
 * The agent card: the api's switch, and the address it hands out.
 *
 * WHAT THIS IS FOR. A developer wants an agent - Claude, Codex, whatever they have - to work on
 * the VBA in front of them. The agent needs to know that this api exists, where it is, and how to
 * behave once it is in. All of that is one paste, and this is where the paste comes from.
 *
 * THE TEXT IS TAILORED, WHICH IS THE WHOLE POINT (the owner, 2026-08-22: "the instructions should
 * be tailored to that specific instance, including port etc."). Generic instructions would send
 * the agent hunting for a discovery file and guessing which of several running Excels it found.
 * These carry this session's port, this session's token, this host application and this process,
 * so the first request the agent makes is the right one against the right instance.
 *
 * AND THE SWITCH IS HERE BECAUSE THE CONSEQUENCE IS HERE. Turning the api on opens a door into
 * the open projects; the sentence explaining that belongs beside the control that does it, not in
 * a settings list three screens away.
 */

import { openModal } from "./modal.js";

/** How the card reaches the host. One function, because there is only one kind of request. */
export type ApiRequest = (args: Record<string, string>) => Promise<Record<string, unknown>>;

/** The door as the host reports it. */
export interface ApiState {
  on: boolean;
  leansOpen: boolean;
  remembered: boolean;
  host: string;
  pid: number;
  port: number;
  token: string;
  baseUrl: string;
  agentUrl: string;
  progId: string;
  discovery: string;
  project: string;
  error?: string;
}

/** What the card can be driven and read through, for the dev surface. */
export interface AgentDialogProbe {
  state(): {
    open: boolean;
    api: boolean;
    busy: boolean;
    /** The instruction text as it stands, so a check can read what a developer would paste. */
    text: string;
    copied: boolean;
  };
  /** Presses a named control: toggle, copy, close. False when unknown. */
  press(control: string): boolean;
}

let liveDialog: AgentDialogProbe | null = null;

/** The agent card's probe, or null when it is not up. */
export const agentDialogProbe = (): AgentDialogProbe | null => liveDialog;

const BLANK: ApiState = {
  on: false,
  leansOpen: false,
  remembered: false,
  host: "",
  pid: 0,
  port: 0,
  token: "",
  baseUrl: "",
  agentUrl: "",
  progId: "",
  discovery: "",
  project: "",
};

/**
 * The paste itself.
 *
 * Written to be READ BY A MODEL, so it leads with the address and the first call rather than with
 * prose about what xlide is: an agent that has the URL can ask the api to describe itself, and
 * `agent` answers better than any summary kept in step by hand here. What cannot be discovered by
 * asking - which instance this is, and how to behave while writing someone's code - is what the
 * rest of it spends its words on.
 */
export function instructionsFor(state: ApiState): string {
  if (!state.on) {
    return "";
  }

  const where = state.project ? state.project : "no workbook open yet";
  return [
    `xlide is running inside ${state.host || "Office"} and has opened a local api for you.`,
    "",
    `  Base URL    ${state.baseUrl}`,
    `  Start here  GET ${state.agentUrl}`,
    "",
    "That first GET describes itself: what this api is, which application it is running in, and",
    "where to go next. Follow it rather than guessing - `agent/routes` lists every route with its",
    "arguments and an example, `agent/route?name=<route>` explains one, and `agent/examples` gives",
    "runnable recipes. Every reply carries `next` links, so there is no step where you have to",
    "already know the answer.",
    "",
    `This session: ${state.host || "Office"}, process ${state.pid}, showing ${where}.`,
    `Several editors can be open at once - this address is THIS one, and the \`agent\` reply's`,
    "`pid` is how you confirm you are talking to it.",
    "",
    "How to behave in here:",
    "",
    `- Put \`by=<your name>\` on every write. It is what the developer sees in the Changes pane,`,
    "  and a write without it is recorded as `unattributed`.",
    "- Read a module before you write it. A write replaces the module's whole text.",
    "- There is no revert route, deliberately. To undo, write the previous text back - your undo",
    "  lands in the change log like any other edit, which is what makes it reviewable.",
    "- Prefer the named routes to `eval`. `eval` runs script in the editor's own page; it exists",
    "  for diagnosis, not for getting work done.",
    "",
    "Keep this address to yourself. It is loopback-only, but it carries a token that is the whole",
    "of its security: anything holding it can read, write and run code in the open projects. Do",
    "not paste it into code, commits, issues, or anywhere it outlives this conversation.",
    "",
    `From VBA or another local process on this machine: GetObject(, "${state.progId}").`,
  ].join("\n");
}

/**
 * Opens the card. One at a time; a second call while one stands does nothing.
 */
export function openAgentDialog(ask: ApiRequest, closed?: () => void): void {
  if (liveDialog) {
    return;
  }

  let state: ApiState = BLANK;
  let busy = false;
  let copied = false;

  const { card, dismiss } = openModal({
    backdropId: "agent-backdrop",
    cardId: "agent-card",
    label: "Work with an agent",
    closed: () => {
      liveDialog = null;
      closed?.();
    },
  });

  const head = document.createElement("div");
  head.id = "agent-head";

  const title = document.createElement("h2");
  title.id = "agent-title";
  title.textContent = "Work with an agent";

  const close = document.createElement("button");
  close.type = "button";
  close.id = "agent-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';
  close.addEventListener("click", () => dismiss());

  head.append(title, close);

  // ---- the switch ------------------------------------------------------------------------

  const gate = document.createElement("div");
  gate.id = "agent-gate";

  const toggle = document.createElement("input");
  toggle.type = "checkbox";
  toggle.id = "agent-toggle";
  toggle.setAttribute("aria-describedby", "agent-consequence");

  const toggleLabel = document.createElement("label");
  toggleLabel.htmlFor = "agent-toggle";
  toggleLabel.id = "agent-toggle-label";
  toggleLabel.textContent = "Let agents reach this editor";

  const lit = document.createElement("span");
  lit.id = "agent-lit";

  const consequence = document.createElement("p");
  consequence.id = "agent-consequence";

  gate.append(toggle, toggleLabel, lit);

  // ---- the paste -------------------------------------------------------------------------

  const paste = document.createElement("div");
  paste.id = "agent-paste";

  const pasteLabel = document.createElement("label");
  pasteLabel.htmlFor = "agent-text";
  pasteLabel.id = "agent-paste-label";
  pasteLabel.textContent = "Paste this to your agent";

  const copy = document.createElement("button");
  copy.type = "button";
  copy.id = "agent-copy";
  copy.className = "agent-action";
  copy.innerHTML = '<span class="codicon codicon-copy" aria-hidden="true"></span><span id="agent-copy-word">Copy</span>';

  const pasteHead = document.createElement("div");
  pasteHead.id = "agent-paste-head";
  pasteHead.append(pasteLabel, copy);

  const text = document.createElement("textarea");
  text.id = "agent-text";
  text.rows = 16;
  text.spellcheck = false;
  text.readOnly = true;

  // Said out loud when it changes, because the button's own label going back to "Copy" after two
  // seconds is a change nobody watching the text area would see.
  const said = document.createElement("p");
  said.id = "agent-said";
  said.setAttribute("role", "status");

  paste.append(pasteHead, text, said);

  card.append(head, gate, consequence, paste);

  // ---- drawing ---------------------------------------------------------------------------

  let copiedTimer = 0;

  const draw = (): void => {
    toggle.checked = state.on;
    toggle.disabled = busy;
    lit.textContent = state.on ? "on" : "off";
    lit.className = state.on ? "agent-lit-on" : "agent-lit-off";

    consequence.textContent = state.on
      ? "A program on this machine that has your token can read, write and run code in the open"
        + " projects, exactly as this editor can. It stays on until you turn it off, including"
        + " after a restart."
      : "Off, nothing outside this editor can reach your projects. Turning it on opens a door on"
        + " this machine only - no network - and writes a token to a file under your profile."
        + " Anything running as you can read that file.";

    const script = instructionsFor(state);
    text.value = script || (state.error
      ? `The editor could not answer: ${state.error}`
      : "Turn the switch on and this fills with the address to hand your agent.");
    text.disabled = !state.on;
    copy.disabled = !state.on || busy;
    paste.classList.toggle("agent-paste-shut", !state.on);
  };

  const refresh = async (action: "state" | "on" | "off"): Promise<void> => {
    busy = true;
    draw();
    try {
      const answer = await ask({ action });
      state = { ...BLANK, ...(answer as unknown as ApiState) };
    } catch {
      state = { ...state, error: "the editor did not answer" };
    } finally {
      busy = false;
      draw();
    }
  };

  toggle.addEventListener("change", () => {
    void refresh(toggle.checked ? "on" : "off");
  });

  const copyOut = async (): Promise<void> => {
    const script = instructionsFor(state);
    if (!script) {
      return;
    }

    let done = false;
    try {
      await navigator.clipboard.writeText(script);
      done = true;
    } catch {
      // The page is served over loopback, which IS a secure context, so the clipboard api is
      // normally there - but a host that has withheld the permission would otherwise leave the
      // button doing nothing at all, and a button that silently fails is worse than no button.
      text.disabled = false;
      text.select();
      done = document.execCommand("copy");
      text.disabled = !state.on;
    }

    copied = done;
    said.textContent = done
      ? "Copied. Paste it into your agent's chat."
      : "Could not reach the clipboard - select the text and copy it yourself.";
    (document.getElementById("agent-copy-word") as HTMLElement).textContent = done ? "Copied" : "Copy";

    window.clearTimeout(copiedTimer);
    copiedTimer = window.setTimeout(() => {
      copied = false;
      said.textContent = "";
      const word = document.getElementById("agent-copy-word");
      if (word) {
        word.textContent = "Copy";
      }
    }, 4000);
  };

  copy.addEventListener("click", () => void copyOut());

  liveDialog = {
    state: () => ({ open: true, api: state.on, busy, text: text.value, copied }),
    press: (control) => {
      if (control === "toggle") {
        toggle.checked = !toggle.checked;
        toggle.dispatchEvent(new Event("change"));
        return true;
      }

      if (control === "copy") {
        copy.click();
        return true;
      }

      if (control === "close") {
        dismiss();
        return true;
      }

      return false;
    },
  };

  draw();
  toggle.focus();
  void refresh("state");
}
