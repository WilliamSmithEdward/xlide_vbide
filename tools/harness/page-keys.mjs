/*
 * Keys pressed as keys, for the suites that need the page to see a real key event.
 *
 * Through the page's DevTools protocol (Input.dispatchKeyEvent), which delivers a trusted key
 * event to the page and runs it through everything a key from the keyboard runs through: the
 * editor's key listeners, Monaco's keybinding resolution, and the browser's own rule that a key
 * whose keydown was cancelled types nothing. Nothing goes to the desktop: the event reaches only
 * the page, so whatever is to receive it must have the focus first.
 *
 * What it does NOT run through is the host's accelerator hook, which sits in front of the page
 * (F5, F8, F9, Ctrl+S and the rest of VbeCommands.ForKey): those keys are the host's and are
 * driven through `api.command`.
 *
 * The keys are named as a US keyboard has them, which is what `code` and the virtual key say.
 */

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { wait } from "./xlide-api.mjs";
import { loopbackUrl } from "./loopback-url.mjs";

/** The DevTools port of the session the api client is aimed at. */
export function devtoolsPort(api) {
  const home = join(process.env.LOCALAPPDATA ?? "", "xlide_vbide");
  const pid = Number(process.env.XLIDE_PID) || api.pid;
  return JSON.parse(readFileSync(join(home, `xlide-api-${pid}.json`), "utf8")).devtoolsPort;
}

/** Attaches to the editor page. Answers `{ socket, send(method, params) }`; close the socket. */
export async function connectToEditorPage(port) {
  const version = await (await fetch(loopbackUrl(port, "/json/version"), { redirect: "error" })).json();
  const socket = new WebSocket(version.webSocketDebuggerUrl);
  let nextId = 0;
  const waiting = new Map();
  socket.onmessage = (event) => {
    const message = JSON.parse(event.data);
    const pending = message.id && waiting.get(message.id);
    if (pending) {
      waiting.delete(message.id);
      if (message.error) { pending.reject(new Error(message.error.message)); } else { pending.settle(message.result); }
    }
  };
  await new Promise((settle, reject) => {
    socket.onopen = settle;
    socket.onerror = () => reject(new Error("could not reach the DevTools door"));
  });
  const send = (method, params = {}, sessionId) => new Promise((settle, reject) => {
    const id = ++nextId;
    waiting.set(id, { settle, reject });
    socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
  });

  const targets = await (await fetch(loopbackUrl(port, "/json"), { redirect: "error" })).json();
  const page = targets.find((one) => one.type === "page" && one.title === "xlide editor");
  if (!page) { throw new Error(`no editor page; saw ${targets.map((one) => one.title).join("|")}`); }
  const { sessionId } = await send("Target.attachToTarget", { targetId: page.id, flatten: true });
  return { socket, send: (method, params) => send(method, params, sessionId) };
}

/**
 * For a suite whose keys are meant as plain typing: turns Vim mode off for the run and answers
 * the function that puts the developer's choice back, for its `finally`. With Vim mode on a key
 * is a command, so a suite that types `'` to make an edit would be setting a mark (#173).
 */
export async function plainTyping(api) {
  const was = (await api.settings()).vimMode === true;
  if (was) {
    await api.settings({ vimMode: false });
    for (let waited = 0; waited < 5000 && (await api.ui()).vim?.enabled; waited += 150) {
      await wait(150);
    }
  }
  return async () => { if (was) { await api.settings({ vimMode: true }).catch(() => {}); } };
}

const NAMED = {
  Enter: { key: "Enter", code: "Enter", windowsVirtualKeyCode: 13, text: "\r" },
  Escape: { key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 },
  Backspace: { key: "Backspace", code: "Backspace", windowsVirtualKeyCode: 8 },
  ArrowUp: { key: "ArrowUp", code: "ArrowUp", windowsVirtualKeyCode: 38 },
  ArrowDown: { key: "ArrowDown", code: "ArrowDown", windowsVirtualKeyCode: 40 },
  ArrowLeft: { key: "ArrowLeft", code: "ArrowLeft", windowsVirtualKeyCode: 37 },
  ArrowRight: { key: "ArrowRight", code: "ArrowRight", windowsVirtualKeyCode: 39 },
};

/** Each symbol's key on a US keyboard: [code, virtual key, the character without Shift]. */
const SYMBOL_KEYS = [
  ["Semicolon", 186, ";", ":"], ["Slash", 191, "/", "?"], ["Period", 190, ".", ">"],
  ["Comma", 188, ",", "<"], ["BracketLeft", 219, "[", "{"], ["BracketRight", 221, "]", "}"],
  ["Minus", 189, "-", "_"], ["Equal", 187, "=", "+"], ["Quote", 222, "'", '"'],
  ["Backquote", 192, "`", "~"], ["Backslash", 220, "\\", "|"],
];
const SHIFTED_DIGITS = ")!@#$%^&*(";

/** The key that types a character on a US keyboard, as the protocol wants it described. */
function keyTyping(character) {
  if (character === " ") {
    return { key: " ", code: "Space", windowsVirtualKeyCode: 32, text: " " };
  }

  if (/^[a-z]$/i.test(character)) {
    const upper = character.toUpperCase();
    return {
      key: character, code: `Key${upper}`, windowsVirtualKeyCode: upper.charCodeAt(0),
      text: character, shift: character === upper,
    };
  }

  if (/^[0-9]$/.test(character)) {
    return { key: character, code: `Digit${character}`, windowsVirtualKeyCode: character.charCodeAt(0), text: character };
  }

  const digit = SHIFTED_DIGITS.indexOf(character);
  if (digit >= 0) {
    return { key: character, code: `Digit${digit}`, windowsVirtualKeyCode: 48 + digit, text: character, shift: true };
  }

  for (const [code, virtualKey, plain, shifted] of SYMBOL_KEYS) {
    if (character === plain || character === shifted) {
      return { key: character, code, windowsVirtualKeyCode: virtualKey, text: character, shift: character === shifted };
    }
  }

  throw new Error(`no key is known to type ${JSON.stringify(character)}`);
}

/**
 * A keyboard for the page: `press("Enter")`, `press("d")`, `press("[", { ctrl: true })`, and
 * `type("ciw")` for one press per character. A key held with Ctrl or Alt types no character,
 * as on a keyboard. `settleMs` is waited after each press, for the editor to act on it.
 */
export function pageKeyboard(page, { settleMs = 150 } = {}) {
  const press = async (name, { ctrl = false, alt = false, shift = false } = {}) => {
    const { shift: needsShift, text, ...key } = NAMED[name] ?? keyTyping(name);
    const modifiers = (alt ? 1 : 0) | (ctrl ? 2 : 0) | (shift || needsShift ? 8 : 0);
    const typed = ctrl || alt ? undefined : text;
    await page.send("Input.dispatchKeyEvent", { type: "keyDown", ...key, modifiers, ...(typed ? { text: typed } : {}) });
    await page.send("Input.dispatchKeyEvent", { type: "keyUp", ...key, modifiers });
    await wait(settleMs);
  };

  const type = async (text) => {
    for (const character of text) {
      await press(character);
    }
  };

  return { press, type };
}
