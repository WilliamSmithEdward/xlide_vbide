/*
 * The Theme setting and the light surface (#114).
 *
 * "Use system" follows Windows' app mode; "Light" and "Dark" hold one. Each choice has to reach
 * four places: the page's root (data-theme, which the light surface's rules key on), the code
 * editor's theme, the title bar the shim dresses (read from its log line, the compositor itself
 * cannot be asked), and the Object Browser palette, which is a page of its own told by the host.
 *
 * Then the light surface is read for contrast: every visible text run against the ground it sits
 * on, composited through its ancestors' backgrounds and its own opacity, held to WCAG AA - 4.5:1
 * for text, 3:1 for large text and icons. Disabled controls are exempt, as WCAG exempts them.
 * And dark is read back as dark, because every light rule is scoped to light and a slip in that
 * scoping would show here first. Last, a name typed into the Rename Symbol box is read in both
 * themes, which the scan cannot see (#128).
 *
 * THE CHANGES PANE IS READ WITH A CHANGE IN IT, not only as it happens to stand. Empty, it shows
 * one sentence, and the scan passed it for as long as nothing before this suite had left a module
 * changed; with rows, its summary button rested at the strip's 0.8 and carried the green count at
 * 3.8:1 and the module tally at 4.3:1 on the light ground. Nothing here would have said so, and
 * the suite went red only on the day it was run after two others (2026-11-25). So it writes a
 * module itself, checks that the summary really came up, reads the pane, the opened summary and
 * a picked change with its diff, and puts the module's text back. Then it presses Accept, with
 * the text already back so nothing counts as changed since the mark, and reads the accepted line:
 * a state nothing reached, where the light green at 0.7 read 3.3:1. Accept cannot be taken back,
 * so a session this suite has run in has its rounds marked reviewed.
 *
 * THAT PANE IS READ IN DARK TOO, and only that pane. Dark is the surface the product was drawn
 * on and nobody had measured it: the removed count's red read 3.4:1 and 2.4:1 on a picked row,
 * round numbers 3.3:1, the diff's line numbers 2.9:1. Its rules were brought to AA on 2026-11-25
 * and the reading here is scoped to `#changes`, because the rest of the dark surface has not
 * been audited and a whole-page reading would go red on things this suite has no claim on yet.
 *
 * Windows' mode is emulated for the page through DevTools (prefers-color-scheme), so the suite
 * does not depend on, or change, the machine's own setting. The developer's Theme choice is put
 * back afterwards.
 *
 * Run against the debug fixture:
 *   node tools\harness\theme.mjs
 */

import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { open, reporter, wait } from "./xlide-api.mjs";
import { loopbackUrl } from "./loopback-url.mjs";
import { contrastScan } from "./contrast.mjs";

const api = await open({});
const { check, done } = reporter();

const targetOf = async (match) => {
  const targets = await (await fetch(loopbackUrl(api.devtoolsPort, "/json"), { redirect: "error" })).json();
  return targets.find(match) ?? null;
};
const connect = async (target) => {
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
  let nextId = 0;
  const waiting = new Map();
  socket.onmessage = ({ data }) => {
    const reply = JSON.parse(data);
    const pending = waiting.get(reply.id);
    if (!pending) return;
    waiting.delete(reply.id);
    reply.error ? pending.reject(new Error(reply.error.message)) : pending.resolve(reply.result);
  };
  const cdp = (method, params) => new Promise((resolve, reject) => {
    const id = ++nextId;
    waiting.set(id, { resolve, reject });
    socket.send(JSON.stringify({ id, method, params }));
  });
  const js = (expression) => cdp("Runtime.evaluate", { expression, returnByValue: true }).then((r) => r.result.value);
  return { socket, cdp, js };
};

const editor = await connect(await targetOf((row) => row.title === "xlide editor"));
const pid = api.pid;
const logDir = join(process.env.LOCALAPPDATA ?? "", "xlide_vbide", "logs");
const log = readdirSync(logDir).filter((name) => name.endsWith(`-${pid}.log`)).map((name) => join(logDir, name))
  .sort((a, b) => statSync(b).mtimeMs - statSync(a).mtimeMs)[0];
const titleBarAskedFor = () => {
  const asks = readFileSync(log, "utf8").split(/\r?\n/).filter((line) => /asked for a (dark|light) title bar/.test(line));
  return /asked for a (dark|light)/.exec(asks.at(-1) ?? "")?.[1] ?? null;
};
const pageShows = () => editor.js(`({ theme: document.documentElement.dataset.theme ?? null,
  editor: [...(document.querySelector('.monaco-editor')?.classList ?? [])].find((c) => c === 'vs' || c === 'vs-dark') ?? null,
  ground: getComputedStyle(document.body).backgroundColor })`);
const emulate = (scheme) => editor.cdp("Emulation.setEmulatedMedia", { features: [{ name: "prefers-color-scheme", value: scheme }] });

// The palette reads its own document; it is opened once and asked after each change.
await api.act("toolbar", { command: "objectBrowser" });
await wait(2500);
const paletteShows = async () => {
  const target = await targetOf((row) => /view=objbrowser/.test(row.url));
  if (!target) return null;
  const palette = await connect(target);
  const theme = await palette.js("document.documentElement.dataset.theme ?? null");
  palette.socket.close();
  return theme;
};

const kept = (await api.settings()).theme;
const project = await api.project();

// The module this suite changes so the Changes pane has something to show, and its text before.
// Putting the text back is what takes the summary down again: it counts what differs from the
// accept mark, and nothing does once the module reads as it did.
const changedModule = "Helper";
let textBefore = null;
const summaryShows = () => editor.js(`(() => { const s = document.getElementById('changes-summary');
  return s && !s.hidden && s.getBoundingClientRect().width > 0 ? s.textContent.trim() : null; })()`);
const writeAChange = async () => {
  textBefore = (await api.readModule(changedModule, project.projectId)).text;
  const lines = textBefore.split(/\r?\n/);
  let last = lines.length - 1;
  while (last > 0 && !lines[last].trim()) last--;
  // One line changed and one added, so the summary carries a green count AND a red one.
  const changed = [...lines.slice(0, last), `${lines[last]} ' theme suite`, "' theme suite: an added line", ...lines.slice(last + 1)].join("\r\n");
  await api.writeModule(changedModule, changed, project.projectId, { by: "theme-suite" });
  let shows = await summaryShows();
  for (let waited = 0; shows === null && waited < 5000; waited += 250) {
    await wait(250);
    shows = await summaryShows();
  }
  return shows;
};
const putTheTextBack = async () => {
  if (textBefore === null) return true;
  const before = textBefore;
  await api.writeModule(changedModule, before, project.projectId, { by: "theme-suite" });
  textBefore = null;
  return (await api.readModule(changedModule, project.projectId)).text === before;
};
// Asks the page until it answers something, for the states a click takes a moment to draw.
const until = async (expression, budgetMs = 5000) => {
  let answer = await editor.js(expression);
  for (let waited = 0; !answer && waited < budgetMs; waited += 250) {
    await wait(250);
    answer = await editor.js(expression);
  }
  return answer;
};
// A change picked in the list: its row takes the selection's ground and its diff comes up beside
// it, line numbers and all. Answers whether both are really there, so a reading that follows is
// of that state and not of whatever the pane showed before.
const paneShows = () => editor.js(`(() => { const list = document.getElementById('changes-list'), diff = document.getElementById('changes-diff');
  return { rows: list.querySelectorAll('.changes-entry').length, picked: list.querySelectorAll('.changes-entry-showing').length,
    lineNumbers: diff.querySelectorAll('.sync-gutter').length, diff: diff.textContent.trim().slice(0, 48) }; })()`);
const pickAChange = async () => {
  const tries = [];
  for (let attempt = 1; attempt <= 3; attempt++) {
    // Clicked every time, picked already or not: a click on a row always shows its comparison
    // as it stands now, which is the pane's own rule.
    await editor.js(`document.querySelector('#changes-list .changes-entry')?.click(), true`);
    let shows = null;
    for (let waited = 0; waited < 2000; waited += 250) {
      await wait(250);
      shows = await paneShows();
      if (shows.picked > 0 && shows.lineNumbers > 0) {
        return { picked: true, said: attempt === 1 ? "" : `on click ${attempt}; before that ${JSON.stringify(tries)}` };
      }
    }
    tries.push(shows);
  }
  return { picked: false, said: JSON.stringify(tries) };
};

try {
  console.log("the Theme setting\n");

  for (const windows of ["dark", "light"]) {
    await emulate(windows);
    for (const choice of ["system", "light", "dark"]) {
      await api.settings({ theme: choice });
      await wait(800);
      const want = choice === "system" ? windows : choice;
      const shows = await pageShows();
      const label = `Windows ${windows}, Theme ${choice}`;
      check(`${label}: the page shows ${want}`, shows.theme === want && shows.editor === (want === "light" ? "vs" : "vs-dark"),
        JSON.stringify(shows));
      check(`${label}: the title bar is asked for ${want}`, titleBarAskedFor() === want, String(titleBarAskedFor()));
      check(`${label}: the Object Browser shows ${want}`, (await paletteShows()) === want);
      check(`${label}: the ground is ${want}`,
        shows.ground === (want === "light" ? "rgb(255, 255, 255)" : "rgb(30, 30, 30)"), shows.ground);
    }
  }
  check("the setting reads back as the last choice", (await api.settings()).theme === "dark");

  console.log("\nthe light surface's contrast\n");

  await api.settings({ theme: "light" });
  await wait(800);
  // The whole page unless a root is named: the light surface is read whole, dark only in the pane
  // whose rules have been brought to AA (see the head of this file). The reading is the one every
  // suite shares, in contrast.mjs; what comes back here is the runs that fell short.
  const scan = async (root) => (await editor.js(contrastScan(root))).fails;
  const surfaces = [
    ["the workspace", async () => {}],
    ["the Problems pane", () => editor.js(`document.querySelector('.panel-tab[data-panel="problems"]')?.click(), true`)],
    ["the Tests pane", () => editor.js(`document.querySelector('.panel-tab[data-panel="tests"]')?.click(), true`)],
    ["the Changes pane", () => editor.js(`document.querySelector('.panel-tab[data-panel="changes"]')?.click(), true`)],
    ["the Changes pane with a change in it", async () => {
      const shows = await writeAChange();
      // The instrument's own guard: without counts on the button, the reading below is of an
      // empty pane again and would pass whatever the stylesheet says.
      check("a written module puts its counts on the Changes summary", /^\+\d+\s*-\d+\s*\d+ modules?$/.test(shows ?? ""), JSON.stringify(shows));
    }],
    ["the Changes summary, opened", () => editor.js(`document.getElementById('changes-summary')?.click(), true`)],
    ["a change picked in the Changes pane", async () => {
      const pick = await pickAChange();
      check("a picked change shows its row selected and its diff", pick.picked, pick.said);
    }],
    ["the Settings dialog", () => api.act("settings")],
    ["the About dialog", async () => { await api.act("closeDialogs"); await api.act("toolbar", { command: "openHelp" }); }],
    ["the import and export dialog", async () => { await api.act("closeDialogs"); await api.act("toolbar", { command: "openSync" }); }],
    ["the menu", async () => { await api.act("closeDialogs"); await api.act("menuBar"); }],
  ];
  for (const [name, show] of surfaces) {
    await show();
    await wait(600);
    const fails = await scan();
    check(`${name} holds AA contrast in light`, fails.length === 0, fails.slice(0, 4).join("; "));
  }
  await api.act("closeDialogs");
  await editor.js(`document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })), true`);

  // THE SAME PANE IN DARK, while it still holds the change: the picked row on the selection's
  // ground with its diff beside it, then the summary opened over the diff strip.
  console.log("\nthe Changes pane in dark\n");
  await api.settings({ theme: "dark" });
  await wait(800);
  const inDark = [
    ["with a change picked", async () => {
      const pick = await pickAChange();
      check("the change is picked in dark as well", pick.picked, pick.said);
    }],
    ["with its summary opened", () => editor.js(`document.getElementById('changes-summary')?.click(), true`)],
  ];
  for (const [name, show] of inDark) {
    await show();
    await wait(600);
    const fails = await scan("#changes");
    check(`the Changes pane ${name} holds AA contrast in dark`, fails.length === 0, fails.slice(0, 6).join("; "));
  }
  check("the module the suite changed reads as it did before", await putTheTextBack());

  // THE ACCEPTED LINE. Pressed with the text back as it was, so nothing differs from the mark
  // and the summary stays down afterwards; the line itself is what is read, in both themes.
  await editor.js(`document.getElementById('changes-accept')?.click(), true`);
  const accepted = await until(`(() => { const line = document.querySelector('#changes-list .changes-accepted');
    return line && line.getBoundingClientRect().height > 0 ? line.textContent.trim() : null; })()`);
  check("Accept draws the accepted line", !!accepted, JSON.stringify(accepted));
  let fails = await scan("#changes");
  check("the Changes pane's accepted line holds AA contrast in dark", fails.length === 0, fails.slice(0, 4).join("; "));
  await api.settings({ theme: "light" });
  await wait(800);
  fails = await scan();
  check("the Changes pane's accepted line holds AA contrast in light", fails.length === 0, fails.slice(0, 4).join("; "));

  // THE RENAME SYMBOL BOX, typed into (#128). Monaco puts the input foreground on the box and
  // nothing told the input to take it, so a name typed into it was the browser's default black -
  // invisible on the dark input once the selection that hid it was gone. The scan above skips
  // inputs, which hold no text nodes, so the box is read on its own, in both themes, after typing.
  console.log("\nthe Rename Symbol box\n");
  await api.pane("open", { module: "Helper", project: project.projectId });
  await wait(1500);
  for (const theme of ["dark", "light"]) {
    await api.settings({ theme });
    await wait(800);
    await editor.js(`(() => { const e = window.xlideBridge.workspace.activeEditor(), m = e.getModel();
      for (let i = 1; i <= m.getLineCount(); i++) { const t = m.getLineContent(i), k = t.search(/(Sub|Function) [A-Za-z]/);
        if (k >= 0) { e.focus(); e.setPosition({ lineNumber: i, column: t.indexOf(' ', k) + 2 }); break; } }
      e.trigger('theme', 'editor.action.rename', {}); return true; })()`);
    await wait(1200);
    await editor.cdp("Input.insertText", { text: "Renamed" });
    await wait(400);
    const read = await editor.js(`(() => { const i = document.querySelector('.rename-box .rename-input'); if (!i) return null;
      const parse = (c) => /rgba?\\(([^)]+)\\)/.exec(c)[1].split(/[ ,/]+/).filter(Boolean).map(Number);
      const lum = (c) => { const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
        return 0.2126 * f(c[0]) + 0.7152 * f(c[1]) + 0.0722 * f(c[2]); };
      const s = getComputedStyle(i), [a, b] = [lum(parse(s.color)), lum(parse(s.backgroundColor))].sort((p, q) => q - p);
      return { value: i.value, color: s.color, ground: s.backgroundColor, ratio: (a + 0.05) / (b + 0.05) }; })()`);
    check(`${theme}: a name typed into the Rename box holds AA contrast`, read !== null && read.value === "Renamed" && read.ratio >= 4.5,
      read ? `${read.color} on ${read.ground}, ${read.ratio.toFixed(2)}:1` : "no rename box");
    for (const type of ["keyDown", "keyUp"]) {
      await editor.cdp("Input.dispatchKeyEvent", { type, key: "Escape", code: "Escape", windowsVirtualKeyCode: 27 });
    }
    await wait(400);
  }
} finally {
  // A no-op when the text already went back; the catch-all for a run that threw before it did.
  await putTheTextBack().catch((error) => console.log(`  (could not put ${changedModule} back: ${error.message})`));
  await api.settings({ theme: kept }).catch(() => {});
  await editor.cdp("Emulation.setEmulatedMedia", { features: [] }).catch(() => {});
  editor.socket.close();
}

process.exit(done());
