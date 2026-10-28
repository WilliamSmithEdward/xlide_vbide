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
  const scan = `(() => {
    const parse = (c) => { const m = /rgba?\\(([^)]+)\\)/.exec(c); if (!m) return null;
      const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number); return [p[0], p[1], p[2], p.length > 3 ? p[3] : 1]; };
    const over = (top, bottom) => [0, 1, 2].map((i) => top[i] * top[3] + bottom[i] * (1 - top[3])).concat(1);
    const lum = (c) => { const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
      return 0.2126 * f(c[0]) + 0.7152 * f(c[1]) + 0.0722 * f(c[2]); };
    const ratio = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
    const groundOf = (el) => { const chain = [];
      for (let e = el; e; e = e.parentElement) { const bg = parse(getComputedStyle(e).backgroundColor);
        if (bg && bg[3] > 0) chain.push(bg); if (bg && bg[3] >= 1) break; }
      let ground = parse(getComputedStyle(document.body).backgroundColor) ?? [255, 255, 255, 1];
      for (const bg of chain.reverse()) ground = over(bg, ground); return ground; };
    const shown = (el) => { const r = el.getBoundingClientRect();
      if (r.width < 1 || r.height < 1 || r.bottom < 0 || r.top > innerHeight || r.right < 0 || r.left > innerWidth) return false;
      for (let e = el; e; e = e.parentElement) { const s = getComputedStyle(e);
        if (s.visibility === 'hidden' || s.display === 'none' || Number(s.opacity) < 0.05) return false; } return true; };
    const opacityOf = (el) => { let o = 1; for (let e = el; e; e = e.parentElement) o *= Number(getComputedStyle(e).opacity); return o; };
    const fails = [];
    for (const el of document.querySelectorAll('body *')) {
      if (el.closest('.monaco-editor .view-lines, .minimap, canvas, svg, .designer-canvas')) continue;
      const icon = /codicon/.test(typeof el.className === 'string' ? el.className : '');
      if (!icon && ![...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim())) continue;
      if (!shown(el) || el.closest('[disabled], [aria-disabled="true"], .disabled')) continue;
      const s = getComputedStyle(el), fg = parse(s.color); if (!fg) continue;
      const ground = groundOf(el), r = ratio(over([fg[0], fg[1], fg[2], fg[3] * opacityOf(el)], ground), ground);
      const size = parseFloat(s.fontSize), large = size >= 24 || (Number(s.fontWeight) >= 700 && size >= 18.66);
      const need = icon || large ? 3 : 4.5;
      if (r < need) fails.push(r.toFixed(2) + ' ' + (el.id ? '#' + el.id : el.className) + ' "' + (icon ? 'icon' : el.textContent.trim().slice(0, 24)) + '"');
    }
    return [...new Set(fails)];
  })()`;
  const surfaces = [
    ["the workspace", async () => {}],
    ["the Problems pane", () => editor.js(`document.querySelector('.panel-tab[data-panel="problems"]')?.click(), true`)],
    ["the Tests pane", () => editor.js(`document.querySelector('.panel-tab[data-panel="tests"]')?.click(), true`)],
    ["the Changes pane", () => editor.js(`document.querySelector('.panel-tab[data-panel="changes"]')?.click(), true`)],
    ["the Settings dialog", () => api.act("settings")],
    ["the About dialog", async () => { await api.act("closeDialogs"); await api.act("toolbar", { command: "openHelp" }); }],
    ["the import and export dialog", async () => { await api.act("closeDialogs"); await api.act("toolbar", { command: "openSync" }); }],
    ["the menu", async () => { await api.act("closeDialogs"); await api.act("menuBar"); }],
  ];
  for (const [name, show] of surfaces) {
    await show();
    await wait(600);
    const fails = await editor.js(scan);
    check(`${name} holds AA contrast in light`, fails.length === 0, fails.slice(0, 4).join("; "));
  }
  await api.act("closeDialogs");
  await editor.js(`document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })), true`);

  // THE RENAME SYMBOL BOX, typed into (#128). Monaco puts the input foreground on the box and
  // nothing told the input to take it, so a name typed into it was the browser's default black -
  // invisible on the dark input once the selection that hid it was gone. The scan above skips
  // inputs, which hold no text nodes, so the box is read on its own, in both themes, after typing.
  console.log("\nthe Rename Symbol box\n");
  const project = await api.project();
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
  await api.settings({ theme: kept }).catch(() => {});
  await editor.cdp("Emulation.setEmulatedMedia", { features: [] }).catch(() => {});
  editor.socket.close();
}

process.exit(done());
