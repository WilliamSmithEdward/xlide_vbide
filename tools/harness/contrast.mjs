/*
 * Can it be read? One instrument, for every suite that asks.
 *
 * Every visible text run under a root is held to WCAG AA against the ground it is drawn on: 4.5:1
 * for text, 3:1 for large text and icons. The ground is composited through the run's ancestors'
 * backgrounds and the colour through their opacity, because that is where the failures this was
 * written for live: a status colour in a word that is also dimmed, a count on a button that rests
 * at 0.8, a quiet label on a picked row's lighter ground. A rule read on its own shows none of
 * them.
 *
 * Exempt, as WCAG exempts them: disabled controls. Skipped: the code editor's own text, the
 * minimap, canvases and pictures. What is scrolled out of view is skipped when the whole page
 * is read, where a rect off the edge is as likely to be something parked there as something to
 * read, and TAKEN IN when one surface is: a pane's rows below the fold are the pane, and a
 * reading that counted only what a small window happened to show read 12 runs of Source Control
 * where the pane holds 52 - and missed every status word (2026-11-25, the first run of
 * scm.mjs's check, in a window with no saved size). An element laid out with
 * `display: contents` has no box and paints no background, whatever its style says, so its tint is
 * not counted into the ground - counted, it read the full-size comparison's changed lines as
 * twice as tinted as they are.
 *
 * A READING IS OF A STATE. An empty pane holds one sentence and passes whatever its rules say. The
 * Changes pane failed only with rows in it, on a picked row, with a diff showing, and after
 * Accept; the import dialog's "same" chip failed worst on a picked row (2026-11-25). So a suite
 * reads a surface at a point where it has already built the state worth reading, and the reading
 * says how many runs it took in, so a reading of nothing is seen for what it is.
 *
 * WHO READS WHAT. theme.mjs reads the light surface whole, and the Changes pane in dark.
 * module-sync.mjs reads the import and export dialog and scm.mjs the Source Control pane, each in
 * both themes and each scoped to its own surface: the rest of the dark surface has not been
 * audited, and a whole-page reading there would go red on things no suite has a claim on yet.
 */

import { wait } from "./xlide-api.mjs";

/**
 * The reading, as an expression for the page to evaluate. Answers `{ read, fails }`: how many
 * text runs it held to the standard, and one line for each that fell short - the ratio, the
 * element's id or classes, and the start of its text. `outOfView` takes in what is scrolled
 * out of the window as well; see the head of this file for when.
 */
export const contrastScan = (root = "body", { outOfView = false } = {}) => `(() => {
  const parse = (c) => { const m = /rgba?\\(([^)]+)\\)/.exec(c); if (!m) return null;
    const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number); return [p[0], p[1], p[2], p.length > 3 ? p[3] : 1]; };
  const over = (top, bottom) => [0, 1, 2].map((i) => top[i] * top[3] + bottom[i] * (1 - top[3])).concat(1);
  const lum = (c) => { const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
    return 0.2126 * f(c[0]) + 0.7152 * f(c[1]) + 0.0722 * f(c[2]); };
  const ratio = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
  const groundOf = (el) => { const chain = [];
    for (let e = el; e; e = e.parentElement) { const s = getComputedStyle(e); if (s.display === 'contents') continue;
      const bg = parse(s.backgroundColor);
      if (bg && bg[3] > 0) chain.push(bg); if (bg && bg[3] >= 1) break; }
    let ground = parse(getComputedStyle(document.body).backgroundColor) ?? [255, 255, 255, 1];
    for (const bg of chain.reverse()) ground = over(bg, ground); return ground; };
  const shown = (el) => { const r = el.getBoundingClientRect();
    if (r.width < 1 || r.height < 1) return false;
    if (${outOfView ? "false" : "true"} && (r.bottom < 0 || r.top > innerHeight || r.right < 0 || r.left > innerWidth)) return false;
    for (let e = el; e; e = e.parentElement) { const s = getComputedStyle(e);
      if (s.visibility === 'hidden' || s.display === 'none' || Number(s.opacity) < 0.05) return false; } return true; };
  const opacityOf = (el) => { let o = 1; for (let e = el; e; e = e.parentElement) o *= Number(getComputedStyle(e).opacity); return o; };
  const fails = [];
  let read = 0;
  for (const el of document.querySelectorAll(${JSON.stringify(`${root} *`)})) {
    if (el.closest('.monaco-editor .view-lines, .minimap, canvas, svg, .designer-canvas')) continue;
    const icon = /codicon/.test(typeof el.className === 'string' ? el.className : '');
    if (!icon && ![...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim())) continue;
    if (!shown(el) || el.closest('[disabled], [aria-disabled="true"], .disabled')) continue;
    const s = getComputedStyle(el), fg = parse(s.color); if (!fg) continue;
    const ground = groundOf(el), r = ratio(over([fg[0], fg[1], fg[2], fg[3] * opacityOf(el)], ground), ground);
    const size = parseFloat(s.fontSize), large = size >= 24 || (Number(s.fontWeight) >= 700 && size >= 18.66);
    const need = icon || large ? 3 : 4.5;
    read++;
    if (r < need) fails.push(r.toFixed(2) + ' ' + (el.id ? '#' + el.id : el.className) + ' "' + (icon ? 'icon' : el.textContent.trim().slice(0, 24)) + '"');
  }
  return { read, fails: [...new Set(fails)] };
})()`;

/**
 * Reads one surface in each theme named, through the door, and puts the developer's Theme back
 * whatever happens. Answers `{ dark: { read, fails }, light: { read, fails } }`. The whole
 * surface is read, in view or not, so the answer does not depend on the size of the window.
 *
 * For a suite that has the state on screen and is not otherwise about themes. The page must not
 * be reloaded or its surface closed while this runs; nothing here does either.
 */
export async function contrastByTheme(api, root, themes = ["dark", "light"]) {
  const kept = (await api.settings()).theme;
  const readings = {};
  try {
    for (const theme of themes) {
      await api.settings({ theme });
      await wait(800);
      // As a string, so the answer is the same whichever way the door hands a value back.
      const answer = await api.ask(`JSON.stringify(${contrastScan(root, { outOfView: true })})`);
      readings[theme] = typeof answer === "string" ? JSON.parse(answer) : answer;
    }
  } finally {
    await api.settings({ theme: kept });
    await wait(400);
  }
  return readings;
}

/** A reading as one line for a check's detail: how much was read, and what fell short. */
export const contrastSaid = (reading) =>
  `${reading?.read ?? 0} text runs read${reading?.fails?.length ? `; ${reading.fails.slice(0, 6).join("; ")}` : ""}`;
