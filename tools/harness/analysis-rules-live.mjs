/*
 * Machine-wide analyzer rule settings on the LIVE product: the catalog over the door, an
 * override moving the Problems pane, the settings file carrying it, the refusal for an illegal
 * move, the lightbulb's machine-wide entry, and the inline-suppression writer - the same
 * mechanism the rules modal and the problems pane's menu run.
 *
 * Every wait names the state it waits FOR, and a timeout is a failure, never a settle.
 *
 * Run against the debug fixture:
 *   node tools\harness\analysis-rules-live.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const NAME = `Rules${process.pid % 10000}`;

async function problemsFor(module, expected) {
  let rows = [];
  const until = Date.now() + 25000;
  while (Date.now() < until) {
    await wait(600);
    const now = (await api.problems().catch(() => ({ findings: [] }))).findings ?? [];
    rows = now.filter((one) => (one.module ?? "").toLowerCase() === module.toLowerCase());
    if (expected(rows)) { return { rows, arrived: true }; }
  }
  return { rows, arrived: false };
}

const codes = (rows) => rows.map((one) => one.code).sort().join(",") || "(none)";
const settingsFile = () => join(process.env.LOCALAPPDATA, "xlide_vbide", "settings.json");
const settingsText = () => {
  try { return readFileSync(settingsFile(), "utf8"); } catch { return ""; }
};

try {
  await api.command("reset").catch(() => {});
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => {});
  await api.component("add", { kind: "module", name: NAME, project: project.projectId });

  // A clean slate, whatever an interrupted earlier run left standing.
  await api.analysis({ rule: "option-explicit-missing", severity: "default" }).catch(() => {});

  /* ---- the catalog over the door -------------------------------------------------------------- */

  const catalog = await api.analysis();
  check("the catalog reaches the door, from the bundled analyzer",
    (catalog.rules ?? []).length >= 100, `${(catalog.rules ?? []).length} rule(s)`);
  check("and it names the settings file the choices persist in",
    String(catalog.settingsPath ?? "").toLowerCase().endsWith("settings.json"),
    catalog.settingsPath);

  /* ---- an override moves the pane -------------------------------------------------------------- */

  // No Option Explicit, deliberately: the one finding is the style rule the owner named.
  await api.writeModule(NAME, [
    "Public Sub Go()", "    Dim n As Long", "    n = 1", "End Sub",
  ].join(CRLF), project.projectId);
  const before = await problemsFor(NAME,
    (rows) => rows.some((one) => one.code === "option-explicit-missing"));
  check("the style finding reaches the pane", before.arrived, codes(before.rows));

  const turnedOff = await api.analysis({ rule: "option-explicit-missing", severity: "off" });
  check("turning it off answers in words",
    /off everywhere on this machine/.test(String(turnedOff.detail)), turnedOff.detail);
  check("and the reply's catalog row carries the standing override",
    turnedOff.rules.find((one) => one.code === "option-explicit-missing")?.override === "off");

  const after = await problemsFor(NAME, (rows) => rows.length === 0);
  check("the pane empties without the module being touched", after.arrived, codes(after.rows));

  check("the settings file in user space carries the choice",
    settingsText().includes('"option-explicit-missing": "off"'),
    settingsFile());

  /* ---- the lightbulb's machine-wide entry ------------------------------------------------------ */

  // A second module still showing the finding needs the rule back on first.
  const restored = await api.analysis({ rule: "option-explicit-missing", severity: "default" });
  check("clearing answers in words",
    /back to its default/.test(String(restored.detail)), restored.detail);
  const back = await problemsFor(NAME,
    (rows) => rows.some((one) => one.code === "option-explicit-missing"));
  check("and the finding comes back - nothing was suppressed in the code", back.arrived,
    codes(back.rows));

  await api.pane("open", { module: NAME, project: project.projectId });
  await wait(800);
  // AT (1,1), where the module-scoped finding anchors. The first run aimed at column 8 and the
  // marker's zero-width span at column 1 did not touch the word there, so the whole provider
  // answered nothing and the check measured the probe's aim rather than the product.
  const fixes = await api.act("quickFixes", { line: 1, column: 1 });
  const titles = (fixes.data ?? []).map((one) => one.title);
  check("the lightbulb offers the machine-wide switch beside the inline one",
    titles.some((title) => /Turn off 'option-explicit-missing' on this machine/.test(title)),
    titles.join(" | ").slice(0, 140));

  /* ---- illegal moves are refused in words ------------------------------------------------------ */

  const illegal = await api.analysis({ rule: "undeclared-variable", severity: "off" });
  check("an error rule cannot be turned off, and the refusal names the legal move",
    /cannot be 'off'/.test(String(illegal.detail)) && /warning/.test(String(illegal.detail)),
    illegal.detail);

  const unknown = await api.analysis({ rule: "no-such-rule", severity: "off" });
  check("an unknown rule is refused by name",
    /is not an analyzer rule/.test(String(unknown.detail)), unknown.detail);

  /* ---- the inline writer, the problems menu's mechanism ---------------------------------------- */

  const suppressed = await api.analysis({
    module: NAME, line: 1, code: "option-explicit-missing", project: project.projectId,
  });
  check("the inline writer answers in words", /suppressed option-explicit-missing/.test(String(suppressed.detail)),
    suppressed.detail);

  const text = (await api.readModule(NAME, project.projectId)).text ?? "";
  check("the directive is the FILE one, because the rule is module-scoped",
    text.startsWith("' @xlide-analysis-disable-file option-explicit-missing"),
    JSON.stringify(text.split(/\r?\n/)[0]));

  const gone = await problemsFor(NAME, (rows) => rows.length === 0);
  check("and the pane empties from the comment alone", gone.arrived, codes(gone.rows));
} finally {
  await api.analysis({ rule: "option-explicit-missing", severity: "default" }).catch(() => {});
  await api.command("reset").catch(() => {});
  await api.component("remove", { name: NAME, project: project.projectId }).catch(() => {});
}

process.exit(done());
