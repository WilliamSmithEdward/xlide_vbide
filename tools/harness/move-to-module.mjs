/*
 * Move to Module on the LIVE product: the menu path, and every module the HOST holds afterwards.
 *
 * The first of these refactorings to write MORE THAN ONE module, which is the half its engine
 * suite cannot check: the procedure has to leave one component and arrive whole in another, and a
 * qualified call in a third has to follow it, through the host's object model rather than through
 * a returned string.
 *
 * Run against the debug fixture:
 *   node tools\harness\move-to-module.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const stamp = process.pid % 10000;
const SOURCE = `MoveFrom${stamp}`;
const TARGET = `MoveTo${stamp}`;
const CALLER = `MoveUse${stamp}`;

const FROM = [
  "Option Explicit",
  "",
  "Public Sub Stay()",
  '    Debug.Print "here"',
  "End Sub",
  "",
  "Public Sub Travel(ByVal n As Long)",
  "    Debug.Print n * 2",
  "End Sub",
];

const TO = ["Option Explicit", "", "Public Sub Already()", "End Sub"];

const USES = [
  "Option Explicit",
  "",
  "Public Sub Call1()",
  `    ${SOURCE}.Travel 1`,
  "    Travel 2",
  "End Sub",
];

async function hostText(name) {
  const answer = await api.readModule(name, project.projectId);
  return String(answer.source ?? answer.text ?? "").split(/\r\n|\r|\n/);
}

const has = (lines, text) => (lines ?? []).some((one) => one.trim() === text);

try {
  for (const name of [SOURCE, TARGET, CALLER]) {
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }

  for (const [name, lines] of [[SOURCE, FROM], [TARGET, TO], [CALLER, USES]]) {
    await api.component("add", { kind: "module", name, project: project.projectId });
    await api.writeModule(name, lines.join(CRLF), project.projectId);
  }

  await api.caret(7, { module: SOURCE, project: project.projectId });
  await waitFor("the page to be showing the module", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${SOURCE.toLowerCase()}`));

  /* ---- the lightbulb offers it inside a procedure ---------------------------------------------- */

  const titles = (answer) => (answer.data ?? []).map((one) => one.title);
  const offered = await api.act("quickFixes", { line: 7, column: 5 });
  check("a caret inside a procedure is offered the move",
    titles(offered).includes("Move to module..."), titles(offered).join(" | ") || "(none)");
  // A `Public Sub` line is a procedure, not a variable called `Sub`: the encapsulate entry has no
  // business on it, and did until this suite put a caret on one.
  check("and is not offered to encapsulate the word Sub",
    !titles(offered).some((one) => one.startsWith("Encapsulate ")),
    titles(offered).join(" | ") || "(none)");

  /* ---- the menu path ---------------------------------------------------------------------------- */

  const made = await api.act("moveToModule", { line: 7, column: 5, targetModule: TARGET });
  check("the menu path moves it", made.did, made.detail);

  const left = await hostText(SOURCE);
  const arrived = await hostText(TARGET);
  const caller = await hostText(CALLER);

  check("the host's own source module lost it",
    !left.some((one) => one.includes("Sub Travel")),
    left.map((one) => one.trim()).filter(Boolean).join(" | "));
  check("and kept what was not moving",
    has(left, "Public Sub Stay()"), left.map((one) => one.trim()).filter(Boolean).join(" | "));
  check("the target module has it whole",
    has(arrived, "Public Sub Travel(ByVal n As Long)")
    && has(arrived, "Debug.Print n * 2") && has(arrived, "End Sub"),
    arrived.map((one) => one.trim()).filter(Boolean).join(" | "));
  check("and kept what it already had",
    has(arrived, "Public Sub Already()"));
  check("the qualified call site followed it",
    has(caller, `${TARGET}.Travel 1`), caller.map((one) => one.trim()).filter(Boolean).join(" | "));
  check("and the unqualified one was left alone",
    has(caller, "Travel 2"), caller.map((one) => one.trim()).filter(Boolean).join(" | "));

  const parity = await waitFor("the surface to agree with the host",
    async () => { const now = await api.parityAll(); return now.agreed ? now : null; },
    { budgetMs: 15000 }).catch(() => null);
  check("every open module's surface agrees with the host", parity !== null,
    parity ? `${parity.panes.length} pane(s)` : "still disagreeing");

  const clean = await waitFor("the findings to settle", async () => {
    const mine = [SOURCE, TARGET, CALLER].map((one) => one.toLowerCase());
    const rows = ((await api.problems()).findings ?? [])
      .filter((one) => mine.includes((one.module ?? "").toLowerCase()));
    return rows.length === 0 ? [] : null;
  }, { budgetMs: 20000 }).catch(() => null);
  check("all three modules analyse clean afterwards", clean !== null,
    clean === null
      ? ((await api.problems()).findings ?? [])
        .filter((one) => [SOURCE, TARGET, CALLER].map((n) => n.toLowerCase())
          .includes((one.module ?? "").toLowerCase()))
        .map((one) => `${one.module}:${one.code}`).join(",")
      : "(none)");

  /* ---- the refusals ------------------------------------------------------------------------------ */

  await api.writeModule(SOURCE, [
    "Option Explicit", "", "Private held As Long", "",
    "Public Sub Stranded()", "    Debug.Print held", "End Sub",
  ].join(CRLF), project.projectId);
  await wait(700);
  await api.caret(5, { module: SOURCE, project: project.projectId });
  await wait(300);

  const stranded = await api.act("moveToModule", { line: 5, column: 5, targetModule: TARGET });
  check("a procedure using a Private of its module is refused, and the refusal names it",
    !stranded.did && /'held'/.test(String(stranded.detail))
    && /Private to/.test(String(stranded.detail)),
    stranded.detail);

  const untouched = await hostText(SOURCE);
  check("and a refusal wrote nothing",
    untouched.some((one) => one.includes("Sub Stranded")),
    untouched.map((one) => one.trim()).filter(Boolean).join(" | "));

  const nowhere = await api.act("moveToModule", { line: 5, column: 5, targetModule: "NoSuchModule" });
  check("a module the project does not have is refused",
    !nowhere.did && /no module called 'NoSuchModule'/.test(String(nowhere.detail)), nowhere.detail);
} finally {
  await api.act("extractDialog", { press: "cancel" }).catch(() => { });
  for (const name of [CALLER, TARGET, SOURCE]) {
    await api.pane("close", { module: name, project: project.projectId, answer: "discard" }).catch(() => { });
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }
}

done();
