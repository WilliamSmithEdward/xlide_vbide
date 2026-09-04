/*
 * Introduce Parameter on the LIVE product: the signature grows and the callers follow, in the
 * modules the HOST holds.
 *
 * The second refactoring to write across modules, and the one whose failure would be quietest: a
 * signature that grew while a caller in another module still passes the old argument list compiles
 * nowhere and reports somewhere else entirely. So the caller here lives in its own module.
 *
 * Run against the debug fixture:
 *   node tools\harness\introduce-parameter.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const { check, done } = reporter();
const CRLF = "\r\n";
const stamp = process.pid % 10000;
const OWNER = `Param${stamp}`;
const CALLER = `ParamUse${stamp}`;

const OWNS = [
  "Option Explicit",                            // 1
  "",                                           // 2
  "Public Sub Post(ByVal label As String)",     // 3
  "    Dim rate As Double",                     // 4
  "    rate = 1.2",                             // 5
  "    Debug.Print label, rate",                // 6
  "End Sub",                                    // 7
  "",                                           // 8
  "Public Sub Nearby()",                        // 9
  '    Post "here"',                            // 10
  "End Sub",                                    // 11
];

const USES = [
  "Option Explicit",
  "",
  "Public Sub Far()",
  '    Post "there"',
  "End Sub",
];

async function hostText(name) {
  const answer = await api.readModule(name, project.projectId);
  return String(answer.source ?? answer.text ?? "").split(/\r\n|\r|\n/);
}

const has = (lines, text) => (lines ?? []).some((one) => one.trim() === text);

try {
  for (const name of [OWNER, CALLER]) {
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }

  for (const [name, lines] of [[OWNER, OWNS], [CALLER, USES]]) {
    await api.component("add", { kind: "module", name, project: project.projectId });
    await api.writeModule(name, lines.join(CRLF), project.projectId);
  }

  await api.caret(4, { module: OWNER, project: project.projectId, column: 9 });
  await waitFor("the page to be showing the module", async () =>
    ((await api.ui()).focus.model ?? "").toLowerCase().endsWith(`/${OWNER.toLowerCase()}`));
  await wait(300);

  /* ---- the lightbulb offers it on a local -------------------------------------------------------- */

  const titles = (answer) => (answer.data ?? []).map((one) => one.title);
  const offered = await api.act("quickFixes", { line: 4, column: 9 });
  check("a caret on a local is offered the parameter",
    titles(offered).includes("Make 'rate' a parameter"), titles(offered).join(" | ") || "(none)");

  /* ---- the menu path ------------------------------------------------------------------------------ */

  const made = await api.act("introduceParameter", { word: "rate" });
  check("the menu path makes it a parameter", made.did, made.detail);

  const owner = await hostText(OWNER);
  const caller = await hostText(CALLER);

  check("the host's own signature grew",
    has(owner, "Public Sub Post(ByVal label As String, ByVal rate As Double)"),
    owner.find((one) => one.includes("Sub Post")) ?? "(absent)");
  check("the declaration and the assignment are gone",
    !owner.some((one) => one.includes("Dim rate")) && !owner.some((one) => one.trim().startsWith("rate =")),
    owner.map((one) => one.trim()).filter(Boolean).join(" | "));
  check("the body still reads it",
    has(owner, "Debug.Print label, rate"));
  check("the caller in the same module passes the value",
    has(owner, 'Post "here", 1.2'), owner.find((one) => one.includes('"here"')) ?? "(absent)");
  check("and the one in ANOTHER module too, which is the half a single-module check cannot see",
    has(caller, 'Post "there", 1.2'), caller.map((one) => one.trim()).filter(Boolean).join(" | "));
  check("both call sites counted", made.data?.callSites === 2, JSON.stringify(made.data?.callSites));

  // THE MODULES THIS SUITE REWROTE, not every pane in the session.
  //
  // `parityAll` reads every open pane, and in a gate run most of them were opened by suites this
  // one has nothing to do with - so asserting on all of them makes this suite fail for somebody
  // else's stale pane, which is what happened on 2026-09-03. The session-wide question belongs to
  // three-copies.mjs, which exists for it. Anything else found stale is still REPORTED, so the
  // information is not lost, but it does not fail a suite that cannot own it.
  const mine = [OWNER, CALLER].map((one) => one.toLowerCase());
  const isMine = (one) => mine.includes((one.module ?? "").toLowerCase());

  // COUNTED, because `[].every()` is true. Filtering to this suite's modules and asking whether
  // they all agree passes VACUOUSLY if the filter matches nothing - a renamed module, a pane that
  // never opened, a project id that stopped lining up - and a check that cannot go red measures
  // nothing. So the panes have to BE there, all of them, before their agreement means anything.
  const parity = await waitFor("the rewritten modules to agree with the host", async () => {
    const now = await api.parityAll();
    const ours = now.panes.filter(isMine);
    return ours.length === mine.length && ours.every((one) => one.agreed) ? now : null;
  }, { budgetMs: 15000 }).catch(() => null);

  const others = ((parity ?? await api.parityAll()).stale ?? []).filter((one) => !isMine(one));
  check("the modules it rewrote agree with the host", parity !== null,
    parity === null
      ? `of ${mine.length} module(s) ${(await api.parityAll()).panes.filter(isMine).length} had a pane; `
        + `stale: ${((await api.parityAll()).stale ?? []).filter(isMine)
          .map((one) => one.module).join(", ") || "(none named)"}`
      : `${parity.panes.filter(isMine).length} of this suite's pane(s)`
        + (others.length > 0
          ? `; other suites left stale: ${others.map((one) => one.module).join(", ")}`
          : ""));

  const clean = await waitFor("the findings to settle", async () => {
    const mine = [OWNER, CALLER].map((one) => one.toLowerCase());
    const rows = ((await api.problems()).findings ?? [])
      .filter((one) => mine.includes((one.module ?? "").toLowerCase()));
    return rows.length === 0 ? [] : null;
  }, { budgetMs: 20000 }).catch(() => null);
  check("both modules analyse clean afterwards", clean !== null,
    clean === null
      ? ((await api.problems()).findings ?? [])
        .filter((one) => [OWNER, CALLER].map((n) => n.toLowerCase()).includes((one.module ?? "").toLowerCase()))
        .map((one) => `${one.module}:${one.code}`).join(",")
      : "(none)");

  /* ---- Ctrl+Z ------------------------------------------------------------------------------------- */

  await api.act("undo");
  const back = await waitFor("the signature to come back", async () => {
    const now = await hostText(OWNER);
    return now.some((one) => one.includes("ByVal rate As Double")) ? null : now;
  }, { budgetMs: 15000 }).catch(() => null);
  check("undo puts the signature back", back !== null,
    back === null ? "the parameter is still there" : `${back.length} line(s)`);
  check("and the caller in the other module too",
    back !== null && has(await hostText(CALLER), 'Post "there"'),
    (await hostText(CALLER)).map((one) => one.trim()).filter(Boolean).join(" | "));

  /* ---- the refusals --------------------------------------------------------------------------------- */

  await api.writeModule(OWNER, [
    "Option Explicit", "", "Public Sub Post()",
    "    Dim basePrice As Double", "    basePrice = 10",
    "    Dim rate As Double", "    rate = basePrice * 1.2",
    "    Debug.Print rate", "End Sub",
  ].join(CRLF), project.projectId);
  await wait(700);

  const stranded = await api.act("introduceParameter", { word: "rate" });
  check("a value naming a local is refused, and the refusal names it",
    !stranded.did && /'baseprice'/i.test(String(stranded.detail))
    && /not visible where the callers are/i.test(String(stranded.detail)),
    stranded.detail);

  const untouched = await hostText(OWNER);
  check("and a refusal wrote nothing",
    untouched.some((one) => one.includes("Dim rate")),
    untouched.map((one) => one.trim()).filter(Boolean).join(" | "));
} finally {
  for (const name of [CALLER, OWNER]) {
    await api.pane("close", { module: name, project: project.projectId, answer: "discard" }).catch(() => { });
    await api.component("remove", { name, project: project.projectId }).catch(() => { });
  }
}

done();
