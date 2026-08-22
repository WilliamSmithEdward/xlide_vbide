/*
 * The change log: what happened to this project's module code, by whom, in rounds.
 *
 * WHAT IT IS FOR. An agent has been editing this workbook through the door for twenty minutes -
 * what did it change? Before this there was no answer but the shim's diagnostic log, which says a
 * write happened and not what it did.
 *
 * WHAT IT IS NOT. There is no revert here and there is not going to be one: the log shows, and
 * putting text back is a WRITE, done through the route that already has one. That is what makes
 * it safe to trust - nothing in it can lose work, because nothing in it writes any. So every
 * check below reads; none of them asks the log to change the project.
 *
 * The questions only this can ask:
 *
 *   - does a round say what it did to a module ONCE, however many times the module was rewritten?
 *   - are two hands two rounds, so a developer's own edit is not folded into an agent's?
 *   - is the text a module held BEFORE a round still there to be read?
 *   - does the pane draw what the route answers, or has one of the two drifted?
 *   - and does a module nobody touched stay out of it?
 *
 * Run against ChangeFixture.xlsm with the editor open:
 *   tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\ChangeFixture.xlsm
 *   node tools\harness\change-log.mjs
 */

import { open, waitFor } from "./xlide-api.mjs";

const api = await open();
let passed = 0;
let failed = 0;

const check = (name, got, want = true) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (ok) {
    passed++;
    console.log(`ok   ${name}`);
  } else {
    failed++;
    console.log(`FAIL ${name}\n       got  ${JSON.stringify(got)}\n       want ${JSON.stringify(want)}`);
  }
  return ok;
};

const project = (await api.projects()).projects[0];
console.log(`project: ${project.project}\n`);

const held = (name) => api.readModule(name, project.projectId).then((one) => one.text);
const write = (name, text, by) => api.writeModule(name, text, project.projectId, { by });
const log = (args = {}) => api.changes({ ...args, project: project.projectId });

// What every check below is measured against. The fixture's own text, before this run.
const ledgerWas = await held("Ledger");
const untouchedWas = await held("Untouched");

// A LINE OF OUR OWN FIRST. The fixture was built through this door, so the log already holds the
// round that created it; closing here means every round below belongs to this run.
await log({ action: "snapshot", label: "before change-log.mjs" });
const from = (await log()).rounds[0]?.round ?? 0;

const since = (answer) => answer.rounds.filter((round) => round.round > from);

// ---- one hand, three writes ------------------------------------------------------------------

// A module's text comes back with no line ending after its last line, so an appended line takes
// one in front and none behind. `${was}\r\n' one\r\n` would add a BLANK line as well, and the log
// would be right to count it - which is how the first draft of this suite was wrong and the
// product was not.
await write("Ledger", `${ledgerWas}\r\n' one`, "claude");
await write("Ledger", `${ledgerWas}\r\n' one\r\n' two`, "claude");
await write("Ledger", `${ledgerWas}\r\n' one\r\n' two\r\n' three`, "claude");

const oneHand = since(await log());
check("three writes by one hand are one round", oneHand.length, 1);
check("and that round names the hand", oneHand[0]?.by, "claude");

check("the round says what it did to the module once",
  oneHand[0]?.entries.map((entry) => `${entry.module} +${entry.added} -${entry.removed}`),
  ["Ledger +3 -0"]);

check("a module nobody wrote to is not in it",
  oneHand[0]?.entries.some((entry) => entry.module === "Untouched"), false);

// ---- the text from before ----------------------------------------------------------------------

const before = await log({ action: "text", round: oneHand[0].round, module: "Ledger", which: "before" });
check("the text the module held before the round is still readable", before.held, true);
check("and it is exactly what the module held", before.text === ledgerWas);

const after = await log({ action: "text", round: oneHand[0].round, module: "Ledger", which: "after" });
check("the text it holds after is the last of the three writes",
  after.text === `${ledgerWas}\r\n' one\r\n' two\r\n' three`);

const lined = await log({ action: "diff", round: oneHand[0].round, module: "Ledger" });
check("the change lines up, and only the new lines are added",
  lined.rows.filter((row) => row.kind !== "equal").map((row) => row.right),
  ["' one", "' two", "' three"]);

// ---- a different hand is a different round -----------------------------------------------------

await write("Ledger", `${ledgerWas}\r\n' one\r\n' two\r\n' three\r\n' by hand`, "developer");
await write("Ticket", `${await held("Ticket")}\r\n' back to the agent`, "claude");

const hands = since(await log());
check("agent, then developer, then agent again is three rounds", hands.length, 3);
check("and the log says which was which",
  hands.map((round) => round.by).reverse(), ["claude", "developer", "claude"]);

// ---- a snapshot ends a round and names it ------------------------------------------------------

await log({ action: "snapshot", label: "a labelled round" });
const labelled = since(await log());
check("a snapshot labels the round that was running", labelled[0]?.label, "a labelled round");
check("and closes it", labelled[0]?.open, false);

const empty = since(await log({ action: "snapshot", label: "nothing happened" }));
check("a snapshot with nothing running invents no round", empty.length, labelled.length);

// ---- the pane draws what the route says --------------------------------------------------------

await api.ask(
  `(() => { const tab = document.querySelector('.panel-tab[data-panel="changes"]');`
  + ` if (tab) { tab.click(); } return !!tab; })()`);

const drawn = await waitFor("the pane to read the log", async () => {
  const ui = await api.ui();
  return ui.changes && ui.changes.rounds.length > 0 ? ui.changes : false;
}, { budgetMs: 20000 });

const routeNow = await log();
check("the pane and the route agree about the rounds",
  drawn.rounds.map((round) => `${round.round}:${round.by}`),
  routeNow.rounds.map((round) => `${round.round}:${round.by}`));

check("the pane and the route agree about the counts",
  drawn.rounds.flatMap((round) => round.modules.map((one) => `${one.module}+${one.added}-${one.removed}`)),
  routeNow.rounds.flatMap((round) => round.entries.map((one) => `${one.module}+${one.added}-${one.removed}`)));

check("the pane says what the log does not cover", drawn.covers, routeNow.covers);
check("and what it does not cover is said out loud",
  drawn.covers.includes("form designs") && drawn.covers.includes("directly in the VBE"));

// ---- and the pane opens one module's comparison -------------------------------------------------

const shown = await api.act("changesPane", { round: oneHand[0].round, module: "Ledger" });
check("the pane opens a module's comparison from its own row", shown.did, true);

const rows = await waitFor("the comparison to draw", async () => {
  const said = await api.ask(
    `[...document.querySelectorAll('#changes-diff .sync-diff-row')].length`);
  const count = typeof said === "number" ? said : Number(said);
  return count > 0 ? count : false;
}, { budgetMs: 15000 });
check("and it draws rows", rows > 0);

// ---- the accepted line -------------------------------------------------------------------------

const accepted = await log({ action: "accept" });
check("accepting draws a line at the newest round", accepted.acceptedAt, accepted.rounds[0]?.round);
check("and destroys nothing", since(accepted).length, since(routeNow).length);

// ---- nothing here wrote to the project ----------------------------------------------------------

check("Untouched is exactly as the fixture built it", (await held("Untouched")) === untouchedWas);

// Put the fixture's own module back, so a second run starts where the first did.
await write("Ledger", ledgerWas, "change-log.mjs cleanup");
await write("Ticket", (await held("Ticket")).replace("\r\n' back to the agent", ""), "change-log.mjs cleanup");

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
