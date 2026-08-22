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
 * AND IT WRITES NO FILE. The log is held in memory for the life of the session, because nothing in
 * production may depend on an external log file (the owner, 2026-08-22). That is why nothing here
 * looks on disk for anything.
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

// POINTED AT THE FILE THIS SUITE IS WRITING TO. With a second workbook open the pane starts on
// whichever the developer is in, which need not be this one - and that is the pane being right,
// not the suite: every project keeps its own log.
await api.act("changesPane", { file: project.project });

const drawn = await waitFor("the pane to read the log", async () => {
  const ui = await api.ui();
  return ui.changes && ui.changes.project === project.project && ui.changes.rounds.length > 0
    ? ui.changes
    : false;
}, { budgetMs: 20000 });

const routeNow = await log();
check("the pane and the route agree about the rounds",
  drawn.rounds.map((round) => `${round.round}:${round.by}`),
  routeNow.rounds.map((round) => `${round.round}:${round.by}`));

check("the pane and the route agree about the counts",
  drawn.rounds.flatMap((round) => round.modules.map((one) => `${one.module}+${one.added}-${one.removed}`)),
  routeNow.rounds.flatMap((round) => round.entries.map((one) => `${one.module}+${one.added}-${one.removed}`)));

check("the pane says what the log does not cover", drawn.covers, routeNow.covers);
// BOTH HALVES. The exclusions matter, and so does the part a reader doubts: their own typing is
// recorded, under their own name. The first wording led with the exclusions and called the
// uncovered case "edits made directly in the VBE", which the owner read - correctly, from inside
// the VBE - as "anything I type", while their typing was in the log all along.
check("it says what IS recorded, the developer's own edits included",
  drawn.covers.includes("yours and an agent's alike"));
check("and what is not", drawn.covers.includes("form designs"));

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

// ---- and it opens full size ---------------------------------------------------------------------
//
// The pane is a strip along the bottom and a comparison is two columns of code, so anything past a
// few lines is read three words at a time down there.

const grew = await api.act("changesPane", { expand: true });
check("the comparison opens full size", grew.did, true);

const bigger = await api.ask(`JSON.stringify({
  card: !!document.querySelector('#changes-full-card'),
  title: document.querySelector('#changes-full-head span')?.textContent ?? '',
  closes: !!document.querySelector('#changes-full-close'),
  rows: [...document.querySelectorAll('#changes-full-diff .sync-diff-row')].length,
  clipped: getComputedStyle(document.querySelector('#changes-full-diff .sync-code')).textOverflow
})`);
const full = JSON.parse(typeof bigger === "string" ? bigger : JSON.stringify(bigger));

check("it draws the same rows, with a way out", { card: full.card, closes: full.closes, drew: full.rows > 0 },
  { card: true, closes: true, drew: true });
check("and the code is not clipped, because reading it is the point", full.clipped, "clip");
check("the pane says it is up", (await api.ui()).changes.full, true);

const shrank = await api.act("changesPane", { expand: false });
check("and it closes", { did: shrank.did, up: (await api.ui()).changes.full }, { did: true, up: false });

// ---- accepting -------------------------------------------------------------------------

const accepted = await log({ action: "accept" });
check("accepting marks the newest round as reviewed", accepted.acceptedAt, accepted.rounds[0]?.round);
check("and destroys nothing", since(accepted).length, since(routeNow).length);

// ---- a module arriving, being renamed, and leaving --------------------------------------------
//
// Only the write path recorded anything at first, so a module could be added, filled and removed
// and the log would show the filling with no sign of either end (the owner asked, 2026-08-22). A
// removal is the one change whose "before" cannot be recovered afterwards, which makes it the one
// most worth having.

const born = `Arrived${process.pid}`;
const renamed = `Renamed${process.pid}`;

await api.changes({ action: "snapshot", label: "before the component checks", project: project.projectId });
const beforeComponents = (await log()).rounds[0]?.round ?? 0;
const componentsSince = (answer) => answer.rounds.filter((round) => round.round > beforeComponents);

await api.component("add", { name: born, kind: "module", project: project.projectId });
await write(born, ["Option Explicit", "", "Public Sub Hello()", "End Sub"].join("\r\n"), "claude");

// Across the rounds, not within one: the add came through the door unattributed and the write
// named itself, so they are two rounds by the hand-change rule - which is the rule working.
const arrivals = componentsSince(await log()).flatMap((round) => round.entries)
  .filter((entry) => entry.module.toLowerCase() === born.toLowerCase());
check("a module arriving is recorded as an add",
  arrivals.some((entry) => entry.kind === "added"), true);

await api.component("rename", { name: born, newName: renamed, project: project.projectId });

const moved = componentsSince(await log()).flatMap((round) => round.entries)
  .find((entry) => entry.module.toLowerCase() === renamed.toLowerCase());
check("a rename moves the entry rather than starting another",
  moved ? `${moved.module} from ${moved.from}` : "(not recorded)", `${renamed} from ${born}`);

const wasHolding = await held(renamed);
await api.component("remove", { name: renamed, project: project.projectId });

const gone = componentsSince(await log()).flatMap((round) => round.entries)
  .find((entry) => entry.module.toLowerCase() === renamed.toLowerCase());
check("a module leaving is recorded", gone ? gone.kind : "(not recorded)", "removed");

// THE ONE TEXT A REMOVAL CANNOT GET BACK. The round that recorded the removal is the newest one
// naming the module, and its `before` is what the module held when it went.
const burial = componentsSince(await log()).find(
  (round) => round.entries.some((entry) =>
    entry.module.toLowerCase() === renamed.toLowerCase() && entry.kind === "removed"));

const lastWords = await log({
  action: "text", round: burial?.round ?? 0, module: renamed, which: "before",
});
check("and what it held is still readable after it is gone", lastWords.held, true);
check("and it is what the module actually held", lastWords.text === wasHolding);

// ---- the pane is scoped to ONE file ---------------------------------------------------------------
//
// Every project keeps its own log, so the pane's file select is choosing between logs rather than
// filtering one. With a second workbook open the select appears; with one it stays out of the way,
// the same rule the list panes follow. And a file that CLOSES leaves the list, which is what makes
// its changes drop out of the pane rather than lingering as an answer about a workbook nobody has.
const session = (await api.projects()).projects;
const paneNow = (await api.ui()).changes;

check("the pane offers every open file",
  [...paneNow.files].sort(), session.map((one) => one.project).sort());

check("and it is pointed at one of them", session.some((one) => one.project === paneNow.project));

if (session.length > 1) {
  const other = paneNow.files.find((name) => name !== paneNow.project);
  const moved = await api.act("changesPane", { file: other });
  check("the select points the pane at another open file", moved.did, true);

  const after = await waitFor("the pane to answer for the other file", async () => {
    const now = (await api.ui()).changes;
    return now && now.project === other ? now : false;
  }, { budgetMs: 20000 });

  check("and what it shows is that file's log", after.project, other);
} else {
  check("with one file open the select stays hidden", paneNow.files.length, 1);
}

// ---- nothing here wrote to the project ----------------------------------------------------------

check("Untouched is exactly as the fixture built it", (await held("Untouched")) === untouchedWas);

// Put the fixture's own module back, so a second run starts where the first did.
await write("Ledger", ledgerWas, "change-log.mjs cleanup");
await write("Ticket", (await held("Ticket")).replace("\r\n' back to the agent", ""), "change-log.mjs cleanup");

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
