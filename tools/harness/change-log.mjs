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

import { open, waitFor, comparingReporter } from "./xlide-api.mjs";

const api = await open();
const { check, done } = comparingReporter();

const project = (await api.projects()).projects[0];
console.log(`project: ${project.project}\n`);

// IT WRITES, SO IT CHECKS WHOSE WORKBOOK IT IS IN. Every check below appends to a module and puts
// it back, which is safe in the fixture built for it and is somebody's actual work anywhere else.
// The suite took whatever workbook happened to be open and only failed later, on a missing module
// name - by which point it had already written to two of them (noticed 2026-08-22, with a real
// project open in the editor).
if (!/changefixture/i.test(project.project)) {
  console.log(`FAIL this suite writes to the modules of whatever workbook is open, and this one is`
    + `\n     ${project.project}, which is not the fixture it was built for. Refusing.`
    + `\n\n     powershell tools\\harness\\Start-Excel.ps1 -Fresh -Workbook artifacts\\fixtures\\ChangeFixture.xlsm`);
  process.exit(1);
}

const held = (name) => api.readModule(name, project.projectId).then((one) => one.text);
const write = (name, text, by) => api.writeModule(name, text, project.projectId, { by });
const log = (args = {}) => api.changes({ ...args, project: project.projectId });

// EVERY CHECK BELOW WRITES, and a project that is running or stopped refuses writes - rightly.
// Without this the first write THREW, the throw was unhandled, and node buried the one useful
// sentence under a libuv teardown assertion: the developer was running a FORM in the same
// session while the suite started (2026-08-30), and the report looked like a broken product
// instead of a busy host. A suite that cannot run says so and stops; it does not die mid-word.
const modeNow = (await api.state()).debugMode ?? "design";
if (modeNow !== "design") {
  console.log(`FAIL the project is in ${modeNow} mode, so nothing here can write. Is a form or`
    + " a run standing? Stop it (or POST command?name=reset), then run this again.");
  process.exit(1);
}

// THE SECOND MODULE, BY KIND RATHER THAN BY NAME. The checks below need a component of another
// kind joining a round, which the fixture ships as a class - but RENAMING is one of the things
// this product does, and a suite holding the name its fixture shipped with breaks the first time
// anyone exercises that feature on their own copy. It did: `Ticket` came back `Tickets` and the
// run died on a raw "no module named Ticket" from the door (2026-08-22).
const inside = await api.project(project.project);
const classes = (inside.components ?? []).filter((one) => one.kind === "class");
const second = classes[0]?.name;
if (!second) {
  console.log("FAIL this workbook has no class module, so the second-kind checks cannot run."
    + "\n     Rebuild the fixture: powershell tools\\New-ChangeFixture.ps1");
  process.exit(1);
}

// What every check below is measured against. The fixture's own text, before this run.
const ledgerWas = await held("Ledger");
const secondWas = await held(second);
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
await write(second, `${await held(second)}\r\n' back to the agent`, "claude");

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

// ---- and it says when it has not shown everything ---------------------------------------------
//
// The list stops at `limit`, so a reply of two hundred rounds and a complete history were the
// same answer. That matters to a developer hunting an edit they remember making, and to an agent
// told to review what it changed - and this log's stance everywhere else is to say what it cannot
// show, which is why a round whose text has gone reports `held: false` rather than drawing an
// empty comparison. Asked with a small limit, because proving it needs a truncated view and not
// two hundred rounds.

const whole = await log();
const clipped = await log({ limit: "2" });
check("a limited view returns what was asked for", clipped.rounds.length, 2);
check("and says how many the log actually holds", clipped.total, whole.rounds.length);
check("while an unlimited one agrees with itself", whole.total, whole.rounds.length);

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

// ---- and the pane says when they have been overtaken ------------------------------------------
//
// EVERY COUNT HERE IS A COMPARISON OF TWO WHOLE TEXTS, so the pane works them out when it is
// opened and when the developer asks again, never on the write path. The cost of that rule is
// that the numbers age - and they aged SILENTLY: a reading of +54 sat beside a module the editor
// had already grown to 61 lines, with nothing on screen saying which of the two was current (the
// owner, 2026-08-22: "shouldn't + number align with number of lines?" - they did, with the module
// as it was when the pane last looked).
//
// The rule stands. What changed is that the host taps the pane with a bare integer when it
// records something, and the pane says so. This is the check that the tap is actually sent: the
// page half can be driven by hand, but only a real write proves the host makes the call.

check("nothing is newer, having just read", (await api.ui()).changes.behind, false);

await write("Ledger", `${ledgerWas}\r\n' newer than the reading`, "claude");
const overtaken = await waitFor("the pane to notice", async () => {
  const now = await api.ui();
  return now.changes.behind === true ? now.changes : false;
}, { budgetMs: 15000 });
check("a write while the pane stands marks its counts as overtaken", overtaken.behind, true);

const marker = await api.ask(`JSON.stringify({
  shown: !document.getElementById('changes-newer').hidden,
  words: document.getElementById('changes-newer').textContent.trim(),
  announced: document.getElementById('changes-newer').getAttribute('role'),
  onTheButton: document.getElementById('changes-refresh').classList.contains('changes-refresh-newer'),
})`);
const flag = JSON.parse(typeof marker === "string" ? marker : JSON.stringify(marker));
check("and says so in words, not only a dot, where a reader will be told",
  { shown: flag.shown, words: flag.words, announced: flag.announced, onTheButton: flag.onTheButton },
  { shown: true, words: "newer changes", announced: "status", onTheButton: true });

// AND THE COUNTS DID NOT MOVE ON THEIR OWN. The whole point is that nothing recounted.
check("while the counts themselves stayed exactly as they were read",
  (await api.ui()).changes.rounds.flatMap((round) =>
    round.modules.map((one) => `${one.module}+${one.added}`)),
  drawn.rounds.flatMap((round) => round.modules.map((one) => `${one.module}+${one.added}`)));

await api.act("changesPane", { press: "refresh" });
await waitFor("the read", async () => (await api.ui()).changes.behind === false, { budgetMs: 15000 });
check("reading again clears it", (await api.ui()).changes.behind, false);
check("and the counts have caught up",
  (await api.ui()).changes.rounds[0].modules.some((one) => one.module === "Ledger"), true);

await write("Ledger", `${ledgerWas}\r\n' one\r\n' two\r\n' three`, "claude");
await api.act("changesPane", { press: "refresh" });

// The scope is still ANSWERED, for anything reading the log rather than looking at it - an agent
// deciding whether the absence of a row means a module was untouched or merely unwatched. The pane
// no longer draws it: a caveat under every view is noise to somebody who has read it once.
check("the route still says what the log does not cover",
  routeNow.covers.includes("form designs") && routeNow.covers.includes("written through this editor"));

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

// The button that opens it. A glyph given only horizontal padding takes the button's baseline and
// sits high and left of the middle - which is what the focus ring then draws a rectangle around
// (the owner, 2026-08-22: "glyph is not centered in button").
const centred = await api.ask(`JSON.stringify((() => {
  const off = (button) => {
    if (!button) { return null; }
    const box = button.getBoundingClientRect();
    const mark = button.querySelector('.codicon').getBoundingClientRect();
    return [
      Math.round(((mark.left + mark.right) - (box.left + box.right)) / 2),
      Math.round(((mark.top + mark.bottom) - (box.top + box.bottom)) / 2),
    ];
  };
  return { grow: off(document.getElementById('changes-expand')) };
})())`);
const middle = JSON.parse(typeof centred === "string" ? centred : JSON.stringify(centred));
check(`the expand button centres its glyph (off by ${middle.grow})`,
  Math.abs(middle.grow[0]) <= 1 && Math.abs(middle.grow[1]) <= 1, true);

const grew = await api.act("changesPane", { expand: true });
check("the comparison opens full size", grew.did, true);

// AND A DOUBLE-CLICK ON THE ROW DOES THE SAME. A row is a thing you open, and opening a thing
// properly is a double-click - the gesture the tree already uses for a module (the owner,
// 2026-08-22). Driven twice, because the two paths through it are genuinely different:
//
//   WARM - the row's comparison is already on screen, so the second click is just the expand.
//   COLD - nobody had opened the row, so the request is still in flight when the second click
//          lands. The intent is remembered and honoured by the draw that brings the rows in;
//          asking the host again instead would be a second whole-module comparison for an
//          answer already on its way.

await api.act("changesPane", { expand: false });
const warm = await api.act("changesPane",
  { round: oneHand[0].round, module: "Ledger", gesture: "double" });
check("a double-click on a row already showing opens it full size", warm.did, true);
await waitFor("the card", async () => (await api.ui()).changes.full === true, { budgetMs: 10000 });
check("and it is up", (await api.ui()).changes.full, true);

// Point the strip at some OTHER row first, so Ledger's comparison is genuinely not the one on
// screen. Taken from the pane's own state rather than assumed: `second` was written in a later
// round than `oneHand[0]`, so naming that round asked for a row that does not exist there - the
// pane was right to decline and the check was wrong to wait for it.
await api.act("changesPane", { expand: false });
const elsewhereRow = (await api.ui()).changes.rounds
  .flatMap((round) => round.modules.map((one) => ({ round: round.round, module: one.module })))
  .find((one) => one.module !== "Ledger");

await api.act("changesPane", { round: elsewhereRow.round, module: elsewhereRow.module });
await waitFor("the other row", async () =>
  (await api.ui()).changes.showing === `${elsewhereRow.module}@${elsewhereRow.round}`,
{ budgetMs: 10000 });

// Cold: point the pane away first, so Ledger's rows are genuinely not the ones on screen.
const cold = await api.act("changesPane",
  { round: oneHand[0].round, module: "Ledger", gesture: "double" });
check("a double-click on a row nobody had opened opens it too", cold.did, true);
const arrived = await waitFor("the card to catch up", async () => {
  const now = await api.ui();
  return now.changes.full === true ? now.changes.showing : false;
}, { budgetMs: 15000 });
check("and it shows the row that was double-clicked, not the one that was on screen",
  arrived, `Ledger@${oneHand[0].round}`);

const hint = await api.ask(
  `document.querySelector('#changes-list .changes-entry')?.title ?? ''`);
check("and the row says the gesture is there at all", /Double-click/.test(String(hint)), true);

await api.act("changesPane", { expand: false });
await api.act("changesPane", { round: oneHand[0].round, module: "Ledger" });
await waitFor("the strip back on Ledger", async () =>
  (await api.ui()).changes.showing === `Ledger@${oneHand[0].round}`, { budgetMs: 10000 });
await api.act("changesPane", { expand: true });

const bigger = await api.ask(`JSON.stringify((() => {
  const card = document.querySelector('#changes-full-card');
  const body = document.querySelector('#changes-full-diff');
  const drawn = [...document.querySelectorAll('#changes-full-diff .sync-diff-row')];
  const box = card?.getBoundingClientRect() ?? { top: -1, bottom: -1, left: -1, right: -1 };
  const first = drawn[0]?.getBoundingClientRect() ?? { top: 0 };
  return {
    card: !!card,
    title: document.querySelector('#changes-full-head span')?.textContent ?? '',
    closes: !!document.querySelector('#changes-full-close'),
    rows: drawn.length,
    clipped: getComputedStyle(body.querySelector('.sync-code')).textOverflow,
    onScreen: box.top >= 0 && box.left >= 0
      && Math.ceil(box.bottom) <= window.innerHeight && Math.ceil(box.right) <= window.innerWidth,
    below: Math.round(box.bottom - window.innerHeight),
    fills: Math.round((box.height / window.innerHeight) * 100),
    // Head-room over the first line, which is the top half of the breathing space.
    airTop: Math.round(first.top - body.getBoundingClientRect().top),
  };
})())`);
const full = JSON.parse(typeof bigger === "string" ? bigger : JSON.stringify(bigger));

check("it draws the same rows, with a way out", { card: full.card, closes: full.closes, drew: full.rows > 0 },
  { card: true, closes: true, drew: true });
check("and the code is not clipped, because reading it is the point", full.clipped, "clip");

// WHERE IT LANDS. The shared scaffold's head-room is sized for a small confirm, so a card that
// wants the screen has to fit under what it asked for. This one asked for 92vh under an inherited
// 18vh of padding and hung a tenth of a screen off the bottom edge, on screen, for a release (the
// owner, 2026-08-22: "the popout geometry is off"). Opening was checked; landing was not.
check(`and it lands on the screen (${full.below}px past the bottom)`, full.onScreen, true);

// FULL HEIGHT, whatever it holds, so the comparison is the same shape every time it is opened.
check(`and fills the height (${full.fills}vh)`, full.fills >= 88 && full.fills <= 96, true);

// With the first and last lines off the edges rather than against them.
check(`and the first line is not against the head (${full.airTop}px over it)`, full.airTop >= 6, true);
check("the pane says it is up", (await api.ui()).changes.full, true);

// ---- and it carries the snapshots with it --------------------------------------------------
//
// Opening one comparison full size and having to close it to reach the next one is the dialog
// asking the reader to hold the list in their head. The rail is that list, kept where it can be
// pointed at - and it is a CONTROL, so it is driven here rather than merely counted.

const rail = await api.ask(`JSON.stringify((() => {
  const list = document.getElementById('changes-full-list');
  const opts = [...list.querySelectorAll('[role="option"]')];
  const here = list.querySelector('.changes-full-showing');
  const box = here?.getBoundingClientRect() ?? { height: 0 };
  return {
    listbox: list.getAttribute('role'),
    width: Math.round(list.getBoundingClientRect().width),
    options: opts.length,
    selected: opts.filter((one) => one.getAttribute('aria-selected') === 'true').length,
    // Roving tabindex: ONE stop on the way round the card, arrows moving within it.
    stops: opts.filter((one) => one.tabIndex === 0).length,
    // The current row said other than by fill, and a row big enough to hit.
    bar: here ? getComputedStyle(here).borderLeftColor : '',
    fill: here ? getComputedStyle(here).backgroundColor : '',
    rowHeight: Math.round(box.height),
  };
})())`);
const nav = JSON.parse(typeof rail === "string" ? rail : JSON.stringify(rail));

check("the card carries the snapshot list", { listbox: nav.listbox, offered: nav.options > 1 },
  { listbox: "listbox", offered: true });

// AND IT IS THERE WHATEVER THE LOG HOLDS. The rail used to fold itself away below two snapshots,
// on the rule the file select follows - a list offering the single thing already on screen is
// chrome charging rent. Wrong rule here: the file select sits in a strip that stands either way,
// while this rail and its button are the card's own furniture, so hiding them moved the title,
// moved the code, and took away the control that would bring them back - exactly when a log is
// new, which is the first time anybody opens the card (the owner, 2026-08-22, looking at a card
// with no sidebar and no button). The count is not what decides it.
const furniture = await api.ask(`JSON.stringify({
  rail: !!document.getElementById('changes-full-list')?.offsetParent,
  divider: !!document.getElementById('changes-full-splitter')?.offsetParent,
  button: !!document.getElementById('changes-full-toggle')?.offsetParent,
})`);
check("and the rail, its divider and its button stand whatever the log holds",
  JSON.parse(typeof furniture === "string" ? furniture : JSON.stringify(furniture)),
  { rail: true, divider: true, button: true });

// The head's own two buttons, held to the same square.
const heads = await api.ask(`JSON.stringify((() => {
  const off = (id) => {
    const button = document.getElementById(id);
    const box = button.getBoundingClientRect();
    const mark = button.querySelector('.codicon').getBoundingClientRect();
    return [
      Math.round(((mark.left + mark.right) - (box.left + box.right)) / 2),
      Math.round(((mark.top + mark.bottom) - (box.top + box.bottom)) / 2),
    ];
  };
  return { toggle: off('changes-full-toggle'), close: off('changes-full-close') };
})())`);
const inHead = JSON.parse(typeof heads === "string" ? heads : JSON.stringify(heads));
check(`and the head's buttons centre their glyphs (${JSON.stringify(inHead)})`,
  [...inHead.toggle, ...inHead.close].every((one) => Math.abs(one) <= 1), true);
check(`and it is compact beside the code (${nav.width}px wide)`, nav.width <= 240, true);
check("and exactly one snapshot is the selected one", { selected: nav.selected, stops: nav.stops },
  { selected: 1, stops: 1 });
check("and which one is said other than by colour", nav.bar !== nav.fill && nav.bar !== "", true);
check(`and its rows can be hit (${nav.rowHeight}px)`, nav.rowHeight >= 24, true);
check("and it reports what it offers", (await api.ui()).changes.fullChoices, nav.options);

// Picking another one, from the CARD's rail rather than the pane's list.
const onScreen = (await api.ui()).changes;
const elsewhere = onScreen.rounds
  .flatMap((one) => one.modules.map((each) => ({ round: one.round, module: each.module })))
  .find((one) => !(one.round === oneHand[0].round && one.module === "Ledger"));

const picked = await api.act("changesPane",
  { round: elsewhere.round, module: elsewhere.module, in: "full" });
check("the rail picks another snapshot without closing the card", picked.did, true);

const swapped = await waitFor("the card to swap", async () => {
  const said = await api.ask(`document.getElementById('changes-full-title')?.textContent ?? ''`);
  const title = String(said);
  return title.includes(`round ${elsewhere.round}`) ? title : false;
}, { budgetMs: 15000 });
check("and says which one it is now showing", swapped,
  `${elsewhere.module}, round ${elsewhere.round}`);
check("and the card is still up", (await api.ui()).changes.full, true);

// ONE SOURCE OF TRUTH. The rail and the pane's list are two controls onto one comparison, so a
// pick in the card has to move the pane behind it as well - otherwise closing the card drops the
// reader back onto whatever they were looking at before they started navigating.
check("and the pane behind agrees", (await api.ui()).changes.showing,
  `${elsewhere.module}@${elsewhere.round}`);

// ---- the rail folds away ---------------------------------------------------------------------
//
// One button, and the divider stays behind as the way home.

const wide = (await api.ui()).changes.railWidth;
await api.act("changesPane", { press: "rail" });
const away = await waitFor("the rail to fold", async () => {
  const now = (await api.ui()).changes;
  return now.railUp === false ? now : false;
}, { budgetMs: 10000 });
check("one button puts the snapshots away", away.railUp, false);

const home = await api.ask(`JSON.stringify((() => {
  const button = document.getElementById('changes-full-toggle');
  const head = document.getElementById('changes-full-head');
  return {
    // The way back stays put, in the corner it was pressed in. The DIVIDER goes with the rail:
    // there is nothing left to resize, and it is not what brings it back any more.
    stays: !!button?.offsetParent,
    inHead: head?.firstElementChild === button,
    says: button?.getAttribute('aria-pressed') ?? '',
    divider: !!document.getElementById('changes-full-splitter')?.offsetParent,
  };
})())`);
const folded = JSON.parse(typeof home === "string" ? home : JSON.stringify(home));
check("and the button stays, top left, as the way back",
  { stays: folded.stays, inHead: folded.inHead, says: folded.says, divider: folded.divider },
  { stays: true, inHead: true, says: "false", divider: false });

await api.act("changesPane", { press: "rail" });
const back = await waitFor("the rail to come back", async () => {
  const now = (await api.ui()).changes;
  return now.railUp === true ? now : false;
}, { budgetMs: 10000 });
check("and the same button brings them back, the width it was left",
  { up: back.railUp, width: back.railWidth }, { up: true, width: wide });

// ---- and it is draggable ----------------------------------------------------------------------
//
// How much of the card a list of module names is worth is the reader's call, not this pane's. The
// floor and the ceiling are the pane's, though: a rail at 20px is a control nobody can use, and one
// at 80% of the card has eaten the code it was meant to be a way into.

const pull = async (by) => {
  await api.ask(`(() => {
    const bar = document.getElementById('changes-full-splitter');
    const box = bar.getBoundingClientRect();
    const y = box.top + box.height / 2;
    const at = (type, x) => bar.dispatchEvent(new PointerEvent(type, {
      bubbles: true, cancelable: true, pointerId: 1, isPrimary: true, button: 0, buttons: 1,
      clientX: x, clientY: y,
    }));
    const from = box.left + box.width / 2;
    at('pointerdown', from);
    at('pointermove', from + ${by});
    at('pointerup', from + ${by});
    return true;
  })()`);
  return (await api.ui()).changes.railWidth;
};

// HOW FAR IT CAN DRAG DEPENDS ON THE WINDOW, so the distance comes from the room that exists.
// The ceiling is 40% of the card, and a fixed pull of 80 asserted a rail of 280 inside a card
// that allows 247 - measured on a 644px frame, where this failed as "the divider drags" about a
// divider that had dragged exactly as far as it is permitted to. That is #17's shape: an ambient
// geometry read as a broken feature. The room is checked on its own line so that a window with
// none fails saying so, rather than passing this vacuously on a pull of zero.
const ceiling = Number(await api.ask(
  `Math.round(document.getElementById('changes-full-card').clientWidth * 0.4)`));
const room = Math.min(80, ceiling - wide);
check(`the card is wide enough to have a drag to test (rail ${wide}, ceiling ${ceiling})`,
  room >= 20, true);
check("the divider drags", await pull(room), wide + room);
check("and stops before the rail is too small to use", await pull(-4000), 140);
check("and before it has eaten the code", await pull(4000), ceiling);

const separator = await api.ask(`JSON.stringify((() => {
  const bar = document.getElementById('changes-full-splitter');
  return {
    role: bar.getAttribute('role'),
    orient: bar.getAttribute('aria-orientation'),
    now: bar.getAttribute('aria-valuenow'),
    grip: !!document.getElementById('changes-full-grip')?.offsetParent,
    reachable: bar.tabIndex,
  };
})())`);
const bar = JSON.parse(typeof separator === "string" ? separator : JSON.stringify(separator));
check("and says what it is, with a grip to say it can be pulled",
  { role: bar.role, orient: bar.orient, grip: bar.grip, reachable: bar.reachable >= 0 },
  { role: "separator", orient: "vertical", grip: true, reachable: true });
check("and reports where it stands", Number(bar.now), ceiling);

await pull(-4000);

const shrank = await api.act("changesPane", { expand: false });
check("and it closes", { did: shrank.did, up: (await api.ui()).changes.full }, { did: true, up: false });

// ---- accepting -------------------------------------------------------------------------

// Counted RIGHT BEFORE the accept, not from the reading taken further up: what this check means
// is "accepting removed no rounds", and anything written in between - the overtaken-counts checks
// above write to Ledger - legitimately adds one. Measured against the older reading it failed for
// the writes rather than for the accept, which is the check reporting on the wrong thing.
const beforeAccept = await log();

const accepted = await log({ action: "accept" });
check("accepting marks the newest round as reviewed", accepted.acceptedAt, accepted.rounds[0]?.round);
check("and destroys nothing", since(accepted).length, since(beforeAccept).length);

// WHERE THE MARK SITS. The list runs newest first, so a mark naming the newest reviewed round
// belongs IN FRONT of it: everything below has been seen, everything above has not. Drawn after
// that round instead, the round sat on the unreviewed side - so accepting everything left the mark
// one row down from the top rather than at it.
await api.act("changesPane", { press: "refresh" });
const laidOut = async () => {
  const said = await api.ask(`JSON.stringify(
    [...document.querySelectorAll('#changes-list > *')].map((el) =>
      el.classList.contains('changes-accepted') ? 'accepted' : 'round'))`);
  return JSON.parse(typeof said === "string" ? said : JSON.stringify(said));
};

await waitFor("the pane to show the mark", async () => (await laidOut()).includes("accepted"),
  { budgetMs: 15000 });
check("with everything accepted the mark is at the top", (await laidOut())[0], "accepted");

// And once something new arrives it sits between the new work and the reviewed work.
await write("Ledger", [ledgerWas, "' after the accept"].join("\r\n"), "claude");
await log({ action: "snapshot", label: "since you said yes" });
await api.act("changesPane", { press: "refresh" });
await waitFor("the new round to draw", async () => {
  const now = (await api.ui()).changes;
  return now.rounds[0]?.label === "since you said yes";
}, { budgetMs: 15000 });

check("and a round written since sits above it", (await laidOut()).slice(0, 2), ["round", "accepted"]);

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

// ---- RESTORE: make what the log remembers true again ------------------------------------------
//
// The pane shipped show-only and the owner reversed it (2026-08-30: "full ability to restore from
// any arbitrary snapshot, and revert to last accepted"). What keeps the old rule's substance is
// checked here first and hardest: a restore lands as a ROUND, so the log can always take back its
// own restores, and nothing it does is out of its own reach.

const restoreBase = await held("Ledger");
await log({ action: "snapshot", label: "restore baseline" });
const boundary = (await log()).rounds[0]?.round ?? 0;

await write("Ledger", `${restoreBase}\r\n' restore probe A`, "claude");
await log({ action: "snapshot", label: "A" });
const afterA = await held("Ledger");
const boundaryA = (await log()).rounds[0].round;

const secondBeforeB = await held(second);
await write("Ledger", `${restoreBase}\r\n' restore probe A\r\n' restore probe B`, "claude");
await write(second, `${secondBeforeB}\r\n' restore probe B`, "claude");
await log({ action: "snapshot", label: "B" });

const restoreStarted = Date.now();
const backToA = await log({ action: "restore", round: boundaryA, by: "claude" });
const restoreTook = Date.now() - restoreStarted;
check("restore answers each module's outcome",
  (backToA.outcomes ?? []).map((one) => `${one.module}:${one.did}`).sort().join(","),
  [`Ledger:written`, `${second}:written`].sort().join(","));
check(`and it is quick (${restoreTook}ms for two modules)`, restoreTook < 5000, true);

check("the project is byte-for-byte at the chosen boundary",
  (await held("Ledger")) === afterA && (await held(second)) === secondBeforeB, true);

const restoreRound = (await log()).rounds.find((round) => !round.open);
check("the restore landed as a round of its own",
  `${restoreRound?.label} by ${restoreRound?.by}`, `restore to round ${boundaryA} by claude`);

// THE RESTORE IS RESTORABLE. The boundary to aim at is the round BEFORE the restore's own:
// "after the restore round" is the restored state, which is where we already stand.
await log({ action: "restore", round: backToA.newRound - 1, by: "claude" });
check("restoring the restore away brings the newest text back",
  (await held("Ledger")).includes("restore probe B"), true);

// One module only, back past both probes - the second module must not move.
const secondNow = await held(second);
await log({ action: "restore", round: boundary, module: "Ledger", by: "claude" });
check("a per-module restore reverts that module byte-for-byte",
  (await held("Ledger")) === restoreBase, true);
check("and leaves every other module alone", (await held(second)) === secondNow, true);

// REJECT: everything since the accept mark. Accept here, change, reject, compare.
await log({ action: "accept" });
const acceptedText = await held("Ledger");
await write("Ledger", `${acceptedText}\r\n' past the mark`, "claude");
await log({ action: "snapshot", label: "past the mark" });

await log({ action: "reject", by: "claude" });
check("reject restores to the accept mark", (await held("Ledger")) === acceptedText, true);

// A component round trip: add a CLASS, write it, restore to before - it must LEAVE - then
// restore forward again - it must come BACK, as a class, holding its text.
const phoenix = `Phoenix${process.pid}`;
await log({ action: "snapshot", label: "before the phoenix" });
const beforePhoenix = (await log()).rounds[0].round;

await api.component("add", { name: phoenix, kind: "class", project: project.projectId });
await write(phoenix, ["Option Explicit", "", "Public Sub Rise()", "End Sub"].join("\r\n"), "claude");
await log({ action: "snapshot", label: "the phoenix stands" });
const phoenixText = await held(phoenix);
const phoenixStands = (await log()).rounds[0].round;

await log({ action: "restore", round: beforePhoenix, by: "claude" });
const components = (proj) => (proj.components ?? []).map((one) => one.name.toLowerCase());
check("restoring to before an add removes the module",
  components(await api.project(project.project)).includes(phoenix.toLowerCase()), false);

await log({ action: "restore", round: phoenixStands, by: "claude" });
const reborn = (await api.project(project.project)).components
  ?.find((one) => one.name.toLowerCase() === phoenix.toLowerCase());
check("restoring forward re-adds it", reborn !== undefined, true);
check("as the class it was", reborn?.kind, "class");
check("holding the text it held", (await held(phoenix)) === phoenixText, true);
await api.component("remove", { name: phoenix, project: project.projectId });

// A rename round trip: the restore carries the NAME back as well as the text.
const wanderer = `Wander${process.pid}`;
const returned = `Return${process.pid}`;
await api.component("add", { name: wanderer, kind: "module", project: project.projectId });
await write(wanderer, "Option Explicit\r\n' the wanderer", "claude");
await log({ action: "snapshot", label: "before the wander" });
const beforeWander = (await log()).rounds[0].round;

await api.component("rename", { name: wanderer, newName: returned, project: project.projectId });
await write(returned, "Option Explicit\r\n' the wanderer\r\n' far from home", "claude");

await log({ action: "restore", round: beforeWander, by: "claude" });
const namesNow = components(await api.project(project.project));
check("a restore across a rename carries the name back",
  namesNow.includes(wanderer.toLowerCase()) && !namesNow.includes(returned.toLowerCase()), true);
check("and the text with it", (await held(wanderer)) === "Option Explicit\r\n' the wanderer", true);
await api.component("remove", { name: wanderer, project: project.projectId });

// RENAMED AND THEN REMOVED, restored forward - the owner's live test, 2026-08-30. The plan
// used to emit a rename-back whose source no longer existed, painting "failed - nothing named
// X to rename back" in red over a restore that then succeeded through the re-add. The rename
// step stands down when nothing is standing; the add row speaks for the identity alone.
const ghost = `Ghost${process.pid}`;
const ghostAway = `GhostAway${process.pid}`;
await api.component("add", { name: ghost, kind: "module", project: project.projectId });
await write(ghost, "Option Explicit\r\n' the ghost", "claude");
await log({ action: "snapshot", label: "the ghost stands" });
const ghostStood = (await log()).rounds[0].round;
await api.component("rename", { name: ghost, newName: ghostAway, project: project.projectId });
await api.component("remove", { name: ghostAway, project: project.projectId });
await log({ action: "snapshot", label: "the ghost is gone" });

const seance = await log({ action: "restore", round: ghostStood, by: "claude" });
check(`a renamed-then-removed module restores forward with no failed row `
  + `(${(seance.outcomes ?? []).map((one) => `${one.module}:${one.did}`).join(",")})`,
  (seance.outcomes ?? []).every((one) => one.did !== "failed"), true);
check("and no outcome row is said twice",
  (seance.outcomes ?? []).length
    === new Set((seance.outcomes ?? []).map((one) => JSON.stringify(one))).size, true);
check("and it comes back under its boundary name, text and all",
  (await held(ghost)) === "Option Explicit\r\n' the ghost", true);
await api.component("remove", { name: ghost, project: project.projectId });

// The refusals answer in words, not in silence.
const noSuch = await log({ action: "restore", round: 9999 });
check("a boundary the log does not hold is refused in words",
  /holds no round 9999/.test(noSuch.detail ?? ""), true);

// Without a round the door answers an ERROR, which the client surfaces as a throw - the same
// contract every other malformed request has.
const noArg = await log({ action: "restore" }).then(() => "(answered)").catch((ex) => ex.message);
check("restore without a round asks for one", /needs round=N/.test(noArg), true);

// AND THE PANE'S OWN GESTURE, which is the path a hand takes: the row's Restore control, then
// the confirm its modal raises. Proving only the route would prove the worker and skip the
// buttons, which are the half a developer touches.
await write("Ledger", `${restoreBase}\r\n' pane probe`, "claude");
await log({ action: "snapshot", label: "for the pane" });
const paneListing = await log();
const paneBoundary = paneListing.rounds.filter((one) => !one.open)[1]?.round;

await api.act("changesPane", { press: "refresh" });
await waitFor("the pane to hold the rounds", async () =>
  ((await api.ui()).changes?.rounds ?? []).length > 0, { budgetMs: 15000 }).catch(() => {});

const pressed = await api.act("changesPane", { restore: paneBoundary });
check(`the pane's restore control presses and confirms (${pressed.detail})`, pressed.did, true);
await waitFor("the pane's restore to land", async () =>
  !(await held("Ledger")).includes("' pane probe"), { budgetMs: 20000 }).catch(() => {});
check("and the pane path restores like the route does",
  (await held("Ledger")).includes("' pane probe"), false);

// ---- the SUMMARY: what Reject would take back, worn as a number -------------------------------
//
// A hybrid control beside Accept/Reject (the owner, 2026-08-30): the whole story since the
// accept mark in one glance, the changed modules behind the click, and a module's row the way
// into the editor with the changes highlighted on the lines themselves. Text against text, not
// rounds added up: a module written +5 and hand-reverted -5 has changed nothing.

await log({ action: "accept" });
const summaryBase = await held("Ledger");

// DELETE one line and ADD two unique ones, rather than replacing a line with fixed text: a
// fixed marker already sitting on that line - an earlier probe, a prior run - makes the
// replacement a no-op and the counts collapse (exactly how this check first failed). A
// deletion and two pid-stamped tails read the same whatever the baseline holds.
const baseLines = summaryBase.split(/\r?\n/);
baseLines.splice(3, 1);
baseLines.push(`' tail one ${process.pid}`, `' tail two ${process.pid}`);
await write("Ledger", baseLines.join("\r\n"), "claude");
await write(second, `${await held(second)}\r\n' summary tail ${process.pid}`, "claude");

const summarised = (await log()).sinceAccept;
check("the summary counts text against text, per module and in total",
  JSON.stringify(summarised),
  JSON.stringify({ files: 2, added: 3, removed: 1, entries: [
    { module: "Ledger", added: 2, removed: 1 },
    { module: second, added: 1, removed: 0 },
  ] }));

const sinceRows = await log({ action: "diff", since: "accept", module: "Ledger" });
check("the since-accept diff lines the mark against the LIVE text",
  (sinceRows.rows ?? []).filter((row) => row.kind !== "equal")
    .map((row) => `${row.kind}@${row.rightNumber}`).join(","),
  `removed@null,added@${baseLines.length - 1},added@${baseLines.length}`);

// The pane's own gesture: press the hybrid button, click a module row, and the module opens in
// the EDITOR with the changes painted - which the probe reports as what the paint counted.
await api.act("changesPane", { press: "refresh" });
await waitFor("the pane to hold the summary", async () =>
  (await api.ui()).changes?.sinceAccept?.files === 2, { budgetMs: 15000 });
const summaryPress = await api.act("changesPane", { press: "summary" });
check("the summary button presses", summaryPress.did, true);
const summaryRowClick = await api.act("changesPane", { summaryRow: "Ledger" });
check("and its module row clicks", summaryRowClick.did, true);

const painted = await waitFor("the editor to open Ledger with the paint", async () => {
  const ui = await api.ui();
  return ui.statusModule === "Ledger" && ui.changes?.highlighted?.module === "Ledger"
    ? ui.changes.highlighted : false;
}, { budgetMs: 20000 });
check("the module opens in the editor with its changes painted",
  `+${painted.added} -${painted.removed}`, "+2 -1");

// Reject takes it all back, and the summary stands down - the button with nothing to say leaves.
await log({ action: "reject", by: "claude" });
check("reject clears what the summary was counting", (await log()).sinceAccept, null);
check("and the texts are back at the mark", (await held("Ledger")) === summaryBase, true);

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

// Put the fixture's own modules back, so a second run starts where the first did. The
// second module goes back to its CAPTURED start rather than having one known marker
// stripped: the restore checks append markers of their own, and a strip-one cleanup let
// them compound across runs until a baseline carried a marker and a correct restore
// read as a wrong one (caught on this suite's own second consecutive run, 2026-08-30).
await write("Ledger", ledgerWas, "change-log.mjs cleanup");
await write(second, secondWas, "change-log.mjs cleanup");

process.exit(done());
