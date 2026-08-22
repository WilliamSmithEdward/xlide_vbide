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

await api.act("changesPane", { expand: false });
await api.act("changesPane", { round: oneHand[0].round, module: second });
await waitFor("the other row", async () =>
  (await api.ui()).changes.showing === `${second}@${oneHand[0].round}`, { budgetMs: 10000 });

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

check("the divider drags", await pull(80), wide + 80);
check("and stops before the rail is too small to use", await pull(-4000), 140);

const ceiling = Number(await api.ask(
  `Math.round(document.getElementById('changes-full-card').clientWidth * 0.4)`));
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

const accepted = await log({ action: "accept" });
check("accepting marks the newest round as reviewed", accepted.acceptedAt, accepted.rounds[0]?.round);
check("and destroys nothing", since(accepted).length, since(routeNow).length);

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
await write(second, (await held(second)).replace("\r\n' back to the agent", ""), "change-log.mjs cleanup");

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
