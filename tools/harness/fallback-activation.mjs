/*
 * The page's own activations are answers to one tab list, and the host takes them only while
 * that list is current.
 *
 * WHY THIS EXISTS. Two of the page's activations are not the developer's: when a tab list closes
 * the tab on screen the page falls back to the one it showed before, and when a list names
 * nothing active the page says what it is showing instead. Closing a pane makes the host send
 * exactly such a list for a beat before its own choice follows, and the page's answer to it can
 * reach the host after the host has moved on - here, after it has added a class and shown it.
 * Obeyed, it took the surface back to the module before, in every round of the probe that found
 * it. The leak sweep met it as "waiting for the page to show LeakField" whenever the sweep's own
 * navigation happened to land first (2026-09-22). Each list now carries a revision, the page's
 * answers say which list they answer, and the host leaves an answer to a replaced list alone.
 *
 * Two halves. THE RACE replays the sequence - two classes shown, both closed and removed, a third
 * added - and asks what the page shows afterwards, with nothing navigating to hide a stale
 * answer. It lost every round before the fix. THE RULE drives the page's bridge directly: an
 * answer to the current list is taken, an answer to an older one is left alone and the host says
 * so, and the developer's own activations are taken whatever list is current.
 *
 * It brings its own classes and takes them away. The names are fixed and reused every round,
 * because every add spends the session's identifier budget and the gate runs suites after this.
 *
 *   node tools\harness\fallback-activation.mjs
 */

import { open, wait, waitFor, reporter } from "./xlide-api.mjs";

const api = await open({});
const { check, done } = reporter();

const project = await api.project();
const projectId = project.projectId;
const CRLF = "\r\n";

const IFACE = "IXlideFallback";
const IMPL = "XlideFallback";
const TARGET = "XlideFallbackTarget";
const OURS = [IMPL, IFACE, TARGET];

const ROUNDS = 2;

/** The module the page is showing, by the last segment of its model's address. */
const pageShows = async () => String((await api.ui())?.focus?.model ?? "").split("/").pop() ?? "";
const shows = async (module) => (await pageShows()).toLowerCase() === module.toLowerCase();

/** Asks for a module the way a developer's click does, and waits for the page to show it. */
async function onScreen(module) {
  await api.caret(1, { module, project: projectId });
  return waitFor(`the page to show ${module}`, () => shows(module), { budgetMs: 8000 })
    .then(() => true, () => false);
}

/** Pane closed with its edits discarded, then the component removed; never throws. */
async function takeAway(name) {
  await api.pane("close", { module: name, project: projectId, answer: "discard" }).catch(() => null);
  await api.component("remove", { name, project: projectId }).catch(() => null);
}

/** Host log lines matching `text` since `mark` (a log() answer's `next`). */
async function saidSince(mark, text) {
  return (await api.log({ since: mark, match: text, max: 50 })).lines ?? [];
}

try {
  // THE RACE.
  for (let round = 1; round <= ROUNDS; round += 1) {
    for (const [name, text] of [
      [IFACE, ["Option Explicit", "", "Public Sub Go()", "End Sub"]],
      [IMPL, ["Option Explicit", "", `Implements ${IFACE}`]],
    ]) {
      await api.component("remove", { name, project: projectId }).catch(() => null);
      await api.component("add", { kind: "class", name, project: projectId });
      await api.writeModule(name, text.join(CRLF), projectId);
    }

    await onScreen(IFACE);
    await onScreen(IMPL);
    await wait(300);

    const mark = (await api.log({ max: 1 })).next;

    // THE PAGE FALLS BEHIND, ON PURPOSE. The race is the page answering a list after the host
    // has moved on, and whether a small fixture's page lags enough to show it varies from one
    // session to the next: the same probe lost three rounds in three and then held four in
    // four on the same build, its answers landing 80ms late in one session and 20ms in the
    // other. A large project's repaint is what makes the lag in real use, so the page's thread
    // is held for a moment here: the lists below queue behind it, and every answer the page
    // makes to them leaves after the host has already shown the new class.
    await api.eval("setTimeout(() => { const until = Date.now() + 600; while (Date.now() < until) {} }, 0)");

    // The tab on screen closes, and then the one it falls back to - each close a list naming
    // nothing for a beat, each answered by the page - and the host moves on to a new class
    // before the answers can arrive.
    for (const name of [IMPL, IFACE]) {
      await takeAway(name);
    }

    await api.component("remove", { name: TARGET, project: projectId }).catch(() => null);
    await api.component("add", { kind: "class", name: TARGET, project: projectId });
    await api.writeModule(TARGET, ["Option Explicit", "", `Public Held${round} As String`].join(CRLF), projectId);

    // Long enough for any answer in flight to land and be acted on; the stale ones arrived
    // within a tenth of a second of the add.
    await wait(1500);

    const left = await saidSince(mark, "left alone");
    check(`round ${round}: the class the host added is still what the page shows`,
      await shows(TARGET),
      `page shows ${await pageShows()}; ${left.length} stale answer(s) left alone`);

    await takeAway(TARGET);
    await wait(500);
  }

  // THE RULE, through the bridge itself.
  const [first, second] = (project.components ?? [])
    .map((one) => one.name)
    .filter((name) => !OURS.includes(name));

  await api.eval(`window.xlideBridge.activateModule(${JSON.stringify(first)})`);
  check(`the developer's own activation is taken (${first})`,
    await waitFor(`the page to show ${first}`, () => shows(first), { budgetMs: 5000 }).then(() => true, () => false),
    `page shows ${await pageShows()}`);
  await wait(500);

  const current = await api.ask("window.xlideBridge.modulesRevision");
  await api.eval(`window.xlideBridge.activateModule(${JSON.stringify(second)}, undefined, undefined, true)`);
  check(`an answer to the current list (${current}) is taken (${second})`,
    await waitFor(`the page to show ${second}`, () => shows(second), { budgetMs: 5000 }).then(() => true, () => false),
    `page shows ${await pageShows()}`);
  await wait(500);

  const mark = (await api.log({ max: 1 })).next;
  const older = (await api.ask("window.xlideBridge.modulesRevision")) - 1;
  await api.eval(`(() => {
    const bridge = window.xlideBridge;
    const was = bridge.modulesRevision;
    bridge.modulesRevision = ${older};
    bridge.activateModule(${JSON.stringify(first)}, undefined, undefined, true);
    bridge.modulesRevision = was;
  })()`);
  await wait(1500);
  const said = await saidSince(mark, "left alone");
  check(`an answer to an older list (${older}) is left alone`, await shows(second),
    `page shows ${await pageShows()}`);
  check("and the host says so, naming both lists",
    said.some((line) => line.includes(`activate ${first}`) && line.includes(`tab list ${older}`)),
    said[0] ?? "nothing logged");

  await api.eval(`window.xlideBridge.activateModule(${JSON.stringify(first)})`);
  check("the developer's own activation after it is still taken",
    await waitFor(`the page to show ${first}`, () => shows(first), { budgetMs: 5000 }).then(() => true, () => false),
    `page shows ${await pageShows()}`);
} finally {
  for (const name of OURS) {
    await takeAway(name);
  }

  const still = ((await api.project(projectId)).components ?? [])
    .map((one) => one.name)
    .filter((name) => OURS.includes(name));
  check("the classes it brought were taken away", still.length === 0, still.join(", ") || undefined);

  process.exitCode = done();
}
