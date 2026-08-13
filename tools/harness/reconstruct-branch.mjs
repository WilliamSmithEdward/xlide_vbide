/*
 * The reconstruction path - the only path, held at both of the old threshold's sides.
 *
 * Every keystroke's contentChanged carries the ranges alone, and the host REBUILDS its shadow
 * text by applying them; the write-back writes that rebuilt text into the workbook, so drift
 * here corrupts modules. Small modules used to ship their whole text as well and the host
 * preferred that copy, which left the rebuild exercised only above 64,000 characters, where
 * nothing tested it (the audit's C14). The copy was dropped on 2026-08-12 once this suite
 * existed and the measurement favoured the rebuild (88ms against 68ms for the same typing) -
 * so this now pins the one path at a small size and a large one, and pins that the whole-text
 * copy STAYS gone: the tap's contentChanged must carry no fullText on either side.
 *
 * Runs in the -Deep tier, in its own session:
 *   node tools\harness\reconstruct-branch.mjs
 */
import { open, reporter, scratchModule, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();

const { check, done } = reporter();

/** Text the two sides can be compared on: the documented CRLF and trailing-line differences. */
const comparable = (text) => (text ?? "").replace(/\r\n/g, "\n").replace(/\n+$/, "");

/**
 * A module of procedures with a comment first line to type into, sized to a character count.
 * The generated shape mirrors the perf fixture's: identical procedures, so the only variable
 * between the two modules is size.
 */
function moduleOf(targetChars) {
  const lines = ["' seeded", "Option Explicit", ""];
  let n = 0;
  while (lines.join("\r\n").length < targetChars) {
    lines.push(
      `Public Function Fill${n}(ByVal seed As Long) As Long`,
      `    Fill${n} = seed + ${n}`,
      "End Function",
      "");
    n += 1;
  }

  return lines.join("\r\n");
}

/**
 * Seeds one module, types a marker into its first line through the real keyboard path, and
 * proves the three copies agree afterwards. Answers the wall-clock of the type call and
 * whether the change that crossed the wire carried fullText.
 */
async function drive(label, name, targetChars) {
  const marker = `' zzq${process.pid}${label}`;
  const seeded = moduleOf(targetChars);

  await api.component("add", { kind: "module", name, project: project.projectId });
  await api.writeModule(name, seeded, project.projectId);
  await api.pane("open", { module: name, project: project.projectId });
  await waitFor(`${name} to be the shown module`, async () =>
    (await api.state()).shownModule === name);

  // The seed must be ON the surface before typing means anything: the pane open publishes
  // the document, and typing into a model still holding nothing tests nothing.
  await waitFor(`${name}'s seed to reach the surface`, async () =>
    comparable((await api.readModule(name, project.projectId, { live: true })).text) === comparable(seeded));

  // The old threshold's two sides, kept as the size spread: proof at one scale is not proof
  // at the other, and these sizes are the ones the dropped copy used to split between.
  const length = (await api.readModule(name, project.projectId, { live: true })).text.length;
  check(`${label}: the module sits on the intended side of the old 64,000 line`,
    label === "under" ? length < 64_000 : length > 64_000,
    `${length} chars`);

  await api.caret(1, { module: name, column: 1, project: project.projectId });

  const began = Date.now();
  await api.type(marker);
  const typedMs = Date.now() - began;

  // The newest contentChanged for THIS module proves the whole-text copy stays gone. A
  // reintroduced fullText would sit inside the tap row's kept head by field order, even
  // though rows truncate at 2048 characters, so absence here is meaningful.
  const tap = await api.messages(200);
  const changes = (tap.messages ?? [])
    .filter((one) => (one.text ?? "").includes('"contentChanged"')
      && (one.text ?? "").includes(`"${name}"`));
  const newest = changes[changes.length - 1];
  const carried = newest ? (newest.text ?? "").includes('"fullText"') : null;

  check(`${label}: the change crossed the wire as ranges alone, no fullText`,
    carried === false,
    newest ? `fullText ${carried ? "present" : "absent"} in the newest contentChanged` : "no contentChanged seen in the tap");

  // The typed text reaches the workbook through the debounced write-back - through the
  // reconstructed shadow, for the over module, which is the whole point.
  await waitFor(`${name}'s marker to reach the workbook`, async () =>
    ((await api.readModule(name, project.projectId)).text ?? "").includes(marker), { budgetMs: 30000 });

  const live = comparable((await api.readModule(name, project.projectId, { live: true })).text);
  const stored = comparable((await api.readModule(name, project.projectId)).text);
  check(`${label}: the workbook holds the surface's text byte for byte`, stored === live,
    stored === live ? null : `first divergence at char ${[...stored].findIndex((c, i) => c !== live[i])}`);

  const engine = await api.engineSource(name, { text: true });
  check(`${label}: the engine's live copy agrees as well`,
    comparable(engine.engineText) === live,
    engine.engineHolds === false ? "the engine holds no live copy" : `engine ${engine.engineLines} line(s)`);

  return typedMs;
}

const under = scratchModule(api, project.projectId, `Under${process.pid}`);
const over = scratchModule(api, project.projectId, `Over${process.pid}`);

try {
  const underMs = await drive("under", `Under${process.pid}`, 60_000);
  const overMs = await drive("over", `Over${process.pid}`, 68_000);

  // Reported rather than asserted, as a drift watch: both sizes ride the same path now, so
  // these should sit close together - the 88-against-68 gap was the whole-text copy's cost,
  // and it went with the copy.
  console.log(`\n  the same typing, one path: ${underMs}ms at the small size, ${overMs}ms at the large`);
} finally {
  await under.dispose();
  await over.dispose();
  await wait(300);
}

process.exit(done());
