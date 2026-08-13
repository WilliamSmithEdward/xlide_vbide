/*
 * The 64,000-character threshold, exercised from BOTH sides for the first time.
 *
 * Below it, every keystroke's contentChanged carries the full document and the host prefers
 * that; above it, the message carries only the ranges and the host RECONSTRUCTS its shadow
 * text by applying them. Typical VBA modules sit far below the ceiling, so the reconstruction
 * branch - the one whose drift would make the write-back write a corrupted module into the
 * workbook - ran close to never, and no test anywhere exercised either branch (the audit's
 * C14). This is that test, plus the measurement C14 asked for: the same typing against a
 * module just under and a module just over, timed.
 *
 * The branch taken is PROVEN, not assumed from arithmetic: the message tap holds the page's
 * own contentChanged, and the under module's must carry fullText while the over module's must
 * not. Without that pin, a mis-sized fixture would pass both phases through the fullText path
 * and this suite would guard nothing.
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

  const length = (await api.readModule(name, project.projectId, { live: true })).text.length;
  check(`${label}: the module sits on the intended side of 64,000`,
    label === "under" ? length < 64_000 : length > 64_000,
    `${length} chars`);

  await api.caret(1, { module: name, column: 1, project: project.projectId });

  const began = Date.now();
  await api.type(marker);
  const typedMs = Date.now() - began;

  // The newest contentChanged for THIS module says which branch the host took. The tap
  // truncates rows at 2048 characters, which is fine here BY FIELD ORDER: the message puts
  // type, module, revision and the (one-keystroke) changes ahead of fullText, so the KEY
  // '"fullText"' sits inside the kept head even when its 60KB value is cut.
  const tap = await api.messages(200);
  const changes = (tap.messages ?? [])
    .filter((one) => (one.text ?? "").includes('"contentChanged"')
      && (one.text ?? "").includes(`"${name}"`));
  const newest = changes[changes.length - 1];
  const carried = newest ? (newest.text ?? "").includes('"fullText"') : null;

  check(`${label}: the change crossed the wire ${label === "under" ? "WITH" : "WITHOUT"} fullText`,
    label === "under" ? carried === true : carried === false,
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

  // The measurement C14 called for, reported rather than asserted: machines vary, and the
  // decision this feeds - move the threshold, or drop fullText now that the reconstruction
  // branch is tested - is a person's call made from these numbers.
  console.log(`\n  the same typing: ${underMs}ms under the threshold (fullText), ${overMs}ms over it (reconstruction)`);
} finally {
  await under.dispose();
  await over.dispose();
  await wait(300);
}

process.exit(done());
