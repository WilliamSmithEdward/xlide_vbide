/*
 * Aggressive, seeded, non-deterministic stress against a LIVE editor.
 *
 * Every other suite in here asks a question it already knows the shape of. This one does not: it
 * picks operations at random, fires some of them at once, and watches for the product falling
 * over in a way nobody wrote a check for. What it is hunting is the class of defect a scripted
 * suite structurally cannot find - an ordering nobody thought of, a re-entrancy, a wrapper leaked
 * only on the path where two things arrive together, a modal raised by one operation that the
 * next one waits on for ever.
 *
 * SEEDED, because a fuzzer that cannot be replayed reports ghosts. The seed is printed at the
 * start and on every failure, and `--seed=N` replays that exact run. Randomness is the point;
 * irreproducibility is not.
 *
 * WHAT IT WATCHES, checked between rounds rather than asserted per operation - an individual
 * call failing is often a legitimate refusal, and treating refusals as failures is how a chaos
 * suite becomes noise:
 *
 *   the host still has the same pid     a crash is the loudest possible answer
 *   the door still answers              a hang is the second loudest
 *   COM wrappers are not climbing       the leak that kills Excel, over a long random walk
 *   handles are not climbing            the leak that is not made of wrappers
 *   the log has no unhandled exception  something threw where nobody was catching
 *   the project still reads             the state survived whatever just happened
 *
 * IT DRIVES A COPY, under `artifacts\chaos\`, so `save` is a legal operation to fuzz - which
 * matters, because the saved file is where a class module's predeclared flag is read from - and
 * no fixture is at risk. It creates and destroys modules under a `Chaos_` prefix and never
 * touches the ones it found.
 *
 * AFTER A CRASH RUN, clear the Excel Resiliency key under HKCU and close the recovered workbook
 * before running the next one. Excel brings a crashed workbook back as an untrusted `.xlsb` with
 * MACROS BLOCKED, and the banner it shows reads exactly like a trust problem of the machine's
 * rather than the wreckage of the last run. Leaving it there also means the next launch opens the
 * recovery instead of the workbook asked for. (This is also why the copy lives beside the
 * fixtures rather than in %TEMP%, where it was harder still to tell the two apart.)
 *
 *   copy artifacts\fixtures\LanguageFixture.xlsm artifacts\chaos\ChaosTarget.xlsm
 *   tools\harness\Start-Excel.ps1 -Workbook artifacts\chaos\ChaosTarget.xlsm -Fresh
 *   set XLIDE_CHAOS_PROJECT=ChaosTarget.xlsm
 *
 *   node tools\harness\chaos.mjs                     300 rounds, a fresh seed
 *   node tools\harness\chaos.mjs --rounds=2000       a long walk
 *   node tools\harness\chaos.mjs --seed=1849302113   replay one exactly
 *   node tools\harness\chaos.mjs --quiet             only the failures and the summary
 */

import { open, wait } from "./xlide-api.mjs";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";

/* ---- the dice ------------------------------------------------------------------------------ */

const argOf = (name, fallback) => {
  const found = process.argv.find((one) => one.startsWith(`--${name}=`));
  return found === undefined ? fallback : Number(found.slice(name.length + 3));
};

const SEED = argOf("seed", (Math.floor(Math.random() * 0xffffffff)) >>> 0);
const ROUNDS = argOf("rounds", 300);
const QUIET = process.argv.includes("--quiet");

/** mulberry32: small, fast, and the same sequence for the same seed on every machine. */
function diceFrom(seed) {
  let held = seed >>> 0;
  return () => {
    held = (held + 0x6d2b79f5) >>> 0;
    let t = held;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const dice = diceFrom(SEED);
const upTo = (n) => Math.floor(dice() * n);
const oneOf = (list) => list[upTo(list.length)];
const chance = (p) => dice() < p;

/* ---- what a module might hold ---------------------------------------------------------------
 *
 * Pathological on purpose. The interesting failures are at the edges: text at the size limit,
 * text the host's code page cannot spell, text that does not parse, and text that is empty.
 */

const CRLF = "\r\n";

const bodies = [
  () => ["Option Explicit", "", "Public Sub Small()", "    Debug.Print 1", "End Sub"].join(CRLF),

  // Findings on purpose, so the analyzer has work to do and the problems pane keeps moving.
  () => ["Option Explicit", "", "Public Sub Undeclared()",
    `    missing${upTo(50)} = 1`, "    Dim v As Variant", "    v = alsoMissing", "End Sub"].join(CRLF),

  // Long. Not at VBA's 64,802-line ceiling - the point is a big reparse per keystroke, and the
  // ceiling itself is pinned by RecordCostTests where it costs nothing to run.
  () => ["Option Explicit", "", "Public Sub Long()",
    ...Array.from({ length: 400 + upTo(2600) }, (_, at) => `    Debug.Print "row ${at}"`),
    "End Sub"].join(CRLF),

  // One enormous line, which is a different shape of large from many lines.
  () => ["Option Explicit", "", "Public Sub Wide()",
    `    Debug.Print "${"x".repeat(2000 + upTo(8000))}"`, "End Sub"].join(CRLF),

  // Does not parse, and must not take anything down with it.
  () => ["Option Explicit", "", "Public Sub Broken(", "    If Then Else", "    Dim", "End"].join(CRLF),

  // An unterminated string, which is the classic way to make a lexer run to the end of the file.
  () => ["Option Explicit", "", "Public Sub Unterminated()", '    Debug.Print "never closed', "End Sub"].join(CRLF),

  // Conditional compilation, nested. Which branch is live decides what is analysed at all.
  () => ["Option Explicit", "", "#If VBA7 Then", "#If Win64 Then",
    "Public Sub Wide64()", "End Sub", "#Else", "Public Sub Wide32()", "End Sub", "#End If",
    "#Else", "Public Sub Old()", "End Sub", "#End If"].join(CRLF),

  // Scripts the machine's code page may not hold. The writer is expected to REFUSE some of
  // these rather than mangle them, and a refusal is a pass.
  () => ["Option Explicit", "", "Public Sub Scripts()",
    `    Debug.Print "${oneOf(["ждать", "待つ", "περιμένω", "لانتظار", "chờ đợi", "🎲🎲🎲"])}"`,
    "End Sub"].join(CRLF),

  // Nothing at all. An empty module has no lines, and asking one for line 1 raises.
  () => "",

  // Whitespace only, which is not the same as empty.
  () => CRLF.repeat(1 + upTo(20)),
];

/* ---- the run --------------------------------------------------------------------------------- */

/** Thrown by a move that has checked something and found it WRONG. Ends the walk. */
class Wrong extends Error {}

const api = await open({});
const ledger = [];
const failures = [];
let ops = 0;

const note = (what) => {
  ledger.push(what);
  if (ledger.length > 60) { ledger.shift(); }
};

const project = process.env.XLIDE_CHAOS_PROJECT;
if (!project) {
  console.error("Set XLIDE_CHAOS_PROJECT to the open workbook's file name first.");
  process.exit(2);
}

/*
 * IT WILL NOT RUN AGAINST ANYTHING BUT A CHAOS TARGET, and this is not caution for its own sake.
 *
 * `save` is one of the moves, deliberately - the saved file is where a class module's predeclared
 * flag is read from, so a walk that never saved would not exercise it. That makes this the one
 * suite here that writes the workbook it is pointed at. Pointed at a fixture, it saves its own
 * `Chaos_` modules and the immediate window's scratch INTO it, and the next suite to use that
 * fixture opens a project that no longer compiles.
 *
 * Which is exactly what happened on 2026-08-23: LanguageFixture.xlsm came back with eight Chaos
 * modules and an XlideImmediateScratch baked in, and the next launch raised a Compile error
 * dialog before anything had run. The convention is enforced rather than remembered.
 */
if (!/chaos/i.test(project)) {
  console.error(`Refusing to run against "${project}".`);
  console.error("This walk SAVES the workbook it drives, so it only runs against a copy whose");
  console.error("name says so. Make one and point it there:");
  console.error("");
  console.error("  copy artifacts\\fixtures\\LanguageFixture.xlsm artifacts\\chaos\\ChaosTarget.xlsm");
  console.error("  tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\chaos\\ChaosTarget.xlsm -Fresh");
  console.error("  set XLIDE_CHAOS_PROJECT=ChaosTarget.xlsm");
  process.exit(2);
}

const startPid = api.pid;
const before = await api.stats();
const known = (await api.project(project)).components.map((one) => one.name);
const chaosNames = () => Array.from({ length: 6 }, (_, at) => `Chaos_${at}`);

/**
 * A module belonging to ONE operation, so a check that compares what it wrote against what came
 * back is comparing its own text and nobody else's.
 *
 * The shared pool above is deliberately contended - that is what most of the walk is for - but a
 * move that ASSERTS needs a subject no other move is writing to, or it reports the product wrong
 * for a race of the harness's own making. Named per operation and left behind; the walk removes
 * its own modules at the end.
 */
/**
 * Clears a stopped project before a move that means to write, the way the refusal message tells
 * a caller to.
 *
 * NOT POLITENESS - COVERAGE. Between rounds was not often enough: an evaluation stops the project
 * again within the round, and the assertion moves were being turned away before they compared
 * anything. Measured over two walks of 640 and 679 operations, the cross-workbook check landed 2
 * times out of 32 and then 0 out of 52. A check that cannot run is not a check, so the moves that
 * need design mode ask for it first, and the count of how often stays in the summary as #7's size.
 */
const readyToWrite = async () => {
  if ((await api.state().catch(() => ({}))).debugMode !== "break") { return; }
  unwedged += 1;

  // THE SIGNATURE OF ISSUE #9. A reset that answers "currently disabled" while the editor says it
  // is stopped is the break nothing can leave: every debug command refused, no dialog standing,
  // no form showing, and only restarting Excel clears it. Counted apart from the ordinary wedge,
  // because the two look identical from here until the reset answers.
  const said = await api.command("reset").catch(() => ({ ran: false, detail: "unreachable" }));
  if (said.ran === false && String(said.detail).includes("disabled")) {
    stuck += 1;
  }

  await wait(250);
};

const claimed = [];
const mine = async (what) => {
  await readyToWrite();

  // THE SEED IS IN THE NAME, because the operation counter restarts every run and Excel does not.
  // Two walks against one instance both reached `Chaos_rt104`, and the second was refused as a
  // name already taken - which the summary reported eleven times as though the product had done
  // something wrong. A run's modules are its own.
  const name = `Chaos_${what}${SEED.toString(36)}_${ops}`;
  await api.component("add", { kind: "module", name, project });
  claimed.push(name);
  return name;
};

/** Modules made in a NAMED project, so the twin's are given back too. */
const claimedIn = [];

/**
 * A SECOND workbook, when one is open, so the walk can ask the questions a single file cannot.
 *
 * Found rather than configured: whatever else the session holds that is not the target. Null when
 * there is only one, and the moves that need two say so and stand down rather than failing.
 */
const twin = ((await api.projects().catch(() => ({ projects: [] }))).projects ?? [])
  .map((one) => one.display ?? one.project ?? one)
  .find((one) => typeof one === "string" && one !== project) ?? null;

console.log(`chaos: seed ${SEED}, ${ROUNDS} rounds, pid ${startPid}, project ${project}`);
console.log(`       found ${known.length} component(s): ${known.join(", ")}`);
console.log(`       wrappers ${before.comWrappersLive}, handles ${before.handleCount}`);
console.log(twin === null
  ? "       one workbook open, so the cross-workbook checks stand down\n"
  : `       twin workbook: ${twin}\n`);

// The dialog guard ON, so a modal raised by one operation does not park every later one behind
// it. A fuzzer that deadlocks on the first message box tests one operation and then nothing.
// `--noguard` leaves it off, which is how the product ships. The difference between the two
// runs is what scopes a failure: one that only happens with the guard on belongs to agents
// driving the door, and one that happens either way belongs to everybody.
if (!process.argv.includes("--noguard")) {
  await api.guard(true).catch(() => {});
}

/** Every operation the walk can pick, and how often. A refusal is not a failure. */
const moves = [
  [8, "write", async () => {
    const name = oneOf(chaosNames());
    await api.component("add", { kind: oneOf(["module", "class", "module"]), name, project }).catch(() => {});
    const text = oneOf(bodies)();
    await api.writeModule(name, text, project);
    return `${name} <- ${text.length} chars`;
  }],

  [6, "roundtrip", async () => {
    // WRITE, THEN READ BACK, AND COMPARE. Everything else here watches for the product falling
    // over; this watches for it staying up and being WRONG, which is the worse of the two and
    // the one nothing in this walk could previously see. A module that comes back different from
    // what went in is a developer's code quietly altered.
    // ITS OWN MODULE, not one from the shared pool. Two of these landing in the same burst wrote
    // the same name and each read back the other's text, and the walk called the product WRONG
    // for something no contract covers: two writers to one module have no defined outcome. It
    // still races everything else - pane opens, analysis, saves, the reseed - which is the part
    // worth racing.
    const name = await mine("rt");

    // Distinctive, so a mix-up with another module's text is obvious rather than plausible.
    const stamp = `${SEED}_${ops}`;
    const sent = ["Option Explicit", "", `' roundtrip ${stamp}`, `Public Sub R${upTo(1000)}()`,
      `    Debug.Print "${stamp}"`, "End Sub"].join(CRLF);

    // NOT CAUGHT. A write this walk could not make is a refusal, and it has to be COUNTED as
    // one - swallowing it and returning a tidy string made every roundtrip report `ok` while
    // comparing nothing at all. Measured: with the comparison deliberately broken, the move
    // still said "7 ok, 0 refused", because the project happened to be stopped and every write
    // was being turned away before anything was checked.
    await api.writeModule(name, sent, project);

    const back = (await api.readModule(name, project)).text ?? "";

    // Line endings are the editor's to decide - it stores what it stores - so the comparison is
    // on the lines, not the bytes between them. A trailing newline is likewise not a difference.
    const lines = (text) => text.replace(/\r\n/g, "\n").replace(/\n+$/, "").split("\n");
    const [a, b] = [lines(sent), lines(back)];
    if (a.length !== b.length || a.some((line, at) => line !== b[at])) {
      throw new Wrong(`ROUNDTRIP MISMATCH in ${name}: sent ${a.length} line(s), got back `
        + `${b.length}. First difference at ${a.findIndex((line, at) => line !== b[at])}: `
        + `${JSON.stringify(a.find((line, at) => line !== b[at]))} vs `
        + `${JSON.stringify(b[a.findIndex((line, at) => line !== b[at])])}`);
    }

    return `${name} ${a.length} line(s) intact`;
  }],

  [5, "twin", async () => {
    /*
     * THE SAME MODULE NAME IN TWO WORKBOOKS, which is where this product's nastiest defects have
     * lived. A module name is not an identity: two open files both hold a Sheet1, and everything
     * keyed by name alone - the engine's live copy, the write's target, the findings, the
     * breakpoints - has at some point picked whichever file enumerated first.
     *
     * So both get a module of the same name and DIFFERENT text, and each must read back its own.
     * Nothing else in the harness can ask this: every other suite drives one file.
     */
    if (twin === null) { return "no second workbook open"; }

    await readyToWrite();

    // Seeded like the others, so two walks against one Excel do not collide on a name.
    const name = `Chaos_tw${SEED.toString(36)}_${ops}`;
    for (const where of [project, twin]) {
      await api.component("add", { kind: "module", name, project: where });
      claimedIn.push([where, name]);
    }

    const textFor = (where) => ["Option Explicit", "", `' this module belongs to ${where}`,
      `Public Sub Belongs${upTo(1000)}()`, "End Sub"].join(CRLF);
    const sent = new Map([[project, textFor(project)], [twin, textFor(twin)]]);

    // Written one after the other on purpose: what is being asked is whether the SECOND write
    // lands on the second file, not whether two concurrent writes are ordered.
    for (const where of [project, twin]) {
      await api.writeModule(name, sent.get(where), where);
    }

    for (const where of [project, twin]) {
      const back = (await api.readModule(name, where)).text ?? "";
      if (!back.includes(`belongs to ${where}`)) {
        const other = where === project ? twin : project;
        throw new Wrong(`CROSSED WORKBOOKS: ${where}'s ${name} came back holding `
          + (back.includes(`belongs to ${other}`)
            ? `${other}'s text. A module name was taken for an identity.`
            : `neither file's text: ${JSON.stringify(back.slice(0, 80))}`));
      }
    }

    return `${name} kept apart in both files`;
  }],

  [4, "coherence", async () => {
    // DOES THE SURFACE AGREE WITH THE WORKBOOK? `live=1` is what the editor holds and the bare
    // read is what the module holds, and after a write-back the two must be the same text. When
    // they are not, the analyzer is diagnosing something nobody can see: it is what left the
    // problems of a discarded edit alive after a close and a reopen, and what made a workbook
    // hold 42 lines while the editor showed an empty document and refused every breakpoint on it.
    // Its own module, for the same reason the roundtrip has one.
    const name = await mine("co");
    await api.pane("open", { module: name, project });

    const stamp = `${SEED}_${ops}`;
    await api.writeModule(name, ["Option Explicit", "", `' coherence ${stamp}`,
      `Public Sub C${upTo(1000)}()`, "End Sub"].join(CRLF), project);

    const stored = (await api.readModule(name, project)).text ?? "";
    const live = (await api.readModule(name, project, { live: true })).text ?? "";

    const flat = (text) => text.replace(/\r\n/g, "\n").replace(/\n+$/, "");
    if (flat(stored) !== flat(live)) {
      throw new Wrong(`SURFACE AND WORKBOOK DISAGREE about ${name}: the module holds `
        + `${flat(stored).length} chars, the editor shows ${flat(live).length}. `
        + `stored ${JSON.stringify(flat(stored).slice(0, 90))} vs `
        + `live ${JSON.stringify(flat(live).slice(0, 90))}`);
    }

    return `${name} agrees, ${flat(stored).length} chars`;
  }],

  [5, "read", async () => {
    const name = oneOf([...known, ...chaosNames()]);
    const answer = await api.readModule(name, project, { live: chance(0.5) });
    return `${name} -> ${(answer.text ?? "").length} chars`;
  }],

  [4, "pane", async () => {
    const name = oneOf([...known, ...chaosNames()]);
    const action = oneOf(["open", "open", "close"]);
    await api.pane(action, { module: name, project, answer: "discard" });
    return `${action} ${name}`;
  }],

  [4, "caret", async () => {
    const name = oneOf(known);
    const line = 1 + upTo(60);
    await api.caret(line, { module: name, project, column: 1 + upTo(40) });
    return `${name}:${line}`;
  }],

  [3, "component", async () => {
    const name = oneOf(chaosNames());
    const action = oneOf(["add", "remove", "rename"]);
    if (action === "rename") {
      await api.component("rename", { name, newName: `${name}_r${upTo(4)}`, project });
    } else {
      await api.component(action, { kind: oneOf(["module", "class", "form"]), name, project });
    }
    return `${action} ${name}`;
  }],

  [3, "problems", async () => `${((await api.problems()).findings ?? []).length} finding(s)`],

  [2, "project", async () => `${(await api.project(project)).components.length} component(s)`],

  [2, "outline", async () => {
    const name = oneOf([...known, ...chaosNames()]);
    return `${name} -> ${((await api.outline(name, project)).procedures ?? []).length}`;
  }],

  [2, "changes", async () => {
    const action = oneOf([undefined, "text", "diff", "snapshot"]);
    const answer = await api.changes({ action, project, label: action === "snapshot" ? `chaos${ops}` : undefined });
    return `${action ?? "list"} -> ${Object.keys(answer).join(",")}`;
  }],

  [2, "state", async () => `${(await api.state()).debugMode}`],

  [1, "compile", async () => `${JSON.stringify(await api.compile({ waitMs: 8000 })).slice(0, 60)}`],

  [1, "save", async () => {
    // The saved file is where a class module's predeclared flag is read from, so saving is a
    // state change the analyzer can see - and one nothing else in the harness exercises at random.
    await api.command("save");
    return "saved";
  }],

  [1, "sync", async () => `plan ${((await api.syncPlan("export", { project })).items ?? []).length} row(s)`],

  [1, "perf", async () => `${Object.keys(await api.perf({ reset: chance(0.2) })).length} key(s)`],

  [1, "undoRename", async () => `${JSON.stringify(await api.undoRename()).slice(0, 60)}`],

  [1, "tests", async () => `${JSON.stringify(await api.tests({ action: "list" })).slice(0, 60)}`],

  [1, "drain", async () => {
    await api.drainFinalizers();
    return "drained";
  }],

  /* ---- and the ones most likely to find something ------------------------------------------
   *
   * Everything above is one subsystem at a time. These reach across two, which is where an
   * ordering nobody thought of actually lives.
   */

  [4, "tab", async () => {
    const action = oneOf(["activate", "cycleTab", "split", "focusEditor", "changesPane", "unfoldModule"]);
    const answer = await api.act(action, { name: oneOf([...known, ...chaosNames()]), project });
    return `${action} -> ${answer.did}`;
  }],

  [3, "search", async () => {
    if (chance(0.4)) {
      await api.act("search", { close: 1 });
      return "closed";
    }

    const query = oneOf(["Debug", "Sub", "x".repeat(200), "", "[", "\\", "Ждать", "End"]);
    const answer = await api.act("search", {
      query, scope: oneOf(["module", "project", undefined]),
      matchCase: chance(0.5) ? 1 : undefined, wholeWord: chance(0.5) ? 1 : undefined,
    });
    return `${JSON.stringify(query).slice(0, 24)} -> ${answer.did}`;
  }],

  [3, "editorAction", async () => {
    // Whatever the surface will take. Formatting rewrites every line, which is the operation
    // that historically moved every finding in a module underneath itself.
    const id = oneOf(["xlide.format", "xlide.undoRename", "xlide.toggleComment", "xlide.indent"]);
    const answer = await api.act("editorAction", { id });
    return `${id} -> ${answer.did}`;
  }],

  [3, "designer", async () => {
    const form = oneOf(known.filter((one) => /form/i.test(one)));
    if (form === undefined) { return "no form here"; }

    const action = oneOf(["add", "remove", "set"]);
    const name = `Chaos${upTo(4)}`;
    if (action === "add") {
      await api.designerEdit("add", { module: form, project, type: oneOf(["label", "textBox", "commandButton"]),
        name, left: upTo(300), top: upTo(200), width: 10 + upTo(120), height: 10 + upTo(40) });
    } else if (action === "remove") {
      await api.designerEdit("remove", { module: form, project, name });
    } else {
      await api.designerEdit("set", { module: form, project, name,
        property: oneOf(["Caption", "BackColor", "Width", "Visible"]),
        value: oneOf(["hello", "255", "-1", "0", "not a number", ""]) });
    }
    return `${action} ${name} on ${form}`;
  }],

  [2, "immediate", async () => {
    const answer = await api.immediate(oneOf(["?1+1", "?Now", "?ThisWorkbook.Name", "Debug.Print 1/0", "?"]));
    return `${JSON.stringify(answer).slice(0, 60)}`;
  }],

  [2, "breakpoint", async () => {
    const name = oneOf([...known, ...chaosNames()]);
    await api.breakpoint(name, 1 + upTo(30), { project, state: oneOf(["on", "off", undefined]) });
    return `${name}`;
  }],

  [2, "renameModule", async () => {
    const from = oneOf(chaosNames());
    const to = `${from}_n${upTo(3)}`;
    const answer = await api.act("renameModule", { name: from, newName: to, project });
    return `${from} -> ${to}: ${answer.did}`;
  }],

  [1, "layout", async () => {
    if (chance(0.3)) { await api.resetLayout(); return "reset"; }
    return `${Object.keys(await api.layout()).length} key(s)`;
  }],

  [1, "reload", async () => {
    // The heaviest thing a caller can ask for: the whole page torn down and rebuilt, while
    // everything else here is still arriving.
    await api.reload({ waitMs: 25000 });
    return "page reloaded";
  }],
];

/*
 * `--only=pane,immediate` narrows the walk to those moves.
 *
 * This is how a find becomes a diagnosis. The whole point of the full walk is that it reaches
 * states nobody designed; the point of this is that once it has broken something, the set of
 * moves involved can be halved until what is left is small enough to read.
 */
const only = process.argv.find((one) => one.startsWith("--only="));
const chosen = only === undefined ? moves : moves.filter(([, name]) =>
  only.slice(7).split(",").includes(name));

if (chosen.length === 0) {
  console.error(`--only names no move. Available: ${moves.map(([, name]) => name).join(", ")}`);
  process.exit(2);
}

if (only !== undefined) {
  console.log(`narrowed to: ${chosen.map(([, name]) => name).join(", ")}
`);
}

const total = chosen.reduce((sum, [weight]) => sum + weight, 0);
const pick = () => {
  let at = upTo(total);
  for (const move of chosen) {
    at -= move[0];
    if (at < 0) { return move; }
  }
  return chosen[0];
};

/**
 * One operation. A refusal is recorded and forgiven; the invariants decide what is a failure.
 *
 * The TALLY matters as much as the verdict. A fuzzer whose operations are nearly all being
 * refused is bouncing off the front door and testing nothing, and the only way to know that is
 * to count - so the summary prints how many of each landed and how many were turned away.
 */
const tally = new Map();
const count = (name, outcome) => {
  const held = tally.get(name) ?? { ok: 0, refused: 0, reasons: new Map() };
  held[outcome] += 1;
  tally.set(name, held);
  return held;
};

const fire = async () => {
  const [, name, run] = pick();
  ops += 1;
  try {
    const detail = await run();
    count(name, "ok");
    note(`ok   ${name}  ${detail}`);
  } catch (err) {
    // A REFUSAL IS NOT A DEFECT, but some of what a move can discover is. A move that has
    // CHECKED something and found it wrong throws Wrong, and that ends the walk rather than
    // being counted among the ordinary noise of asking for things that are not there.
    if (err instanceof Wrong) {
      note(`WRONG ${name}  ${err.message}`);
      fail(`a move found the product wrong: ${name}`, err.message);
      return;
    }

    // The WHOLE message, keyed by a normalised form. Slicing it to a colon-tail read tidily
    // and threw away which route said it and what it was about, which is most of the value.
    const full = String(err.message);
    const why = full.replace(/[0-9]+/g, "#").slice(0, 160);
    const held = count(name, "refused");
    const seen = held.reasons.get(why) ?? { n: 0, full };
    seen.n += 1;
    held.reasons.set(why, seen);
    note(`REFUSED ${name}  ${String(err.message).slice(0, 120)}`);
  }
};

/**
 * The host's own last words, read off the log FILE rather than through the door.
 *
 * When the interesting failure is the host dying, the door dies with it - so the one account of
 * what happened is the file the shim was writing at the time, and fetching it by hand every time
 * is how a crash goes uninvestigated.
 */
const lastWords = (pid, lines = 22) => {
  try {
    const folder = join(process.env.LOCALAPPDATA, "xlide_vbide", "logs");
    const name = readdirSync(folder).filter((one) => one.endsWith(`-${pid}.log`)).pop();
    if (name === undefined) { return ["(no log for this pid)"]; }
    return readFileSync(join(folder, name), "utf8")
      .split(/\r?\n/).filter((one) => one.length > 0).slice(-lines)
      .map((one) => one.slice(0, 150));
  } catch (err) {
    return [`(the log would not be read: ${err.message})`];
  }
};

const fail = (what, detail) => {
  failures.push({ round: ops, what, detail });
  console.log(`\nBROKE: ${what}\n  ${detail}`);
  console.log(`  seed ${SEED}, after ${ops} operation(s). The last of them:`);
  for (const line of ledger.slice(-25)) { console.log(`    ${line}`); }
  console.log(`\n  and what the host itself was doing (pid ${startPid}):`);
  for (const line of lastWords(startPid)) { console.log(`    ${line}`); }
  console.log("");
};

/** Whether a process id still names a running process. Signal 0 tests, it does not kill. */
const alive = (pid) => {
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err.code === "EPERM";
  }
};

let unwedged = 0;
let stalls = 0;
let stuck = 0;
let peakWrappers = before.comWrappersLive;
let peakHandles = before.handleCount;

for (let round = 1; round <= ROUNDS && failures.length === 0; round += 1) {
  // Sometimes one at a time, sometimes a burst arriving together. The burst is the interesting
  // one: the host thread is a single lane and everything here has to queue for it.
  const burst = chance(0.55) ? 2 + upTo(7) : 1;
  await Promise.allSettled(Array.from({ length: burst }, () => fire()));

  // And usually no pause at all, so operations land inside each other's debounce windows and
  // inside each other's host-thread crossings.
  if (chance(0.25)) { await wait(upTo(30)); }

  /*
   * UNWEDGE, the way the product tells a caller to.
   *
   * An evaluation against a project that will not compile leaves it stopped, and while it is
   * stopped every write and every component change is refused (issue #7). Left alone, the walk
   * wedges in the first few rounds and spends the rest of its life being turned away: measured
   * at 2 roundtrips landing and 35 refused, which is a fuzzer that has stopped fuzzing.
   *
   * So it does what the refusal message says to do. That is not the harness working around the
   * product - it is the harness behaving like the competent caller the message is addressed to,
   * and the COUNT below is worth as much as the coverage it buys: how often a random walk has
   * to reach for the escape hatch is the size of #7, measured rather than argued.
   */
  if ((await api.state().catch(() => ({}))).debugMode === "break") {
    unwedged += 1;
    await api.command("reset").catch(() => {});
  }

  if (round % 10 !== 0) { continue; }

  /* ---- the invariants ---------------------------------------------------------------------- */

  // THE HOST ITSELF, asked of the operating system rather than of the door. `api.pid` is read
  // from the discovery file when the client connects and never changes afterwards, so comparing
  // it to itself was a crash check that could not fail - which it duly did not, through a run
  // where Excel died and came back under a new pid with "Recovered" in its title.
  if (!alive(startPid)) {
    fail("the host is gone", `pid ${startPid} is no longer a process. Excel died under this walk.`);
    break;
  }

  if (!(await api.waitUntilResponsive({ timeout: 30000 }))) {
    fail("the door stopped answering",
      `no reply for 30s after ${ops} operations, though pid ${startPid} is still alive - `
      + "so the host is up and its add-in is not answering");
    break;
  }

  const now = await api.stats().catch(() => null);
  if (now === null) {
    fail("stats would not answer", "the door replied but this route did not");
    break;
  }

  peakWrappers = Math.max(peakWrappers, now.comWrappersLive);
  peakHandles = Math.max(peakHandles, now.handleCount);

  // A ceiling rather than a delta. The walk legitimately holds panes and modules open, so the
  // count is expected to move; what must not happen is unbounded growth.
  if (now.comWrappersLive > before.comWrappersLive + 2000) {
    fail("COM wrappers are climbing",
      `${before.comWrappersLive} at rest, ${now.comWrappersLive} now`);
    break;
  }

  if (now.handleCount > before.handleCount + 4000) {
    fail("handles are climbing", `${before.handleCount} at rest, ${now.handleCount} now`);
    break;
  }

  /*
   * A STALL THAT CLEARS IS NOT A HANG, and the difference decides whether the walk stops.
   *
   * The host thread is a single lane, and every analysis pass reads every project along it, so
   * under a burst a request can miss its three-second crossing and answer "the host thread did
   * not answer in time" (issue #8). That is worth counting and not worth stopping for: treating
   * it as a hang ends the walk at the first busy moment and hides everything it would have found
   * afterwards.
   *
   * So it is asked TWICE, with a pause. Answering the second time means the thread was busy;
   * still refusing means it is gone, which is the thing this check is actually for.
   */
  let readable = await api.project(project).catch((err) => String(err.message));
  if (typeof readable === "string") {
    stalls += 1;
    await wait(3000);
    readable = await api.project(project).catch((err) => String(err.message));
  }

  if (typeof readable === "string") {
    fail("the project stopped reading, twice, three seconds apart", readable);
    break;
  }

  if (!QUIET) {
    console.log(`  round ${String(round).padStart(4)}  ops ${String(ops).padStart(5)}`
      + `  wrappers ${now.comWrappersLive}  handles ${now.handleCount}`
      + `  components ${readable.components.length}`);
  }
}

/* ---- and what the host thought of it all ----------------------------------------------------- */

// The per-operation modules go back, so a long walk does not leave hundreds of them behind for
// the next one to enumerate. Best effort: a walk that ended because the host died has nothing to
// tidy, and saying so twice helps nobody.
for (const name of claimed) {
  await api.component("remove", { name, project }).catch(() => {});
}
for (const [where, name] of claimedIn) {
  await api.component("remove", { name, project: where }).catch(() => {});
}

const survived = failures.length === 0;

let logTail = "";
try {
  logTail = (await api.log({ max: 4000 })).text ?? "";
} catch { /* the door may be gone, which the invariants have already reported */ }

const angry = logTail.split(/\r?\n/).filter((line) =>
  /unhandled|System\.\w+Exception|0xc0000005|stack overflow/i.test(line));

console.log("\n---------------------------------------------------------------");
console.log(`seed        ${SEED}`);
console.log(`operations  ${ops}`);
console.log(`host        pid ${startPid} ${alive(startPid) ? "still running" : "IS GONE - Excel died under this walk"}`);
console.log(`wrappers    ${before.comWrappersLive} at rest, peak ${peakWrappers}`);
console.log(`handles     ${before.handleCount} at rest, peak ${peakHandles}`);
console.log(`log         ${angry.length} line(s) naming an unhandled fault`);
console.log(`unwedged    ${unwedged} time(s) - found stopped and reset, which is issue #7's size`);
console.log(`stalled     ${stalls} time(s) - a crossing missed its budget and caught up (#8)`);
console.log(`STUCK       ${stuck} time(s) - stopped with Reset itself disabled, which is #9`);
console.log("");
console.log("what landed, and what was turned away:");
for (const [name, held] of [...tally.entries()].sort((a, b) => (b[1].ok + b[1].refused) - (a[1].ok + a[1].refused))) {
  const top = [...held.reasons.entries()].sort((a, b) => b[1].n - a[1].n)[0];
  console.log(`  ${name.padEnd(14)} ${String(held.ok).padStart(4)} ok  ${String(held.refused).padStart(4)} refused`);
  if (top !== undefined) { console.log(`      x${top[1].n} ${top[1].full.slice(0, 200)}`); }
}
for (const line of angry.slice(0, 10)) { console.log(`              ${line.slice(0, 160)}`); }

console.log(`\nRESULT: ${survived && angry.length === 0 ? "SURVIVED" : "BROKE"} - ${ops} operations at seed ${SEED}`);
process.exit(survived && angry.length === 0 ? 0 : 1);
