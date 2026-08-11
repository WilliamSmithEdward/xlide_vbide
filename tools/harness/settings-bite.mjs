/*
 * Does every setting actually change what the editor does?
 *
 * WHY THIS EXISTS. On 2026-08-09 the export mode dropdown was found to have never reached the plan
 * at all: the control rebuilt itself and discarded the selection four lines before it was read, so
 * choosing "Delete them" did nothing and the dropdown snapped back on its own. Nothing caught it
 * because the api passes the mode as an argument and never touches the control. A setting that
 * saves correctly and changes nothing is a defect this product has already had.
 *
 * Of seven settings, two were named in any harness file before this. Three of the others govern
 * Enter and could not be driven at all until `act("press")` arrived the same day.
 *
 * Each is set BOTH WAYS and the behaviour observed. The stored value is not the subject: a setting
 * that round-trips through the store and reaches nothing is exactly what is being looked for.
 *
 *   node tools\harness\settings-bite.mjs
 */
import { open, waitFor, waitUntilStable } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const name = `Bite${process.pid}`;

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
  if (ok) { passed += 1; console.log(`ok   ${what}`); }
  else { failures.push(what); console.log(`FAIL ${what}${detail ? `\n     ${detail}` : ""}`); }
};

const live = async () =>
  ((await api.readModule(name, project.projectId, { live: true })).text ?? "")
    .split("\n").map((one) => one.replace("\r", ""));

/**
 * Seeds the module, puts the caret at the end of a line, and presses Enter.
 *
 * THE SEED WAIT IS WHAT KEEPS THE REST HONEST. This is called six times against the same module,
 * so at the moment the write goes out the module still holds the PREVIOUS run's result - which
 * already has the extra line Enter is about to add. Wait only for "a line appeared" and it is
 * satisfied by the leftovers before the seed has even landed, and the suite compares two settings
 * by reading the same stale text twice. So the seed is waited for as a whole text first, and only
 * then does the line count mean anything.
 */
async function enterOn(source, line) {
  const seed = source.join("\n").trim();

  await api.writeModule(name, source.join("\r\n"), project.projectId);
  await waitFor("the seed to be the module's whole text", async () =>
    (await live()).join("\n").trim() === seed);

  await api.caret(line, { module: name, project: project.projectId, column: source[line - 1].length + 1 });
  await waitFor("the caret to reach the end of the line Enter is pressed on", async () => {
    const focus = (await api.ui()).focus;
    return focus?.line === line && focus?.column === source[line - 1].length + 1;
  });

  await api.act("press", { key: "Enter" });
  // A line appeared, which is neutral about WHAT was built - that is what the checks compare.
  await waitFor("Enter to add a line", async () => (await live()).length > source.length);
  return live();
}

const restore = await api.settings();
let made = false;

try {
  await api.component("add", { kind: "module", name, project: project.projectId });
  made = true;
  await api.writeModule(name, "Option Explicit\r\n", project.projectId);
  await waitFor("the new module to hold its first line", async () =>
    ((await api.readModule(name, project.projectId)).text ?? "").includes("Option Explicit"));

  await api.pane("open", { module: name, project: project.projectId });
  await waitFor("the module to be the one on screen", async () =>
    (await api.ui()).focus.model?.toLowerCase().endsWith(`/${name.toLowerCase()}`));

  /*
   * THE OPENER MUST NOT ALREADY HAVE ITS CLOSER. Smart Enter builds a block when there is one to
   * build; pressing Enter on a Sub that already has its End Sub below it is an ordinary newline,
   * and both layouts then look identical. The first version of this check measured exactly that
   * and reported a working setting dead.
   */
  const opener = ["Option Explicit", "", "Public Sub Go()", ""];
  await api.settings({ blockLayout: "comfy" });
  const comfy = (await enterOn(opener, 3)).slice(2, 8).join("|");
  await api.settings({ blockLayout: "compact" });
  const compact = (await enterOn(opener, 3)).slice(2, 8).join("|");
  check("blockLayout changes the block smart Enter builds", comfy !== compact,
    `comfy ${JSON.stringify(comfy)} and compact ${JSON.stringify(compact)} are the same`);

  const comment = ["Option Explicit", "", "' a comment", ""];
  await api.settings({ continueCommentOnNewline: true });
  const continued = (await enterOn(comment, 3))[3] ?? "";
  await api.settings({ continueCommentOnNewline: false });
  const bare = (await enterOn(comment, 3))[3] ?? "";
  check("continueCommentOnNewline decides whether the next line is a comment",
    continued.trimStart().startsWith("'") && !bare.trimStart().startsWith("'"),
    `on gave ${JSON.stringify(continued)}, off gave ${JSON.stringify(bare)}`);

  const padded = ["Option Explicit", "", "'    padded", ""];
  await api.settings({ continueCommentOnNewline: true, mirrorCommentSpacing: true });
  const mirrored = (await enterOn(padded, 3))[3] ?? "";
  await api.settings({ mirrorCommentSpacing: false });
  const unmirrored = (await enterOn(padded, 3))[3] ?? "";
  check("mirrorCommentSpacing repeats the spaces after the apostrophe",
    mirrored !== unmirrored && mirrored.length > unmirrored.length,
    `on gave ${JSON.stringify(mirrored)}, off gave ${JSON.stringify(unmirrored)}`);

  const explorer = {};
  for (const on of [true, false]) {
    await api.settings({ treeFollowsEditor: on });
    await waitFor(`treeFollowsEditor to read back as ${on}`, async () =>
      (await api.settings()).treeFollowsEditor === on);

    await api.pane("open", { module: name, project: project.projectId });
    // Settled rather than slept on, and stability is a different question from whether the two
    // snapshots DIFFER, which is what the check below asks.
    explorer[String(on)] = JSON.stringify(
      await waitUntilStable(async () => (await api.ui()).explorer));
  }

  check("treeFollowsEditor changes what the explorer shows",
    explorer.true !== explorer.false,
    "the explorer snapshot was identical with it on and off");

  /*
   * THERE ARE SIX SETTINGS, not seven. `formatCanonicalKeywords` was removed on 2026-08-09 after
   * this suite could not find a way to observe it.
   *
   * It reached the formatter and was conditional there, but two paths canonicalise keywords before
   * the formatter is ever asked and neither consults a setting: the HOST respells them as it takes
   * a module, and the page recases every touched line 200ms after it settles. Typing
   * `public sub go()` with the switch OFF still produced `Public Sub go()`.
   *
   * So the switch promised what it could not deliver. Formatting still respells, always. If a
   * `formatCanonicalKeywords` row ever comes back to the dialog, this comment is the reason to
   * ask what changed underneath it first.
   */
  const known = Object.keys(await api.settings());
  check("no setting has appeared that nothing here exercises",
    !known.includes("formatCanonicalKeywords"),
    `settings answers ${JSON.stringify(known)}. A new one needs a row above, or a reason here.`);

} finally {
  await api.settings(restore).catch(() => {});
  if (made) {
    await api.pane("close", { module: name, project: project.projectId, answer: "discard" }).catch(() => {});
    await api.component("remove", { name, project: project.projectId }).catch(() => {});
  }

  console.log(`\n${passed} passed, ${failures.length} failed`);
  for (const one of failures) { console.log(`  ${one}`); }

  process.exitCode = failures.length === 0 ? 0 : 1;
}
