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
import { open } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const name = `Bite${process.pid}`;
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
  if (ok) { passed += 1; console.log(`ok   ${what}`); }
  else { failures.push(what); console.log(`FAIL ${what}${detail ? `\n     ${detail}` : ""}`); }
};

const live = async () =>
  ((await api.readModule(name, project.projectId, { live: true })).text ?? "")
    .split("\n").map((one) => one.replace("\r", ""));

/** Seeds the module, puts the caret at the end of a line, and presses Enter. */
async function enterOn(source, line) {
  await api.writeModule(name, source.join("\r\n"), project.projectId);
  await wait(1500);
  await api.caret(line, { module: name, project: project.projectId, column: source[line - 1].length + 1 });
  await wait(500);
  await api.act("press", { key: "Enter" });
  await wait(900);
  return live();
}

const restore = await api.settings();
let made = false;

try {
  await api.component("add", { kind: "module", name, project: project.projectId });
  made = true;
  await api.writeModule(name, "Option Explicit\r\n", project.projectId);
  await wait(1200);
  await api.pane("open", { module: name, project: project.projectId });
  await wait(2500);

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
    await wait(600);
    await api.pane("open", { module: name, project: project.projectId });
    await wait(1500);
    explorer[String(on)] = JSON.stringify((await api.ui()).explorer);
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
