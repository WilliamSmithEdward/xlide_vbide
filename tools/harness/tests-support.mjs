/*
 * The install chip answers for the file the select is pointing at.
 *
 * THE DEFECT THIS EXISTS FOR (the owner, 2026-08-21). The chip read "XlideAssert Installed", in
 * green, with a workbook selected that plainly did not have it. Three things were wrong at once:
 *
 *   - the pane listed a file only while it held tests, carried the support module, or was the one
 *     being worked in - so a workbook with none of those was not in the file select at all, and
 *     the one file a developer most needs to point the install at was the one they could not
 *     choose;
 *   - nothing republished the pane when the developer moved between files, so even the list it
 *     did have was whatever the last install or refresh had left behind - the route answered
 *     correctly on every read while the pane drew a picture 36 minutes old;
 *   - the session's one-word answer counted only files that HOLD TESTS, which with nothing open
 *     holding tests is vacuously satisfied. Hence a green chip whose tooltip said "every open
 *     file that holds tests carries an XlideAssert" - true, and about no files at all.
 *
 * So: the pane lists every file the tree knows, moving between files repaints it, and the chip is
 * the selected file's - or, on All Files, the worst standing among all of them.
 *
 * NEEDS SEVERAL FILES WITH DIFFERENT ANSWERS, which is the whole point: one file, or several that
 * agree, cannot tell a chip that follows the select from one that does not.
 *
 *   node tools\harness\tests-support.mjs
 */

import { open, waitFor, comparingReporter } from "./xlide-api.mjs";

const api = await open();
const { check, done } = comparingReporter();

const projects = (await api.projects()).projects;
if (projects.length < 2) {
  console.log(`FAIL this suite needs two or more files open; found ${projects.length}`);
  console.log("\n0 passed, 1 failed");
  process.exit(1);
}

/** The chip and the select, as drawn. */
const drawn = async () => {
  const said = await api.ask(`JSON.stringify({
    text: document.querySelector('#tests-install').textContent,
    title: document.querySelector('#tests-install').title,
    disabled: document.querySelector('#tests-install').disabled,
    ok: document.querySelector('#tests-install').classList.contains('tests-install-ok'),
    hidden: document.querySelector('#tests-scope-file').hidden,
    value: document.querySelector('#tests-scope-file').value,
    options: [...document.querySelector('#tests-scope-file').options].map((o) => ({ value: o.value, label: o.textContent }))
  })`);
  return JSON.parse(typeof said === "string" ? said : JSON.stringify(said));
};

/** Points the select, through the change event the control itself fires. */
const point = (value) => api.ask(
  `(() => { const s = document.querySelector("#tests-scope-file");`
  + ` s.value = ${JSON.stringify(value)};`
  + ` s.dispatchEvent(new Event("change")); return s.value; })()`);

// What each file held before this ran, so the session goes back as found.
const before = new Map((await api.tests()).files.map((file) => [file.file, file.support]));
console.log(`open: ${[...before].map(([file, support]) => `${file}/${support}`).join(", ")}\n`);

const at = await drawn();

check("every open file is offered, whatever it holds",
  at.options.slice(1).map((one) => one.label.replace(/ \(\d+\)$/, "")).sort(),
  projects.map((one) => one.project).sort());

check("and the select is showing, because there is a choice to make", at.hidden, false);

// ---- the chip follows the select ------------------------------------------------------------

const missing = [...before].filter(([, support]) => support !== "installed").map(([file]) => file);
const installed = [...before].filter(([, support]) => support === "installed").map(([file]) => file);

if (missing.length === 0 || installed.length === 0) {
  // Make the session disagree with itself, which is the state the chip has to get right.
  const target = missing.length === 0 ? [...before.keys()][0] : missing[0];
  await api.tests({ action: "install", file: target });
  await waitFor("the install to land", async () =>
    (await api.tests()).files.some((file) => file.support === "installed"));
}

const now = (await api.tests()).files;
const holds = now.filter((file) => file.support === "installed").map((file) => file.file);
const lacks = now.filter((file) => file.support !== "installed").map((file) => file.file);
console.log(`\n  installed in: ${holds.join(", ") || "(none)"}`);
console.log(`  missing from: ${lacks.join(", ") || "(none)"}\n`);

check("the session has files that disagree, so the checks below are not vacuous",
  holds.length > 0 && lacks.length > 0);

const keyFor = (file) => `file:${file.toLowerCase()}`;

for (const file of holds.slice(0, 1)) {
  await point(keyFor(file));
  const said = await drawn();
  check(`pointing at ${file}, which has it, reads installed`,
    { text: said.text, ok: said.ok, disabled: said.disabled },
    { text: "XlideAssert Installed", ok: true, disabled: true });
  check(`and its tooltip names that file`, said.title.includes(file));
}

for (const file of lacks.slice(0, 1)) {
  await point(keyFor(file));
  const said = await drawn();
  check(`pointing at ${file}, which does not, offers the install`,
    { ok: said.ok, disabled: said.disabled }, { ok: false, disabled: false });
  check(`and its tooltip names that file`, said.title.includes(file));
}

await point("@allfiles");
const all = await drawn();
check("All Files takes the worst of them, not the best", all.ok, false);

// ---- and it follows the session ---------------------------------------------------------------

// Installing everywhere the chip speaks for is what its label promises when it says a count.
await api.tests({ action: "install" });
await waitFor("every file to carry it", async () =>
  (await api.tests()).files.every((file) => file.support === "installed"));

await point("@allfiles");
const after = await drawn();
check("with every file carrying it, All Files reads installed",
  { text: after.text, ok: after.ok }, { text: "XlideAssert Installed", ok: true });

// ---- put the session back ---------------------------------------------------------------------

for (const [file, support] of before) {
  if (support === "installed") {
    continue;
  }
  const project = projects.find((one) => one.project === file);
  if (project) {
    await api.component("remove", { name: "XlideAssert", project: project.projectId })
      .catch((error) => console.log(`     WARNING: XlideAssert left in ${file} (${error.message})`));
  }
}

process.exit(done());
