/*
 * What a project's own type library references change about its analysis.
 *
 * A workbook that references Word compiles `Dim wd As Word.Application`; one that does not
 * refuses it, and from the analyzer's 10.0.0 that refusal is a finding rather than silence -
 * `missing-library-reference`, an error. So the SAME code is clean or wrong depending on a fact
 * that lives nowhere in the text, and this engine only learns it because the add-in reads the
 * live project's References and sends their GUIDs on the seed.
 *
 * The middle case is the whole suite: before that plumbing there was only the error, on correct
 * code, in every project that automates another application.
 *
 *   node test/project-references.mjs
 */

import { startEngine } from "./harness.mjs";

const { call, stop } = await startEngine("project-references");
const CRLF = "\r\n";

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
  console.log(`  ${ok ? "ok  " : "FAIL"} ${what}${detail !== undefined ? `  -- ${String(detail).slice(0, 130)}` : ""}`);
  if (ok) { passed += 1; } else { failures.push(what); }
};

await call("initialize", {});

// Identities, as a project's references declare them. Word and Excel are both modelled upstream;
// stdole is a real reference every VBA project carries and nothing models, which is the case that
// proves an unknown library is ignored rather than guessed at.
const WORD = "{00020905-0000-0000-C000-000000000046}";
const EXCEL = "{00020813-0000-0000-C000-000000000046}";
const STDOLE = "{00020430-0000-0000-C000-000000000046}";

const AUTOMATES_WORD = [
  "Option Explicit",
  "",
  "Public Sub Automate()",
  "    Dim wd As Word.Application",
  "    Set wd = New Word.Application",
  "    wd.Visible = True",
  "End Sub",
].join(CRLF);

let nextProject = 0;

/** Seeds one project with the given references and answers the diagnostics for its one module. */
async function analyse(source, referenceGuids) {
  const projectId = `refs${nextProject}`;
  const moduleName = `Refs${nextProject}`;
  nextProject += 1;

  await call("project/open", {
    projectId,
    generation: 1,
    modules: [{ moduleName, source, type: "standard" }],
    ...(referenceGuids === undefined ? {} : { referenceGuids }),
  });

  const { diagnostics } = await call("textDocument/diagnostics", {
    documentKey: `${projectId}/${moduleName}`,
    projectId,
    generation: 1,
    source,
    moduleName,
    moduleType: "standard",
  });

  return { projectId, moduleName, diagnostics: diagnostics ?? [] };
}

const missingLibrary = (diagnostics) =>
  diagnostics.filter((one) => one.code === "missing-library-reference");

/* ---- the fact the text cannot carry --------------------------------------------------------- */

const unreferenced = await analyse(AUTOMATES_WORD, undefined);
check("a project that sends no references reports the Word types as unreferenced",
  missingLibrary(unreferenced.diagnostics).length === 1,
  missingLibrary(unreferenced.diagnostics).map((one) => one.message).join("; ") || "(none)");

const referenced = await analyse(AUTOMATES_WORD, [WORD]);
check("THE SAME CODE is clean once the project says it references Word",
  missingLibrary(referenced.diagnostics).length === 0,
  missingLibrary(referenced.diagnostics).map((one) => one.message).join("; ") || "(none)");

const wrongLibrary = await analyse(AUTOMATES_WORD, [EXCEL]);
check("referencing a DIFFERENT application does not excuse it",
  missingLibrary(wrongLibrary.diagnostics).length === 1,
  missingLibrary(wrongLibrary.diagnostics).map((one) => one.code).join(",") || "(none)");

const unmodelled = await analyse(AUTOMATES_WORD, [STDOLE]);
check("a library nothing models is ignored rather than taken for a host",
  missingLibrary(unmodelled.diagnostics).length === 1,
  missingLibrary(unmodelled.diagnostics).map((one) => one.code).join(",") || "(none)");

const empty = await analyse(AUTOMATES_WORD, []);
check("an empty reference list reads as no references, not as an error",
  missingLibrary(empty.diagnostics).length === 1,
  missingLibrary(empty.diagnostics).map((one) => one.code).join(",") || "(none)");

/* ---- the fix is offered under the same knowledge as the squiggle ----------------------------- */

// The worker takes the referenced hosts per ANALYSIS, not per project, so every path that
// analyses has to attach them. Two paths do, and a fix computed under knowledge the squiggle was
// never drawn under is how they last disagreed (the 2026-08-19 host hunt).
//
// ASKED AGAINST EDITED TEXT ON PURPOSE. A fix request whose source matches the last analysis is
// answered from the memo and never analyses at all, so asking with the text just diagnosed
// measures the memo and would pass with this whole plumbing torn out (measured, 2026-09-21). One
// appended comment is the keystroke the fresh pass exists for.
const EDITED = `${AUTOMATES_WORD}${CRLF}' one keystroke later`;
const span = missingLibrary(unreferenced.diagnostics)[0]?.span ?? { start: 0, end: 0 };

const offered = async ({ projectId, moduleName }) => {
  const { actions } = await call("textDocument/codeAction", {
    projectId, moduleName, source: EDITED,
    start: span.start, end: span.end, moduleType: "standard",
  });
  return actions;
};

const withoutFixes = await offered(unreferenced);
check("the fix path DOES answer at that span when the reference is missing",
  withoutFixes.some((one) => one.code === "missing-library-reference"),
  withoutFixes.map((one) => one.title).join(" | ") || "(no actions)");

const withFixes = await offered(referenced);
check("and offers no fix for it once the project references Word",
  withFixes.every((one) => one.code !== "missing-library-reference"),
  withFixes.map((one) => one.title).join(" | ") || "(no actions)");

/* ---- a reference added to a project already open ---------------------------------------------- */

// THE GESTURE THE FEATURE IS FOR: the developer reads `missing-library-reference`, opens the
// References dialog, ticks Word. The module is not touched, so every cheap gate between here and
// the answer sees identical inputs - the host re-seeds the same project id with the same text and
// one more library, and the answer has to move anyway. It did not, first time: the engine's own
// per-document memo compared source, cross-module facts and request shape, and the references
// were in none of them, so the error stayed on a screen where Word was now referenced.
const OPEN_PROJECT = "already-open";
await call("project/open", {
  projectId: OPEN_PROJECT, generation: 1,
  modules: [{ moduleName: "Live", source: AUTOMATES_WORD, type: "standard" }],
});
const diagnoseOpen = (generation) => call("textDocument/diagnostics", {
  documentKey: `${OPEN_PROJECT}/Live`, projectId: OPEN_PROJECT, generation,
  source: AUTOMATES_WORD, moduleName: "Live", moduleType: "standard",
});

const beforeTick = await diagnoseOpen(1);
check("the open project reports it before the reference is added",
  missingLibrary(beforeTick.diagnostics ?? []).length === 1,
  missingLibrary(beforeTick.diagnostics ?? []).map((one) => one.code).join(",") || "(none)");

await call("project/open", {
  projectId: OPEN_PROJECT, generation: 2,
  modules: [{ moduleName: "Live", source: AUTOMATES_WORD, type: "standard" }],
  referenceGuids: [WORD],
});
const afterTick = await diagnoseOpen(2);
check("ticking Word clears it WITHOUT the module being touched",
  missingLibrary(afterTick.diagnostics ?? []).length === 0,
  missingLibrary(afterTick.diagnostics ?? []).map((one) => one.code).join(",") || "(none)");

await call("project/open", {
  projectId: OPEN_PROJECT, generation: 3,
  modules: [{ moduleName: "Live", source: AUTOMATES_WORD, type: "standard" }],
});
const afterUntick = await diagnoseOpen(3);
check("and unticking it brings the error back, same text again",
  missingLibrary(afterUntick.diagnostics ?? []).length === 1,
  missingLibrary(afterUntick.diagnostics ?? []).map((one) => one.code).join(",") || "(none)");

/* ---- and they are forgotten with the project ------------------------------------------------- */

await call("project/close", { projectId: referenced.projectId });
await call("project/open", {
  projectId: referenced.projectId,
  generation: 2,
  modules: [{ moduleName: referenced.moduleName, source: AUTOMATES_WORD, type: "standard" }],
});
const reopened = await call("textDocument/diagnostics", {
  documentKey: `${referenced.projectId}/${referenced.moduleName}`,
  projectId: referenced.projectId,
  generation: 2,
  source: AUTOMATES_WORD,
  moduleName: referenced.moduleName,
  moduleType: "standard",
});
check("closing a project forgets its references - a reopen does not inherit them",
  missingLibrary(reopened.diagnostics ?? []).length === 1,
  missingLibrary(reopened.diagnostics ?? []).map((one) => one.code).join(",") || "(none)");

await stop();
console.log(`\n${passed} passed, ${failures.length} failed`);
if (failures.length > 0) { failures.forEach((one) => console.log(`  ${one}`)); }
process.exit(failures.length > 0 ? 1 : 0);
