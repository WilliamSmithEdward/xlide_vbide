/*
 * Find and Replace, across a scope wider than the module on screen.
 *
 * WHY THIS EXISTS. The find panel had no coverage of any kind, and the reason it went unnoticed is
 * the reason it is worth writing down: the driver and the observer failed in the SAME DIRECTION.
 *
 * `act("search", {query, scope: "project"})` typed the query and searched nothing, because typing
 * raises an `input` event and the only handler for it searches when the scope is "module". And the
 * field a probe would naturally check afterwards - `ui.search.matches` - is fed by the live
 * decorations in the current model, so it is structurally 0 for every scope but that one. An
 * action that did nothing and a count that could not move agreed with each other perfectly, and a
 * test written against them would have passed forever while the feature did nothing at all.
 *
 * REPLACE ALL is the reason this matters rather than merely being untidy. It rewrites text across
 * every module of a project through the host, it is the most destructive operation on this
 * surface, and until 2026-08-11 nothing could trigger it except a person with a mouse.
 *
 * THE TOKEN IS UNIQUE ON PURPOSE. A project-scope replace reaches every module in the workbook, so
 * a test that replaced a real word would rewrite the fixture. Everything here is done against a
 * nonsense token that exists only in a module this suite creates and removes.
 *
 *   node tools\harness\search-features.mjs
 */
import { open, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const name = `Seek${process.pid}`;
const TOKEN = `Zzq${process.pid}Marker`;

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
  if (ok) { passed += 1; console.log(`ok   ${what}`); }
  else { failures.push(what); console.log(`FAIL ${what}${detail ? `\n     ${detail}` : ""}`); }
};

// Three occurrences, so a replace count has something to be wrong about.
const SOURCE = [
  "Option Explicit",
  "",
  `Public Sub ${TOKEN}One()`,
  `    Debug.Print "${TOKEN}"`,
  "End Sub",
  "",
  "Public Sub Caller()",
  `    ${TOKEN}One`,
  "End Sub",
  "",
].join("\r\n");

const searchState = async () => (await api.ui()).search;

let made = false;
try {
  await api.component("add", { kind: "module", name, project: project.projectId });
  made = true;
  await api.writeModule(name, SOURCE, project.projectId);
  await waitFor("the seed to reach the module", async () =>
    ((await api.readModule(name, project.projectId)).text ?? "").includes(TOKEN));

  await api.pane("open", { module: name, project: project.projectId });
  await waitFor("the module to be the one on screen", async () =>
    (await api.ui()).focus.model?.toLowerCase().endsWith(`/${name.toLowerCase()}`));

  // ---- module scope, which is the engine that always worked ----

  await api.act("search", { query: TOKEN, scope: "module" });
  await waitFor("the module search to count its matches", async () =>
    (await searchState()).matches === 3);

  const inModule = await searchState();
  check("a module search counts every occurrence", inModule.matches === 3,
    JSON.stringify(inModule));

  // ---- project scope: the half that did nothing and said it had ----

  const typed = await api.act("search", { query: TOKEN, scope: "project" });
  check("typing a project query is reported as typing, not as searching",
    typed.did && /nothing was run/.test(typed.detail ?? ""), JSON.stringify(typed));

  const beforeRun = await searchState();
  check("and nothing has been asked of the host yet",
    beforeRun.scopedMatches === -1,
    `scopedMatches is ${beforeRun.scopedMatches}; -1 means unasked, which is what typing leaves`);

  const ran = await api.act("search", { query: TOKEN, scope: "project", run: "findAll" });
  check("Find All runs the project search", ran.did, ran.detail);

  await waitFor("the host to answer the project search", async () =>
    (await searchState()).scopedMatches >= 0);

  const found = await searchState();
  check("the project search finds the occurrences the module holds",
    found.scopedMatches === 3,
    `scopedMatches ${found.scopedMatches}, matches ${found.matches} `
    + "(matches is module-scope only and is expected to be 0 here)");

  check("and the module-scope count is NOT what answers for a project search",
    found.matches === 0,
    `matches is ${found.matches}. If this ever becomes non-zero the two engines have been `
    + "confused, and the field a probe reads by habit would start looking right by accident");

  // ---- replace all, against a token nothing else in the workbook has ----

  const replaced = await api.act("search", {
    query: TOKEN,
    scope: "project",
    replacement: `${TOKEN}Done`,
    run: "replaceAll",
  });
  check("Replace All runs", replaced.did, replaced.detail);

  await waitFor("the replace to reach the module", async () =>
    ((await api.readModule(name, project.projectId)).text ?? "").includes(`${TOKEN}Done`));

  const after = (await api.readModule(name, project.projectId)).text ?? "";
  const remaining = after.split(`${TOKEN}Done`).length - 1;
  check("every occurrence was rewritten", remaining === 3,
    `${remaining} occurrence(s) of the replacement, expected 3`);

  check("and the replaced count is reported", (await searchState()).scopedReplaced === 3,
    JSON.stringify(await searchState()));

  // The workbook, the surface and the analyzer must agree after a host-driven rewrite - WAITED
  // for, not sampled. A replace across a project rewrites the module underneath the surface and
  // the two converge a beat later; asking once the instant the text lands catches them mid-hop
  // and reports a defect that is only a race in the asking.
  const sync = await waitFor("the three copies to agree after the replace",
    async () => {
      const seen = await api.inSync();
      return seen.agreed ? seen : null;
    },
    { budgetMs: 12000 }).catch(async () => api.inSync());

  check("the three copies agree after a project-wide replace", sync.agreed,
    JSON.stringify({
      contentAgrees: sync.contentAgrees,
      native: sync.nativeModule,
      surface: sync.surfaceModule,
      page: sync.pageModule,
      nativeLines: sync.nativeLines,
      surfaceLines: sync.surfaceLines,
    }));

  await api.act("search", { close: 1 });
  check("the find box closes", !(await searchState()).open);
} finally {
  await api.act("search", { close: 1 }).catch(() => {});

  if (made) {
    await api.pane("close", { module: name, project: project.projectId, answer: "discard" }).catch(() => {});
    await api.component("remove", { name, project: project.projectId }).catch(() => {});
  }

  console.log(`\n${passed} passed, ${failures.length} failed`);
  for (const one of failures) { console.log(`  ${one}`); }

  process.exitCode = failures.length === 0 ? 0 : 1;
}
