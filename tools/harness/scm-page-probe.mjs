// The Source Control pane and the blame layer with NO git behind them, headless: the demo
// transport answers `scm` with state noGit, which is the first thing a developer without Git
// for Windows sees. The pane must say where git comes from, keep every control that needs a
// repository out of reach, and the blame layer must turn itself OFF and say why when the host
// refuses it - a button left pressed over nothing painted claims a state the host declined.
//
//   node tools\harness\scm-page-probe.mjs      (from PowerShell; see page-probe.mjs)

import { runPageProbe } from "./page-probe.mjs";

const DRIVE = `(async () => {
  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const until = async (what, budget = 5000) => {
    for (let waited = 0; waited < budget; waited += 100) {
      const got = what();
      if (got) { return got; }
      await sleep(100);
    }
    return null;
  };

  await until(() => document.querySelectorAll(".tab").length >= 2, 20000);
  check("the demo page opened its two tabs", document.querySelectorAll(".tab").length >= 2);

  const ui = window.xlideUi;
  const panes = ui.state().panes;
  const seat = panes.find((one) => one.name === "scm");
  check("the scm seat is listed among the panes", !!seat && seat.title === "Source Control",
    JSON.stringify(panes.map((p) => p.name)));

  const tab = document.querySelector('.panel-tab[data-panel="scm"]');
  check("the scm tab is in a dock strip", !!tab);
  const root = document.querySelector("#scm");
  check("the #scm root exists with its parts",
    !!root && !!root.querySelector("#scm-list") && !!root.querySelector("#scm-message") && !!root.querySelector("#scm-diff"));

  const before = ui.state().scm;
  check("ui.scm exists before the pane is shown and is unread",
    before !== null && before.state === "" && before.busy === false, JSON.stringify(before));

  tab.click();
  await sleep(50);
  check("clicking the tab shows the pane body", root.hidden === false, String(root.hidden));

  const scm = await until(() => { const s = ui.state().scm; return s && s.state === "noGit" ? s : null; });
  check("the demo answers noGit and the pane adopts it", scm && scm.project === "Book1.xlsm", JSON.stringify(scm));
  const empty = root.querySelector('.scm-state[data-state="noGit"]');
  check("the noGit empty state names the download",
    !!empty && empty.textContent.includes("git-scm.com/download/win"), empty && empty.textContent);
  check("the message box is hidden with no repository", root.querySelector("#scm-message").hidden === true);
  check("commit is disabled with no repository", root.querySelector("#scm-commit").disabled === true);
  check("blame is disabled with no repository", root.querySelector("#scm-blame").disabled === true);

  const idle = await until(() => (ui.state().scm.busy ? null : true));
  const pressRefresh = ui.act("scmPane", { press: "refresh" });
  check("act scmPane press=refresh answers did once the pane is idle", !!idle && pressRefresh.did === true, pressRefresh.detail);
  const pressCommit = ui.act("scmPane", { press: "commit" });
  check("act scmPane press=commit refuses while disabled", pressCommit.did === false, pressCommit.detail);
  const nothing = ui.act("scmPane", {});
  check("act scmPane with nothing asked says what to pass", nothing.did === false && /press/.test(nothing.detail), nothing.detail);
  check("actions() lists scmPane and blame", ui.actions().includes("scmPane") && ui.actions().includes("blame"));

  const blameBefore = ui.state().blame;
  check("ui.blame is off before the toggle", blameBefore.on === false && blameBefore.lines === 0, JSON.stringify(blameBefore));
  const toggled = await ui.act("blame", { which: "toggle" });
  check("act blame toggle runs the editor's action and reports the host's refusal",
    toggled.did === true && /refused/.test(toggled.detail), toggled.detail);
  const refused = await until(() => {
    const said = ui.state().scm.detail;
    return said && ui.state().blame.on === false ? said : null;
  });
  check("the host's refusal turns blame back off and lands in the pane's notice",
    !!refused && /git/i.test(refused), refused ?? JSON.stringify(ui.state().blame));
  check("ui.blame carries the refusal", typeof ui.state().blame.refused === "string" && /git/i.test(ui.state().blame.refused),
    JSON.stringify(ui.state().blame));
  check("the pane's Blame button reads unpressed", root.querySelector("#scm-blame").getAttribute("aria-pressed") === "false");
  check("and nothing is painted", document.querySelectorAll(".xlide-blame, .xlide-blame-uncommitted").length === 0);
  const bad = await ui.act("blame", { which: "sideways" });
  check("act blame refuses an unknown which", bad.did === false, bad.detail);

  const census = ui.state().census;
  check("the census still balances", census.models === census.documents, JSON.stringify(census));

  const ws = ui.state().workspace;
  check("no tab wears a history face in the demo", ws.groups.every((g) => g.tabs.every((t) => !t.face)));

  return { pass: checks.every((one) => one.ok), checks };
})()`;

await runPageProbe({ label: "xlide-scm", drive: DRIVE });
