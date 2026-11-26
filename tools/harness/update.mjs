/*
 * Self-update, driven end to end against a release of this suite's own.
 *
 * The updater asks a feed for the newest release, judges it against the running build, downloads
 * the installer while checking its length and SHA-256 against what the release published, and
 * starts it. Every step of that is here, without GitHub and without installing anything: the suite
 * serves the feed itself on a loopback port (the `update` route's debug-only `feed=`), and the
 * "installer" it serves is a copy of a Windows program that exits at once - so "the installer was
 * started" is observable, and so is the shim noticing that it ended before Excel did, which is the
 * state a real install never reaches because Excel is closed by then.
 *
 * Driven two ways, deliberately. The `update` route is the api's half; the dialog is the
 * developer's, and the install here is pressed as a BUTTON in that dialog (`act update
 * button=install`), because the route reaching the download proves nothing about the card a
 * person sees.
 *
 * Run against the debug fixture:
 *   node tools\harness\update.mjs
 */

import { createHash } from "node:crypto";
import { existsSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createServer } from "node:http";
import { join } from "node:path";
import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const { check, done } = reporter();

/** Polls the bare route until the phase leaves the ones named, within a budget. */
const settle = async (leaving, budgetMs = 30000) => {
  const started = Date.now();
  let state = await api.update();
  while (leaving.includes(state.phase) && Date.now() - started < budgetMs) {
    await wait(150);
    state = await api.update();
  }
  return state;
};

const dialogs = async () => (await api.ui()).dialogs.map((row) => row.id);
const untilDialog = async (wanted, budgetMs = 8000) => {
  const started = Date.now();
  let shown = await dialogs();
  while (shown.includes("update-card") !== wanted && Date.now() - started < budgetMs) {
    await wait(150);
    shown = await dialogs();
  }
  return shown.includes("update-card") === wanted;
};

// The stand-in installer: Windows' own whoami.exe, which prints the user and exits 0. Signed by
// Microsoft, so no policy on the machine objects to it, and over in a blink.
const standIn = join(process.env.SystemRoot ?? "C:\\Windows", "System32", "whoami.exe");
const installerBytes = readFileSync(standIn);
const installerDigest = createHash("sha256").update(installerBytes).digest("hex");

const before = await api.update();
const installed = before.installed;
check("the route answers the installed version", /^\d+\.\d+\.\d+$/.test(installed), installed);
check("a shipped build reads GitHub", before.feed === "https://api.github.com/repos/WilliamSmithEdward/xlide_vbide/releases/latest", before.feed);
check("the state starts idle or settled", ["idle", "upToDate", "available", "skipped", "failed"].includes(before.phase), before.phase);

/** One release's JSON, in the shape GitHub's releases API answers, pointed at this server. */
const release = (origin, { tag = "v99.0.0", digest = `sha256:${installerDigest}`, asset = "xlide-setup.exe", size = installerBytes.length } = {}) => ({
  html_url: `${origin}/release/${tag}`,
  tag_name: tag,
  name: tag,
  draft: false,
  prerelease: false,
  published_at: "2026-11-25T12:00:00Z",
  body: "A release of the harness's own.\r\n\r\n- Proves the updater end to end.\r\n- Installs nothing.",
  assets: [
    { name: "security-report.md", size: 12, digest: "sha256:" + "0".repeat(64), browser_download_url: `${origin}/security-report.md` },
    { name: asset, size, digest, content_type: "application/x-msdos-program", browser_download_url: `${origin}/xlide-setup.exe` },
  ],
});

let origin = "";
const server = createServer((request, response) => {
  const path = new URL(request.url, origin).pathname;
  const json = (body) => {
    const text = JSON.stringify(body);
    response.writeHead(200, { "content-type": "application/json; charset=utf-8", "content-length": Buffer.byteLength(text) });
    response.end(text);
  };

  switch (path) {
    case "/newer": return json(release(origin));
    case "/older": return json(release(origin, { tag: "v0.0.1" }));
    case "/badhash": return json(release(origin, { digest: "sha256:" + "f".repeat(64) }));
    case "/noasset": return json(release(origin, { asset: "xlide-setup.zip" }));
    case "/short": return json(release(origin, { size: installerBytes.length + 1000 }));
    case "/xlide-setup.exe":
      response.writeHead(200, { "content-type": "application/x-msdos-program", "content-length": installerBytes.length });
      return response.end(installerBytes);
    default:
      response.writeHead(404);
      return response.end("not here");
  }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
origin = `http://127.0.0.1:${server.address().port}`;

const preferenceRoot = process.env.XLIDE_TEST_PREFERENCES_ROOT ?? process.env.LOCALAPPDATA;
const uiStateFile = join(preferenceRoot, "xlide_vbide", "ui-state.json");
const updatesFolder = join(preferenceRoot, "xlide_vbide", "updates");
const uiStateBefore = existsSync(uiStateFile) ? readFileSync(uiStateFile, "utf8") : null;

try {
  console.log("self-update, against a feed of this suite's own\n");

  // ---- the check: a newer release is offered, and the dialog comes up on its own ----
  await api.update({ action: "reset" });
  let state = await api.update({ feed: `${origin}/newer`, action: "check" });
  check("a check answers at once, checking", state.phase === "checking", state.phase);
  state = await settle(["checking"]);
  check("the newer release is available", state.phase === "available", `${state.phase} ${state.failure ?? ""}`);
  check("it names the release", state.version === "99.0.0" && state.tag === "v99.0.0", `${state.version} ${state.tag}`);
  check("it carries the installer's size", state.size === installerBytes.length, String(state.size));
  check("it carries the notes", /Proves the updater end to end/.test(state.notes ?? ""), (state.notes ?? "").slice(0, 40));
  check("it records when it checked", /^\d{4}-\d{2}-\d{2}T/.test(state.checkedAt ?? ""), state.checkedAt);
  check("the update dialog opened for the news", await untilDialog(true), (await dialogs()).join(", "));

  // Release notes are not opened from a loopback feed: a browser is not where a suite belongs.
  let refused = null;
  try { await api.update({ action: "notes" }); } catch (error) { refused = error.message; }
  check("release notes are refused for a loopback feed", /GitHub feed/.test(refused ?? ""), refused);

  // ---- install, pressed in the dialog ----
  const pressed = await api.act("update", { button: "install" });
  check("the dialog's Install button is pressed", pressed.did === true, pressed.detail);
  state = await settle(["available", "downloading", "verified", "launched"], 30000);
  check("the installer was started and, Excel still being up, ended", state.phase === "exited", `${state.phase} ${state.failure ?? ""}`);
  check("it exited 0", state.exitCode === 0, String(state.exitCode));
  check("it had a pid", state.pid > 0, String(state.pid));
  check("the download is whole", state.received === installerBytes.length && state.size === installerBytes.length, `${state.received} of ${state.size}`);
  check("the verified installer is on disk", state.path !== null && existsSync(state.path), state.path);
  if (state.path && existsSync(state.path)) {
    const onDisk = readFileSync(state.path);
    check("the file on disk is the bytes served", createHash("sha256").update(onDisk).digest("hex") === installerDigest, String(onDisk.length));
    check("it is named for the version", /xlide-setup-99\.0\.0\.exe$/.test(state.path), state.path);
    check("it lives beside the logs", state.path.toLowerCase().startsWith(updatesFolder.toLowerCase()), updatesFolder);
  }
  check("no partial download is left", !readdirSync(updatesFolder).some((name) => name.endsWith(".partial")), readdirSync(updatesFolder).join(", "));
  check("the dialog says the installer closed", await untilDialog(true), (await dialogs()).join(", "));
  const closed = await api.act("update", { button: "close" });
  check("Close takes the dialog down", closed.did === true && (await untilDialog(false)), closed.detail);

  // A second install finds the verified file and skips the download.
  state = await api.update({ action: "install" });
  state = await settle(["downloading", "verified", "launched"], 30000);
  check("a verified file already there is reused", state.phase === "exited" && state.received === installerBytes.length, `${state.phase} ${state.received}`);
  await api.act("update", { button: "close" });

  // ---- a download whose hash disagrees is refused and deleted ----
  await api.update({ action: "reset" });
  await api.update({ feed: `${origin}/badhash`, action: "check" });
  state = await settle(["checking"]);
  check("the bad-digest release is offered like any other", state.phase === "available", state.phase);
  await api.update({ action: "install" });
  state = await settle(["available", "downloading", "verified"], 30000);
  check("the download whose SHA-256 disagrees fails", state.phase === "failed" && /SHA-256/.test(state.failure ?? ""), `${state.phase} ${state.failure ?? ""}`);
  check("nothing runs and nothing is left", state.pid === 0 && !readdirSync(updatesFolder).some((name) => name.endsWith(".partial")), readdirSync(updatesFolder).join(", "));
  check("the dialog says so", await untilDialog(true), (await dialogs()).join(", "));
  await api.act("update", { button: "close" });

  // ---- a download shorter than the release says is refused ----
  await api.update({ action: "reset" });
  await api.update({ feed: `${origin}/short`, action: "check" });
  await settle(["checking"]);
  await api.update({ action: "install" });
  state = await settle(["available", "downloading", "verified"], 30000);
  check("a download shorter than the release says fails", state.phase === "failed" && /bytes/.test(state.failure ?? ""), `${state.phase} ${state.failure ?? ""}`);
  await api.act("update", { button: "close" });

  // ---- an older release is up to date, a release with no installer is refused ----
  await api.update({ action: "reset" });
  await api.update({ feed: `${origin}/older`, action: "check" });
  state = await settle(["checking"]);
  check("an older release leaves the build up to date", state.phase === "upToDate" && state.version === "0.0.1", `${state.phase} ${state.version}`);
  check("a check by hand shows its answer", await untilDialog(true), (await dialogs()).join(", "));
  refused = null;
  try { await api.update({ action: "install" }); } catch (error) { refused = error.message; }
  check("install is refused while up to date, as the dialog offers none", /nothing to install/.test(refused ?? ""), refused);
  await api.act("update", { button: "close" });

  await api.update({ action: "reset" });
  await api.update({ feed: `${origin}/noasset`, action: "check" });
  state = await settle(["checking"]);
  check("a release without xlide-setup.exe is refused", state.phase === "failed" && /no xlide-setup\.exe/.test(state.failure ?? ""), `${state.phase} ${state.failure ?? ""}`);
  await api.act("update", { button: "close" });

  // ---- a feed off the loopback is refused before anything is asked ----
  refused = null;
  try { await api.update({ feed: "http://example.com/latest" }); } catch (error) { refused = error.message; }
  check("a feed off 127.0.0.1 is refused", /127\.0\.0\.1/.test(refused ?? ""), refused);

  // ---- skip: the version is declined, remembered, and not offered again ----
  await api.update({ action: "reset" });
  await api.update({ feed: `${origin}/newer`, action: "check" });
  await settle(["checking"]);
  await untilDialog(true);
  const skipped = await api.act("update", { button: "skip" });
  check("the dialog's Skip button is pressed", skipped.did === true, skipped.detail);
  await wait(500);
  state = await api.update();
  check("the version is skipped", state.phase === "skipped" && state.skippedVersion === "99.0.0", `${state.phase} ${state.skippedVersion}`);
  check("the choice is written to ui-state.json", /"updateSkipped":\s*"99\.0\.0"/.test(readFileSync(uiStateFile, "utf8")));
  await api.update({ action: "reset" });
  await api.update({ action: "check" });
  state = await settle(["checking"]);
  check("a later check finds it skipped", state.phase === "skipped", state.phase);
  await untilDialog(true);
  await api.act("update", { button: "close" });

  // ---- the setting, through the route and the page ----
  const settingBefore = (await api.settings()).updateCheckOnStart;
  check("the setting is on by default here", settingBefore === true, String(settingBefore));
  let changed = await api.settings({ updateCheckOnStart: false });
  check("the route turns the check on start off", changed.updateCheckOnStart === false);
  const viaPage = await api.act("settings", { key: "updateCheckOnStart", value: "true" });
  await wait(400);
  changed = await api.settings();
  check("the dialog's own path turns it back on", viaPage.did === true && changed.updateCheckOnStart === true, `${viaPage.detail}; ${changed.updateCheckOnStart}`);

  // ---- the About card knows ----
  await api.act("toolbar", { command: "openHelp" });
  await wait(300);
  const status = await api.ask("document.getElementById('help-update-status')?.textContent ?? null");
  check("the About card says what the updater knows", /99\.0\.0/.test(String(status ?? "")), JSON.stringify(status));
  await api.ask("document.getElementById('help-close')?.click(); true");
  check("About closes", await untilDialog(false) && !(await dialogs()).includes("help-card"), (await dialogs()).join(", "));

  // ---- the route is taught to agents ----
  const row = await api.agentRoute("update");
  check("agent/routes teaches the update route", row.name === "update" && row.bareGetIsSafe === true && /install/.test(row.what), row.args);
} finally {
  // Back to GitHub, idle, and the developer's own skipped-version choice as it was.
  try {
    await api.update({ feed: "", action: "reset" });
  } catch (error) {
    console.log(`  (could not put the updater back: ${error.message})`);
  }

  if (uiStateBefore !== null) {
    writeFileSync(uiStateFile, uiStateBefore);
  } else if (existsSync(uiStateFile)) {
    const parsed = JSON.parse(readFileSync(uiStateFile, "utf8"));
    delete parsed.updateSkipped;
    writeFileSync(uiStateFile, JSON.stringify(parsed, null, 2));
  }

  for (const name of existsSync(updatesFolder) ? readdirSync(updatesFolder) : []) {
    if (/^xlide-setup-99\.0\.0\.exe/.test(name)) {
      try { rmSync(join(updatesFolder, name)); } catch { /* left for the next run to replace */ }
    }
  }

  server.close();
}

process.exit(done());
