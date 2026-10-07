import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scratch = await mkdtemp(path.join(tmpdir(), "xlide-preferences-"));
const compiled = path.join(scratch, "preferences.mjs");
await build({ entryPoints: [path.join(root, "src/preferences.ts")], outfile: compiled,
  bundle: true, format: "esm", platform: "node", logLevel: "silent" });
const cache = new Map();
const previousStorage = globalThis.localStorage;
globalThis.localStorage = { getItem: key => cache.get(key) ?? null, setItem: (key, value) => cache.set(key, value) };
try {
  const fresh = await import(pathToFileURL(compiled).href + "?fresh");
  const writes = [];
  assert.equal(fresh.readUiPreference("paneLayout"), null);
  fresh.writeUiPreference("paneLayout", "default");
  fresh.watchUiPreference("paneLayout", value => fresh.writeUiPreference("paneLayout", value));
  fresh.restoreUiPreferences({ paneLayout: "saved by Excel" }, (key, value) => writes.push([key, value]));
  assert.deepEqual(writes, [["paneLayout", "saved by Excel"]]);
  console.log("ok   fresh-process defaults cannot overwrite the shared layout before hydration");

  const reload = await import(pathToFileURL(compiled).href + "?reload");
  assert.equal(reload.readUiPreference("paneLayout"), "saved by Excel");
  reload.writeUiPreference("paneLayout", "new drag");
  reload.watchUiPreference("paneLayout", value => reload.writeUiPreference("paneLayout", value));
  const migrated = [];
  reload.restoreUiPreferences({ paneLayout: "older host snapshot" }, (key, value) => migrated.push([key, value]));
  assert.deepEqual(migrated, [["paneLayout", "new drag"]]);
  console.log("ok   a page reload and legacy cache preserve the latest in-session arrangement");

  reload.writeUiPreference("sourceControl", "divider");
  assert.deepEqual(migrated.at(-1), ["sourceControl", "divider"]);
  assert.equal(reload.readUiPreference("paneLayout"), "new drag");
  console.log("ok   later geometry updates travel separately from the pane layout");
} finally {
  if (previousStorage === undefined) delete globalThis.localStorage;
  else globalThis.localStorage = previousStorage;
  await rm(scratch, { recursive: true, force: true });
}
