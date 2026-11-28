// Which keys are Vim's while Vim mode is on (#173), tested by calling the policy.
//
// The policy is a pure function of the keymap's name for a press, the mode, and whether AltGr
// typed a character, so it is pinned here. The live suite (tools\harness\vim.mjs) presses the
// keys; these fail in milliseconds rather than in Excel.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scratch = await mkdtemp(path.join(tmpdir(), "xlide-vimkeys-"));

const outfile = path.join(scratch, "vimkeys.mjs");
await build({
  entryPoints: [path.join(root, "src", "vimkeys.ts")],
  outfile,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
});
const { routeVimKey } = await import(pathToFileURL(outfile).href);

const normal = { insertMode: false, altGraphCharacter: null };
const insert = { insertMode: true, altGraphCharacter: null };

const checks = [];
const check = (name, body) => checks.push([name, body]);

check("plain, shifted and named keys are Vim's in every mode", () => {
  for (const key of ["'j'", "'G'", "':'", "Esc", "Enter", "Backspace", "Left", "Shift-Left", "Space"]) {
    assert.equal(routeVimKey(key, normal), key, key);
    assert.equal(routeVimKey(key, insert), key, key);
  }
});

check("Ctrl+[ is Escape in every mode", () => {
  assert.equal(routeVimKey("Ctrl-[", normal), "Ctrl-[");
  assert.equal(routeVimKey("Ctrl-[", insert), "Ctrl-[");
});

check("Ctrl+Q, Ctrl+D and Ctrl+U are Vim's outside insert mode only", () => {
  for (const key of ["Ctrl-q", "Ctrl-d", "Ctrl-u"]) {
    assert.equal(routeVimKey(key, normal), key, key);
    assert.equal(routeVimKey(key, insert), null, `${key} in insert mode`);
  }
});

check("every other Ctrl chord stays the editor's, though the keymap binds it", () => {
  // Each of these is a default binding in the keymap: redo, block visual, Escape, page down and
  // up, scroll, jump back and forward, increment and decrement, delete word, next and previous.
  const bound = ["r", "v", "c", "f", "b", "e", "y", "o", "i", "a", "x", "w", "n", "p", "t"];
  for (const letter of bound) {
    assert.equal(routeVimKey(`Ctrl-${letter}`, normal), null, `Ctrl-${letter}`);
    assert.equal(routeVimKey(`Ctrl-${letter}`, insert), null, `Ctrl-${letter} in insert mode`);
  }
});

check("a Ctrl chord with another modifier is the editor's too", () => {
  for (const key of ["Shift-Ctrl-z", "Ctrl-Alt-d", "Shift-Ctrl-d", "Shift-Ctrl-q", "Ctrl-Space", "Ctrl-Enter"]) {
    assert.equal(routeVimKey(key, normal), null, key);
  }
});

check("Alt chords pass to the keymap, which binds none and lets them through", () => {
  assert.equal(routeVimKey("Alt-f", normal), "Alt-f");
});

check("a character typed with AltGr is that character, not the Ctrl+Alt chord it arrives as", () => {
  const typed = (character) => ({ insertMode: false, altGraphCharacter: character });
  assert.equal(routeVimKey("Ctrl-Alt-8", typed("[")), "'['");
  assert.equal(routeVimKey("Ctrl-Alt-7", typed("{")), "'{'");
  assert.equal(routeVimKey("Ctrl-Alt-q", typed("@")), "'@'");
  assert.equal(routeVimKey("Ctrl-Alt-'", typed("'")), "'''");
});

check("AltGr held over a key that types nothing is judged as the chord it is", () => {
  const held = { insertMode: false, altGraphCharacter: "ArrowLeft" };
  assert.equal(routeVimKey("Ctrl-Alt-Left", held), null);
  assert.equal(routeVimKey("Ctrl-d", { insertMode: false, altGraphCharacter: "Dead" }), "Ctrl-d");
});

let failed = 0;
try {
  for (const [name, body] of checks) {
    try {
      body();
      console.log(`  ok    ${name}`);
    } catch (error) {
      failed++;
      console.log(`  FAIL  ${name}\n        ${error.message}`);
    }
  }
} finally {
  await rm(scratch, { recursive: true, force: true });
}

console.log(failed === 0 ? `vim keys: ${checks.length} checks passed` : `vim keys: ${failed} of ${checks.length} FAILED`);
process.exit(failed === 0 ? 0 : 1);
