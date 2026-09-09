// What the blame hint says, tested by calling it.
//
// The words at the end of a line are the whole of what a developer reads of this feature, and
// the layer that paints them cannot run outside monaco. These pin the spelling and the wire
// tolerance, so a change to either fails here in milliseconds rather than in Excel.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-blameformat-"));
const compiled = path.join(scratch, "blameformat.mjs");

await build({
  entryPoints: [path.join(root, "src", "blameformat.ts")],
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
});

const { blameDate, blameHint, blameHover, blameReadingOf, shortHashOf, UNCOMMITTED_HINT }
  = await import(pathToFileURL(compiled).href);

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

const ROW = {
  line: 4,
  hash: "5bbbd77c0ffee1234567890abcdef1234567890a",
  short: "5bbbd77",
  author: "William Smith",
  when: "2026-09-08T23:30:00+10:00",
  summary: "The record-cost bound measures the shape, not the runner",
};

check("the hint is author, date, short with two leading spaces", () => {
  assert.equal(blameHint(ROW), "  William Smith, 2026-09-08, 5bbbd77");
});

check("the date is the author's own day, not the reader's", () => {
  // 23:30 at +10:00 is the previous day in UTC; the hint says the day the author saw.
  assert.equal(blameDate("2026-09-08T23:30:00+10:00"), "2026-09-08");
  assert.equal(blameDate("2026-01-01T00:15:00-05:00"), "2026-01-01");
});

check("a stamp that cannot be read gives no date, and the hint leaves it out", () => {
  assert.equal(blameDate("yesterday"), "");
  assert.equal(blameDate(""), "");
  assert.equal(blameHint({ ...ROW, when: "" }), "  William Smith, 5bbbd77");
});

check("the short hash comes from the hash when the row carries none", () => {
  assert.equal(shortHashOf({ short: "", hash: "abcdef1234567890" }), "abcdef1");
  assert.equal(shortHashOf({ short: "  ab12cd3  ", hash: "zzz" }), "ab12cd3");
  assert.equal(shortHashOf({}), "");
});

check("an empty row hints nothing rather than a row of commas", () => {
  assert.equal(blameHint({ line: 1, hash: "", short: "", author: "", when: "", summary: "" }), "");
});

check("the hover carries the short hash, the author and stamp, then the summary", () => {
  assert.equal(
    blameHover(ROW),
    "**5bbbd77** William Smith, 2026-09-08T23:30:00+10:00\n\nThe record-cost bound measures the shape, not the runner");
  assert.equal(blameHover({ ...ROW, summary: "" }), "**5bbbd77** William Smith, 2026-09-08T23:30:00+10:00");
});

check("the uncommitted hint is a constant the css class pairs with", () => {
  assert.equal(UNCOMMITTED_HINT, "  not committed");
});

check("a reply is coerced: rows sorted by line, bad rows dropped, uncommitted deduplicated", () => {
  const reading = blameReadingOf({
    detail: "blamed",
    head: "5bbbd77",
    lines: [
      { line: 9, hash: "b", short: "bbbbbbb", author: "B", when: "2026-09-01T00:00:00Z", summary: "two" },
      { line: 2, hash: "a", short: "aaaaaaa", author: "A", when: "2026-08-01T00:00:00Z", summary: "one" },
      { line: 0, hash: "x", short: "xxxxxxx", author: "X", when: "", summary: "zero is not a line" },
      { line: "not a number" },
      null,
      "junk",
    ],
    uncommitted: [5, "6", 5, -1, "many", 3.5],
  });

  assert.equal(reading.head, "5bbbd77");
  assert.deepEqual(reading.lines.map((row) => row.line), [2, 9]);
  assert.deepEqual(reading.lines.map((row) => row.author), ["A", "B"]);
  assert.deepEqual(reading.uncommitted, [5, 6]);
});

check("a missing, errored or malformed reply reads as nothing committed", () => {
  assert.deepEqual(blameReadingOf(undefined), { head: "", lines: [], uncommitted: [] });
  assert.deepEqual(blameReadingOf(null), { head: "", lines: [], uncommitted: [] });
  assert.deepEqual(blameReadingOf({ error: "git was not found" }), { head: "", lines: [], uncommitted: [] });
  assert.deepEqual(blameReadingOf({ lines: "nope", uncommitted: 7 }), { head: "", lines: [], uncommitted: [] });
});

check("a row with fields of the wrong type keeps its line and blanks the rest", () => {
  const reading = blameReadingOf({ lines: [{ line: 3, hash: 12, author: null, when: {}, summary: [] }] });
  assert.deepEqual(reading.lines, [{ line: 3, hash: "", short: "", author: "", when: "", summary: "" }]);
});

let failures = 0;
for (const { name, fn } of checks) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}`);
    console.error(`     ${error.message}`);
  }
}

await rm(scratch, { recursive: true, force: true });

console.log(`${checks.length - failures}/${checks.length} passed`);
if (failures > 0) {
  process.exitCode = 1;
}
