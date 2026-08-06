// The page bundles a few smart-editing helpers that belong to the spec repo (xlide_vscode), so
// that typing behaves identically in both products. Those helpers used to be read straight out of
// a neighbouring checkout at build time, which meant the page could only be built on a machine
// that happened to have both repos side by side -- CI could not build it at all. They are now
// vendored into vendor/xlide-spec, and this script is what puts them there.
//
// Run it two ways:
//
//   node tools/sync-spec.mjs           check the vendored copy against the spec repo
//   node tools/sync-spec.mjs --write   re-copy from the spec repo and rewrite the manifest
//
// The check is what keeps the copy honest. It only means anything on a machine that has the spec
// repo, so when the neighbouring checkout is missing -- on CI, or on a clone of this repo alone --
// it reports that and succeeds. Drift is caught where it can be caught, and nowhere does the build
// depend on a directory outside this repository.

import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const pageRoot = path.resolve(here, "..");
const repoRoot = path.resolve(pageRoot, "../..");

// The same neighbouring-checkout convention the other cross-repo tool uses (tools/Compare-Lexers.ps1).
const specRepo = path.resolve(repoRoot, "..", "xlide_vscode");
const specRoot = path.join(specRepo, "src");

const vendorRoot = path.join(pageRoot, "vendor", "xlide-spec");
const manifestPath = path.join(pageRoot, "vendor", "xlide-spec.json");

const SPECIFIER = "xlide-spec";
const write = process.argv.includes("--write");

function fail(message) {
  console.error(message);
  process.exit(1);
}

// Line endings are git's business, not ours. The copy is written exactly as the spec repo holds it,
// but every comparison normalises first: git rewrites line endings on the way in and out of a
// checkout, so hashing raw bytes would make this check pass or fail on a clone's autocrlf setting
// rather than on whether the code is the same code. (Nothing in the vendored set spans a line
// inside a template literal, so normalising cannot change what the bundle does.)
function read(file) {
  return fs.readFileSync(file, "utf8").split("\r\n").join("\n");
}

function sha256(text) {
  return crypto.createHash("sha256").update(text).digest("hex");
}

function posix(p) {
  return p.split(path.sep).join("/");
}

function listFiles(root) {
  if (!fs.existsSync(root)) return [];
  const out = [];
  for (const entry of fs.readdirSync(root, { withFileTypes: true, recursive: true })) {
    if (!entry.isFile()) continue;
    out.push(posix(path.relative(root, path.join(entry.parentPath, entry.name))));
  }
  return out.sort();
}

const IMPORT_SOURCE = /from\s+["']([^"']+)["']/g;

// Which spec modules the page actually imports. Reading this out of the page's own sources rather
// than hard-coding it means a newly added spec import shows up as missing on the next check,
// instead of quietly building against a neighbouring checkout on one machine and failing on CI.
function entryPoints() {
  const entries = new Set();
  for (const rel of listFiles(path.join(pageRoot, "src"))) {
    if (!rel.endsWith(".ts") || rel.endsWith(".d.ts")) continue;
    const text = fs.readFileSync(path.join(pageRoot, "src", rel), "utf8");
    for (const [, source] of text.matchAll(IMPORT_SOURCE)) {
      if (source === SPECIFIER || source.startsWith(SPECIFIER + "/")) {
        entries.add(source.slice(SPECIFIER.length + 1));
      }
    }
  }
  return [...entries].sort();
}

// Everything those entry points reach, following relative imports. The helpers are plain
// TypeScript with no editor dependencies, so the closure closes; anything reaching outside it
// would be a real change in what the page depends on and is reported rather than copied.
function closureFrom(root, entries) {
  const seen = new Set();
  const outside = [];
  const missing = [];
  const queue = entries.map((e) => path.join(root, e + ".ts"));

  while (queue.length) {
    const file = queue.shift();
    const rel = posix(path.relative(root, file));
    if (rel.startsWith("..")) {
      outside.push(rel);
      continue;
    }
    if (seen.has(rel)) continue;
    if (!fs.existsSync(file)) {
      missing.push(rel);
      continue;
    }
    seen.add(rel);
    for (const [, source] of fs.readFileSync(file, "utf8").matchAll(IMPORT_SOURCE)) {
      if (!source.startsWith(".")) continue;
      queue.push(path.resolve(path.dirname(file), source + ".ts"));
    }
  }

  return { files: [...seen].sort(), outside, missing };
}

function specCommit() {
  try {
    return execFileSync("git", ["-C", specRepo, "rev-parse", "HEAD"], { encoding: "utf8" }).trim();
  } catch {
    return null;
  }
}

const entries = entryPoints();
if (entries.length === 0) fail("The page imports nothing from xlide-spec; refusing to guess what to vendor.");

const vendored = listFiles(vendorRoot);
const haveSpecRepo = fs.existsSync(specRoot);

if (write) {
  if (!haveSpecRepo) fail(`Cannot sync: the spec repo is not at ${specRepo}.`);

  const { files, outside, missing } = closureFrom(specRoot, entries);
  if (missing.length) fail(`The spec repo is missing:\n  ${missing.join("\n  ")}`);
  if (outside.length) fail(`These imports reach outside the spec sources:\n  ${outside.join("\n  ")}`);

  fs.rmSync(vendorRoot, { recursive: true, force: true });
  const manifest = { source: "xlide_vscode", commit: specCommit(), entries, files: {} };
  for (const rel of files) {
    const source = path.join(specRoot, rel);
    const target = path.join(vendorRoot, rel);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.copyFileSync(source, target);
    manifest.files[rel] = sha256(read(source));
  }
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + "\n");

  const lines = files.reduce((n, rel) => n + read(path.join(specRoot, rel)).split("\n").length, 0);
  console.log(`vendored ${files.length} files (${lines} lines) from ${manifest.commit?.slice(0, 7) ?? "an unknown commit"}`);
  process.exit(0);
}

if (vendored.length === 0) fail(`Nothing is vendored at ${posix(path.relative(repoRoot, vendorRoot))}. Run: npm run spec:sync`);

// The manifest is checked even without the spec repo, because a hand-edit of a vendored file is
// exactly the mistake this arrangement invites and it does not need a neighbouring checkout to spot.
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
const edited = [];
for (const rel of vendored) {
  const recorded = manifest.files[rel];
  const actual = sha256(read(path.join(vendorRoot, rel)));
  if (recorded !== actual) edited.push(rel);
}
const unlisted = vendored.filter((rel) => !(rel in manifest.files));
const absent = Object.keys(manifest.files).filter((rel) => !vendored.includes(rel));

if (edited.length || unlisted.length || absent.length) {
  const parts = [];
  if (edited.length) parts.push(`edited since it was vendored:\n  ${edited.join("\n  ")}`);
  if (unlisted.length) parts.push(`present but not in the manifest:\n  ${unlisted.join("\n  ")}`);
  if (absent.length) parts.push(`in the manifest but missing:\n  ${absent.join("\n  ")}`);
  fail(`The vendored copy does not match its manifest.\n${parts.join("\n")}\nRun: npm run spec:sync`);
}

if (!haveSpecRepo) {
  console.log(`spec repo not present at ${specRepo}; checked the manifest only (${vendored.length} files)`);
  process.exit(0);
}

const { files, outside, missing } = closureFrom(specRoot, entries);
if (missing.length) fail(`The spec repo is missing:\n  ${missing.join("\n  ")}\nThe page imports something the spec no longer provides.`);
if (outside.length) fail(`These imports reach outside the spec sources:\n  ${outside.join("\n  ")}`);

const stale = files.filter((rel) => {
  const mine = path.join(vendorRoot, rel);
  return !fs.existsSync(mine) || read(mine) !== read(path.join(specRoot, rel));
});
const extra = vendored.filter((rel) => !files.includes(rel));

if (stale.length || extra.length) {
  const parts = [];
  if (stale.length) parts.push(`changed in the spec repo:\n  ${stale.join("\n  ")}`);
  if (extra.length) parts.push(`vendored but no longer reached:\n  ${extra.join("\n  ")}`);
  fail(`The vendored copy has drifted from ${posix(path.relative(repoRoot, specRoot))}.\n${parts.join("\n")}\nRun: npm run spec:sync`);
}

console.log(`vendored copy matches the spec repo (${files.length} files, ${manifest.commit?.slice(0, 7) ?? "unknown commit"})`);
