// Format Module and the editor's outdent, tested by calling them.
//
// Upstream's analyzer 10.7.1 fixed these in ITS formatter and grammar, and none of that reaches
// this page: Format Module is format.ts and the outdent rule is vba.ts's, and neither imports the
// analyzer. Measured on 2026-09-23 with upstream's own cases, ours still closed no block at a
// one-word EndIf, formatted and respelled the line a comment carries on to through ` _`, and
// capitalized a variable named step. The VBE takes EndIf as End If and the carried line as comment
// text, so each case here is the VBE's reading, and each failed on the code before the fix.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-format-"));
const compiled = path.join(scratch, "format.mjs");

// format.ts imports vba.ts, which builds its language configuration from monaco's enums at load.
// Neither the formatter nor the outdent pattern needs monaco itself, so it is a stub that answers
// every property with another stub. The spec's helpers resolve the way the page's build resolves
// them: the vendored copy.
await build({
  stdin: {
    contents: [
      `export { formatVba } from ${JSON.stringify(path.join(root, "src", "format.ts"))};`,
      `export { vbaLanguageConfiguration } from ${JSON.stringify(path.join(root, "src", "vba.ts"))};`,
    ].join("\n"),
    resolveDir: root,
    sourcefile: "format-entry.ts",
    loader: "ts",
  },
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
  alias: { "xlide-spec": path.join(root, "vendor", "xlide-spec") },
  plugins: [{
    name: "stub-monaco",
    setup(on) {
      on.onResolve({ filter: /^monaco-editor\// }, (args) => ({ path: args.path, namespace: "stub" }));
      on.onLoad({ filter: /.*/, namespace: "stub" }, () => ({
        loader: "js",
        contents: [
          "const handler = { get: (t, k) => k === Symbol.toPrimitive ? () => 0 : stub(), apply: () => stub(), construct: () => stub() };",
          "function stub() { return new Proxy(function () {}, handler); }",
          "export const languages = stub(); export const editor = stub(); export const Range = stub();",
          "export const KeyCode = stub(); export const KeyMod = stub(); export const Uri = stub();",
        ].join("\n"),
      }));
    },
  }],
});

const { formatVba, vbaLanguageConfiguration } = await import(pathToFileURL(compiled).href);

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

/** Formats lines and answers lines, so a failure shows the module rather than one long string. */
const format = (...lines) => formatVba(lines.join("\r\n")).split("\r\n");

check("a block If closed with the one-word EndIf is closed (#88)", () => {
  assert.deepEqual(format(
    "Sub T(ByVal n As Long)",
    "If n > 1 Then",
    "Debug.Print n",
    "EndIf",
    "Debug.Print \"after\"",
    "End Sub",
  ), [
    "Sub T(ByVal n As Long)",
    "    If n > 1 Then",
    "        Debug.Print n",
    "    EndIf",
    "    Debug.Print \"after\"",
    "End Sub",
  ]);
});

check("#EndIf sits at the margin with the other directives", () => {
  assert.deepEqual(format(
    "Sub T()",
    "#If VBA7 Then",
    "x = 1",
    "#EndIf",
    "End Sub",
  ), [
    "Sub T()",
    "#If VBA7 Then",
    "    x = 1",
    "#EndIf",
    "End Sub",
  ]);
});

check("the line a comment carries on to through ' _' is left exactly as written (#82)", () => {
  assert.deepEqual(format(
    "Sub T()",
    "' note _",
    "  end sub",
    "Debug.Print 1",
    "End Sub",
  ), [
    "Sub T()",
    "    ' note _",
    "  end sub",
    "    Debug.Print 1",
    "End Sub",
  ]);
});

check("a carried comment runs on while each line ends in ' _', and an empty line ends it", () => {
  assert.deepEqual(format(
    "Sub T()",
    "' first _",
    "second _",
    "third",
    "x = 1",
    "' note _",
    "",
    "y = 2",
    "End Sub",
  ), [
    "Sub T()",
    "    ' first _",
    "second _",
    "third",
    "    x = 1",
    "    ' note _",
    "",
    "    y = 2",
    "End Sub",
  ]);
});

check("Rem carries the same way, and no Rem comment's words are respelled", () => {
  assert.deepEqual(format(
    "Sub T()",
    "Rem if this then that _",
    "end if",
    "x = 1: rem if not",
    "End Sub",
  ), [
    "Sub T()",
    "    Rem if this then that _",
    "end if",
    "    x = 1: rem if not",
    "End Sub",
  ]);
});

check("a comment on an #If line carries too", () => {
  assert.deepEqual(format(
    "Sub T()",
    "#If VBA7 Then ' note _",
    "not code",
    "#End If",
    "End Sub",
  ), [
    "Sub T()",
    "#If VBA7 Then ' note _",
    "not code",
    "#End If",
    "End Sub",
  ]);
});

check("a word that is a keyword only in its own statement keeps its spelling elsewhere (#86)", () => {
  assert.deepEqual(format(
    "Option explicit",
    "Private Declare ptrsafe Function Beep Lib \"kernel32\" (ByVal f As Long) As Long",
    "Private ptrsafe As Long",
    "Sub T()",
    "Dim step As Long, error As String, explicit As Long",
    "For i = 1 To 10 step 2",
    "Next",
    "On error GoTo Fail",
    "step = 2",
    "Debug.Print error, explicit",
    "Fail:",
    "End Sub",
  ), [
    "Option Explicit",
    "Private Declare PtrSafe Function Beep Lib \"kernel32\" (ByVal f As Long) As Long",
    "Private ptrsafe As Long",
    "Sub T()",
    "    Dim step As Long, error As String, explicit As Long",
    "    For i = 1 To 10 Step 2",
    "    Next",
    "    On Error GoTo Fail",
    "    step = 2",
    "    Debug.Print error, explicit",
    "Fail:",
    "End Sub",
  ]);
});

// Already right before the fix, pinned so the fix cannot take them away.
check("already right: an underscore with no space before it carries nothing", () => {
  assert.deepEqual(format(
    "Sub T()",
    "' note_",
    "if x then",
    "y = 1",
    "end if",
    "End Sub",
  ), [
    "Sub T()",
    "    ' note_",
    "    If x Then",
    "        y = 1",
    "    End If",
    "End Sub",
  ]);
});

check("already right: spaces after a continuation's _ continue it (#83), If x Then: is one line (#84)", () => {
  assert.deepEqual(format(
    "Sub T()",
    "If 1 = 1 _   ",
    "Then Debug.Print 1",
    "If x Then: Debug.Print 2",
    "Debug.Print 3",
    "End Sub",
  ), [
    "Sub T()",
    "    If 1 = 1 _",
    "        Then Debug.Print 1",
    "    If x Then: Debug.Print 2",
    "    Debug.Print 3",
    "End Sub",
  ]);
});

check("the editor outdents a line as it becomes EndIf, and End If as before", () => {
  const outdents = vbaLanguageConfiguration.indentationRules.decreaseIndentPattern;
  assert.equal(outdents.test("    EndIf"), true, "EndIf");
  assert.equal(outdents.test("    End If"), true, "End If");
  assert.equal(outdents.test("    #EndIf"), true, "#EndIf");
  assert.equal(outdents.test("    EndIfCount = 1"), false, "a name that begins with EndIf");
});

let failures = 0;
for (const { name, fn } of checks) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}`);
    console.error(`     ${error.message.split("\n").join("\n     ")}`);
  }
}

await rm(scratch, { recursive: true, force: true });

console.log(`${checks.length - failures}/${checks.length} passed`);
if (failures > 0) {
  process.exitCode = 1;
}
