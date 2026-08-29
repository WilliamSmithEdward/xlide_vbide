/*
 * What goes into a module is what the module holds. Character for character.
 *
 * THE DEFECT THIS EXISTS FOR. A whole-module write used `CodeModule.AddFromString`, and the
 * editor appends a line reading `()` to the bottom of any module containing a `Declare` broken
 * over a line continuation:
 *
 *   Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
 *       (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
 *
 * Not a rewrite of the statement - a new line of nonsense at the end of the module, which does
 * not compile. Reproduced outside this product entirely, in a throwaway workbook with no add-in
 * loaded, so it is the editor's doing and not this product's; what was this product's doing was
 * handing it that call and then accepting whatever came back. The owner pasted VBA-JSON, whose
 * Mac branch declares four of these, and got a module that would not compile from a file that
 * was fine (2026-08-21). Put the same declaration on one line and AddFromString is clean, which
 * is why nothing smaller than a real module found it.
 *
 * `InsertLines` reproduces the text exactly, so that is what a whole-module write uses now, and
 * the blank line the editor leaves after it is removed.
 *
 * OVER 400 LINES ON PURPOSE. Below that a write goes out as a line diff, which never used the
 * broken call - so a small module passes this whatever the product does, and would have passed
 * it on the day the owner's module broke.
 *
 * Run against any fixture with the editor open:
 *   node tools\harness\write-fidelity.mjs
 */

import { open, waitFor, scratchModule, comparingReporter } from "./xlide-api.mjs";

const api = await open();
const { check, done } = comparingReporter();

const project = (await api.projects()).projects[0];
console.log(`project: ${project.projectId}\n`);

// The declaration that does it, exactly as VBA-JSON writes it.
const CONTINUED_DECLARE = [
  "#If Mac Then",
  "",
  "#If VBA7 Then",
  'Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _',
  "    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr",
  "#Else",
  'Private Declare Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _',
  "    (ByVal utc_Command As String, ByVal utc_Mode As String) As Long",
  "#End If",
  "",
  "#End If",
].join("\r\n");

/** Padding, so the write is too big to go out as a line diff and takes the whole-module path. */
function padded(head, procedures) {
  const lines = ["Option Explicit", "", head, ""];
  for (let n = 0; n < procedures; n++) {
    lines.push(`Public Function Padding${n}(ByVal seed As Long) As Long`);
    lines.push(`    Padding${n} = seed * ${n + 1}`);
    lines.push("End Function");
    lines.push("");
  }
  return lines.join("\r\n");
}

const name = `WriteProbe${process.pid}`;
const scratch = scratchModule(api, project.projectId, name);

async function writeAndRead(what, text) {
  await api.writeModule(name, text, project.projectId);

  // The door gives up before a big write finishes, so wait for the text rather than the reply.
  const held = await waitFor(
    `${what}: the module to hold the write`,
    async () => {
      const now = (await api.readModule(name, project.projectId)).text;
      return now !== undefined && now !== null && now.length > 0 ? { text: now } : false;
    },
    { budgetMs: 30000, pollMs: 200 });

  return held.text;
}

try {
  await api.component("add", { name, kind: "module", project: project.projectId });

  // 1. THE DEFECT ITSELF.
  const withDeclare = padded(CONTINUED_DECLARE, 120);
  console.log(`the module under test is ${withDeclare.split("\r\n").length} lines\n`);

  const back = await writeAndRead("a continued Declare", withDeclare);

  check("a module holding a continued Declare comes back with the same number of lines",
    back.split("\r\n").length, withDeclare.split("\r\n").length);

  check("no line was appended to the end of it",
    back.split("\r\n").at(-1), withDeclare.split("\r\n").at(-1));

  check("nothing anywhere in it reads '()'",
    back.split("\r\n").filter((line) => line.trim() === "()").length, 0);

  check("and it holds exactly what was written", back === withDeclare);

  // 2. THE SAME DECLARATION ON ONE LINE, which the editor was always happy with. Here so a
  //    failure above can be read as "the continuation" rather than "Declare statements".
  const oneLine = padded(
    'Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" '
    + "(ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr",
    120);

  check("the same declaration on one line holds exactly too",
    (await writeAndRead("a one-line Declare", oneLine)) === oneLine);

  // 3. A TRAILING BLANK LINE IS THE DEVELOPER'S, and survives. The fix removes what the editor
  //    adds past the text; it must not take a blank line that was written on purpose.
  const trailing = `${padded(CONTINUED_DECLARE, 120)}\r\n`;
  const trailingBack = await writeAndRead("a trailing blank line", trailing);

  check("a blank line written at the end is still there", trailingBack === trailing);

  // 4. AND THE SMALL PATH still round-trips, so the two ways of writing agree.
  const small = ["Option Explicit", "", CONTINUED_DECLARE, "", "Public Sub Small()", "End Sub"].join("\r\n");
  check("a module small enough to go out as a line diff holds exactly too",
    (await writeAndRead("a small module", small)) === small);
} finally {
  await scratch.dispose();
}

process.exit(done());
