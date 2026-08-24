// Recursion, of every shape the analyzer might follow forever.
//
// Depth is not the interesting axis - 1000 nested For loops parse in 100ms. A CYCLE is: a
// resolver that walks "what type does this return" has nowhere to stop when the answer is
// itself, and mutual recursion hides the cycle behind two names.

import { startEngine } from "./test/harness.mjs";

const { call, stop } = await startEngine("recurse");
const CRLF = "\r\n";

const cases = {
  "self-recursive function": [
    "Option Explicit", "",
    "Public Function Fact(ByVal n As Long) As Long",
    "    If n <= 1 Then Fact = 1 Else Fact = n * Fact(n - 1)",
    "End Function",
  ],

  "mutual recursion": [
    "Option Explicit", "",
    "Public Function IsEven(ByVal n As Long) As Boolean",
    "    If n = 0 Then IsEven = True Else IsEven = IsOdd(n - 1)",
    "End Function", "",
    "Public Function IsOdd(ByVal n As Long) As Boolean",
    "    If n = 0 Then IsOdd = False Else IsOdd = IsEven(n - 1)",
    "End Function",
  ],

  "a chain of 500 procedures": [
    "Option Explicit", "",
    ...Array.from({ length: 500 }, (_, i) =>
      [`Public Function F${i}() As Long`,
       `    F${i} = F${i + 1}()`,
       "End Function"]).flat(),
    "Public Function F500() As Long", "    F500 = 1", "End Function",
  ],

  "a class whose property returns itself": [
    "Option Explicit", "",
    "Public Property Get Self() As Node",
    "    Set Self = Me",
    "End Property", "",
    "Public Property Get Value() As Long",
    "    Value = 1",
    "End Property",
  ],

  "and a caller walking that cycle": [
    "Option Explicit", "",
    "Public Sub Walk()",
    "    Dim n As Node",
    `    Debug.Print n${".Self".repeat(200)}.Value`,
    "End Sub",
  ],

  "a type that contains itself": [
    "Option Explicit", "",
    "Public Type Cell",
    "    Next_ As Cell",
    "    Value As Long",
    "End Type",
  ],

  "two types containing each other": [
    "Option Explicit", "",
    "Public Type Left_", "    Other As Right_", "End Type", "",
    "Public Type Right_", "    Other As Left_", "End Type",
  ],
};

await call("initialize", {});

for (const [what, lines] of Object.entries(cases)) {
  const source = lines.join(CRLF);
  const id = what.replace(/[^a-z]/gi, "");
  const began = Date.now();
  try {
    await call("project/open", {
      projectId: id, generation: 1,
      modules: [
        { moduleName: "Deep", source, type: "standard" },
        // A class called Node exists, so the self-returning property resolves to something.
        { moduleName: "Node", type: "class", source: ["Option Explicit", "",
          "Public Property Get Self() As Node", "    Set Self = Me", "End Property", "",
          "Public Property Get Value() As Long", "    Value = 1", "End Property"].join(CRLF) },
      ],
    });
    const answer = await call("textDocument/diagnostics", {
      documentKey: `${id}/Deep`, projectId: id, generation: 1,
      source, moduleName: "Deep", moduleType: "standard",
    }, { timeout: 30000 });
    console.log(`${what.padEnd(34)} ${String(Date.now() - began).padStart(6)}ms  `
      + `${(answer.diagnostics ?? []).length} finding(s)`);
  } catch (err) {
    console.log(`${what.padEnd(34)} THREW after ${Date.now() - began}ms: `
      + String(err.message).slice(0, 80));
  }
}

// And completion at the end of that cycle, which is the resolver asked to walk it on purpose.
try {
  const source = ["Option Explicit", "", "Public Sub Walk()", "    Dim n As Node",
    `    n${".Self".repeat(200)}.`, "End Sub"].join(CRLF);
  const began = Date.now();
  await call("project/open", {
    projectId: "Complete", generation: 1,
    modules: [
      { moduleName: "Deep", source, type: "standard" },
      { moduleName: "Node", type: "class", source: ["Option Explicit", "",
        "Public Property Get Self() As Node", "    Set Self = Me", "End Property"].join(CRLF) },
    ],
  });
  const at = source.indexOf(".", source.indexOf("n.Self")) + `n${".Self".repeat(200)}.`.length;
  const answer = await call("textDocument/completion", {
    projectId: "Complete", generation: 1, source, moduleName: "Deep",
    moduleType: "standard", offset: source.lastIndexOf(".") + 1,
  }, { timeout: 30000 });
  console.log(`completion after 200 hops           ${String(Date.now() - began).padStart(6)}ms  `
    + `${(answer.items ?? []).length} item(s)`);
  void at;
} catch (err) {
  console.log(`completion after 200 hops           THREW: ${String(err.message).slice(0, 80)}`);
}

await stop();
