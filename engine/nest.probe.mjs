// How deep can a nested structure go before the analyzer stops coping?
//
// A recursive-descent parser meets a stack it did not choose. Run headless, against the engine,
// because a stack overflow in .NET or in node is not an exception you catch - it ends the process.

import { reporter, startEngine } from "./test/harness.mjs";

const { call, stop } = await startEngine("nest");
const CRLF = "\r\n";

const nested = (depth, kind) => {
  const open = [];
  const close = [];
  for (let i = 0; i < depth; i++) {
    if (kind === "for") {
      open.push(`${"  ".repeat(i + 1)}For i${i} = 1 To 2`);
      close.unshift(`${"  ".repeat(i + 1)}Next i${i}`);
    } else if (kind === "if") {
      open.push(`${"  ".repeat(i + 1)}If x = ${i} Then`);
      close.unshift(`${"  ".repeat(i + 1)}End If`);
    } else if (kind === "with") {
      open.push(`${"  ".repeat(i + 1)}With Application`);
      close.unshift(`${"  ".repeat(i + 1)}End With`);
    } else {
      open.push(`${"  ".repeat(i + 1)}Do While x < ${i}`);
      close.unshift(`${"  ".repeat(i + 1)}Loop`);
    }
  }
  return ["Option Explicit", "", "Public Sub Deep()", "    Dim x As Long",
    ...Array.from({ length: depth }, (_, i) => `    Dim i${i} As Long`),
    ...open, "        x = 1", ...close, "End Sub"].join(CRLF);
};

await call("initialize", {});

for (const kind of ["for", "if", "with", "do"]) {
  for (const depth of [10, 50, 200, 500, 1000]) {
    const source = nested(depth, kind);
    const began = Date.now();
    try {
      await call("project/open", {
        projectId: `${kind}${depth}`, generation: 1,
        modules: [{ moduleName: "Deep", source, type: "standard" }],
      });
      const answer = await call("textDocument/diagnostics", {
        documentKey: `${kind}${depth}/Deep`, projectId: `${kind}${depth}`, generation: 1,
        source, moduleName: "Deep", moduleType: "standard",
      });
      const took = Date.now() - began;
      console.log(`${kind.padEnd(5)} depth ${String(depth).padStart(4)}  ${String(took).padStart(6)}ms  `
        + `${(answer.diagnostics ?? []).length} finding(s)`);
    } catch (err) {
      console.log(`${kind.padEnd(5)} depth ${String(depth).padStart(4)}  THREW after `
        + `${Date.now() - began}ms: ${String(err.message).slice(0, 90)}`);
      break;
    }
  }
}

await stop();
