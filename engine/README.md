# Language engine

Serves VBA analysis to the add-in over a named pipe, in its own process.

## Why it is separate

Analysis of a large project takes seconds on a first pass. The editor is single threaded and owns
the thread the user types on, so that work cannot run inside it. A separate process cannot block the
host, cannot grow its memory, and cannot crash it.

## Why it reuses the editor extension's analyzer

The analyzer already exists, is validated against the real compiler, and meets per-keystroke latency
budgets on large modules. It is imported from the neighbouring `xlide_vscode` checkout rather than
copied, so both products share one implementation and improvements land in both. The build fails
with a clear message when that checkout is missing.

The analyzer was already written to run behind a worker thread's message port: a pure, synchronous
request handler with plain data crossing its boundary. A pipe is the same contract with a different
transport, so the handler is reused rather than reimplemented.

## Protocol

JSON-RPC 2.0, one object per line, over `\\.\pipe\<name>`. Newline framing rather than length
headers, because every message is a single line and a newline reader cannot desynchronise on a
miscounted byte.

Positions are UTF-16 character offsets into the module source, which is the analyzer's own currency.
The add-in converts between those and the editor's one-based line and column at its boundary.

| Method | Purpose |
| --- | --- |
| `initialize` | Handshake. Required before analysis. |
| `project/open` | Replaces everything known about a project: identity, generation, and every module's source. |
| `project/close` | Forgets a project. |
| `module/didClose` | Drops per-document incremental state. |
| `textDocument/diagnostics` | Analyses one module and returns its findings. |
| `shutdown` | Ends the session. |

A generation number accompanies a project and every analysis request. Analysis carrying a generation
the engine was not seeded with is refused with error `-32000` rather than answered from stale
sources, so findings can never describe text the user is not looking at.

## Build

```bash
npm install
npm run build      # bundle only
npm run package    # bundle, then one executable
npm test           # smoke test against the bundle
node test/smoke.mjs --exe   # smoke test against the executable
```

`npm run package` produces `dist/xlide-engine.exe`, which needs nothing installed on the machine it
runs on. It is large, because it contains a runtime as well as the analyzer.

## Status

Diagnostics work end to end and are covered by the smoke test. Completion, hover, and signature help
are resolvable by the analyzer but are not yet exposed as methods here.
