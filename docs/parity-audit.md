# Parity audit: what is wired, what is not

Measured 2026-08-06 against the shipping build, as step 1 of the parity goal in
[handoff.md](handoff.md) section 9. Four layers have to line up for a capability to reach a
developer: the analyzer has to compute it, the engine has to expose an operation, the shim has to
route it, and the page has to register a provider. A capability missing any one of them is
invisible, and the analyzer is almost never the missing one.

## The table

| Capability | Analyzer | Engine op | Shim route | Page provider | State |
| --- | --- | --- | --- | --- | --- |
| Diagnostics | yes | `analyze` | yes | markers | shipped |
| Completion | yes | `completion` | yes | `registerCompletionItemProvider` | shipped |
| Hover | yes | `hover` | yes | `registerHoverProvider` | shipped |
| Signature help | yes | `signature` | yes | `registerSignatureHelpProvider` | shipped |
| Outline | yes | `outline` | yes | explorer and navigation | shipped |
| Typing, formatting | yes | `onType` | yes | two formatting providers | shipped |
| Search | yes | `search` | yes | the search widget | shipped |
| **Quick fixes** | `analyzer/codeActions`, 632 lines | no | no | no | analyzer only |
| **Semantic highlighting** | `analyzer/semantic`, 582 lines | no | no | no | analyzer only |
| **Go to definition** | `analyzer/symbols`, 2,546 lines | no | no | no | analyzer only |
| **Find references** | `analyzer/symbols` | no | no | no | analyzer only |
| **Rename symbol** | `analyzer/symbols` | no | no | no | analyzer only |

The engine exposes eight operations in total (`initialize`, `analyze`, `seed`, `forget`, `module`,
`object`, `all`, `shutdown`) alongside the modules above. The page registers five Monaco providers.

## What the table says

Every remaining parity item is analysis that already ships and cannot be reached. `codeActions` and
`semantic` are compiled into the engine running on this machine right now, answering nothing,
because no operation names them. The symbol index that would answer definition, references and
rename is the same one already backing the outline.

So the work is three layers of plumbing per item, none of it novel: the seven shipped capabilities
are seven worked examples of exactly the same shape.

## The shape of each addition

An engine operation in `engine/src`, following `hover.ts` or `signature.ts`, which are the smallest
of the existing ones. A request and reply pair in the protocol. A route in the shim beside the
others in `AddInSession`, with its message records in `EditorMessages.cs` and the JSON context
entry that ahead-of-time serialisation requires. A message type on the page's bridge union, and a
Monaco provider registration.

Monaco has a first-class provider for every one of these:
`registerCodeActionProvider`, `registerDocumentSemanticTokensProvider`,
`registerDefinitionProvider`, `registerReferenceProvider`, `registerRenameProvider`.

## What is not just plumbing

**Rename.** Everything above answers questions; rename changes code, in modules that may be open,
dirty, or closed. It goes through the host's writer, and it has to agree with the baselines the
unsaved dot and Don't Save both read, or a rename will either lose the dot or make Don't Save
revert to the wrong text. Undo across several modules at once is the part to think hardest about,
since Monaco's undo stack is per model and the edit is not.

**The workbook boundary**, for both rename and references. Two open workbooks can each hold a
`Module1` and a `Recalculate`, and they are unrelated. Everything here is addressed as
(workbook, module) already; reuse that rather than inventing a scope.

**Right-click.** Go to definition and find references belong on the editor's context menu beside
the run and breakpoint commands, not only on F12. That menu already exists and is curated per
object class.

## Suggested order

Quick fixes, then semantic highlighting: both are pure plumbing, both are highly visible, and they
prove the pattern before anything harder. Then definition and references, which build the index
plumbing rename needs. Then rename. Then measure whether the 15,000-line generated Excel object
model in `analyzer/host` actually reaches completions here the way it does in the extension, which
is the largest remaining IntelliSense question and is data plumbing rather than logic.
