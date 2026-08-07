# Parity audit: what is wired, what is not

Measured 2026-08-06 against the shipping build, as step 1 of the parity goal in
[handoff.md](handoff.md) section 9. Four layers have to line up for a capability to reach a
developer: the analyzer has to compute it, the engine has to expose an operation, the shim has to
route it, and the page has to register a provider. A capability missing any one of them is
invisible, and the analyzer is almost never the missing one.

## The table

| Capability | Analyzer | Engine op | Shim route | Page provider | State |
| --- | --- | --- | --- | --- | --- |
Three of the five have been wired since. The table below is as of 2026-08-06; what changed is
recorded under it.

| Capability | Analyzer | Engine op | Shim route | Page provider | State |
| --- | --- | --- | --- | --- | --- |
| Quick fixes | yes | `codeAction` | yes | `registerCodeActionProvider` | shipped |
| Semantic highlighting | yes | `semanticTokens` | yes | `registerDocumentSemanticTokensProvider` | shipped |
| Go to definition | yes | `definition` | yes | `registerDefinitionProvider` | shipped |
| Find references | yes | `references` | yes | `registerReferenceProvider` | shipped, one limit |
| Rename symbol | yes | no | no | no | analyzer only |

Find references answers across the whole workbook, but the editor's references window can only
render modules that have a tab open: a module with no tab has no model to render. Go to definition
has no such limit — a definition in an unopened module goes through the host, which opens it on the
way. What would close the gap is a way for the page to ask the host to open a module without also
moving the caret into it.

## The table as first measured

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

## What the first four taught

Three registrations turned out to cover less than their names: `codeAction`, `semanticTokens` and
`gotoSymbol` each pull in one contribution and leave the rest in modules they never import. In
every case the provider registered cleanly and was then asked nothing, which looks exactly like a
provider that answers nothing. The way to tell them apart is to watch a running editor and see
whether the request goes out at all. Expect this of the next feature rather than rediscovering it.

Two more editor facts, found the same way. A code-action provider must declare
`providedCodeActionKinds`, because the editor gates Ctrl+. and Shift+Alt+. on a context key built
from that list. And semantic highlighting defaults to whatever the theme says, which for a
standalone theme is always no — the flag is hardcoded off on every one of them, so the editor has
to be asked outright.

## Rename, as measured live

Proven in a real host (2026-08-06), reading the result out of the VBA project rather than
believing the editor's report: a rename from a module with a tab open rewrote a module with NO tab
open — both its bare call and its qualified one — because the write goes through the object model,
which does not care what is showing.

| Site | Renamed |
| --- | --- |
| the declaration | yes |
| a bare call inside the declaring module | yes |
| a qualified `Module.Sub` call, anywhere | yes |
| a bare call elsewhere, one definition in the workbook | yes |
| a bare call elsewhere, colliding definitions | no, by the resolver |
| the other module's own declaration and its own calls | never |

**One thing is NOT proven.** The collision row holds at the engine layer, where seeding is explicit
and certain. The live attempt at it was inconclusive: the colliding declaration was added through
the object model rather than through the editor, and there is no way from outside to confirm the
engine had re-seeded with it before the rename ran — so the bare call being renamed may be correct
behaviour against a project the engine still saw as having one definition, or may be the
divergence it looks like. Settle it by making the colliding declaration THROUGH the editor, where
the write-back and reseed are observable, before trusting the row.

The ambiguity report is also over-broad: it currently lists the other procedure's own declaration
and its own calls, which are a different symbol and not something a developer needs to decide
about. Only the genuinely ambiguous bare call belongs in a warning.

## What rename still needs

The engine part is small: `referencesFor` already returns every span, and a rename is those spans
with a validated new name. What decides whether the feature is correct is the part that is not
plumbing.

A rename that stops at a module boundary compiles until the module nobody renamed runs, so a
rename that cannot reach every module must refuse rather than do most of it. Modules with a tab
open are straightforward — the edits go through the same path typing already uses, so the unsaved
dot, Don't Save, and the baselines behind both keep working because nothing new is happening to
them. Modules with no tab are the whole question: writing them behind the developer's back means
inventing baselines for text they never saw, and refusing until they open every affected module
makes the feature rarely usable. Deciding that, and proving it against a live host, is the work.

Undo is the other one. The editor's undo stack is per model and the edit is not, so a multi-module
rename undoes a tab at a time unless something is built to hold it together.
