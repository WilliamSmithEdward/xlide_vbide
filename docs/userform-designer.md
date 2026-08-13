# The UserForm designer: where it stands, and how it lands

The designer is the one VBE document type this product has no story for yet, and the larger of
the two milestones status.md names. Three comments in `VbeMenus.cs` point at "backlog #14" for
it; the numbered list they referenced predates the repository's docs and no longer exists, so
this document is what #14 means now.

Written 2026-08-13, opening with what was measured that day rather than with the plan, because
the plan only makes sense against what already works.

## What is true today, measured

Probed against a live session (RenameFixture.xlsm, `component add kind=3` through the api):

- **The component plumbing is form-aware end to end.** A form is created and removed through
  the `component` route; `project()` reports it as `kind: "form", type: 3`; the tree draws it
  as its own `userform` row kind; `pane open` opens its code-behind; the surface holds the
  text; the analyzer, completion, rename and every other code feature apply to it exactly as
  to a class module. A form's CODE is a first-class citizen already.
- **The designer window exists and nothing can reach it.** The add created
  `RenameFixture.xlsm - EntryForm (UserForm)` (window type 1) with `visible: false`, and it
  stayed that way: no entry in `VbeCommands.ForName` summons it, and `VbeMenus` deliberately
  suppresses every entry point - View Object, the Toolbox (548), Tab Order (469), the Format
  menu's arrange group, Additional Controls - each annotated that it returns with the
  designer. The pane walk carries code panes only, so no tab ever appears for the design side.
- **Sync already meets the binary half.** Export writes the `.frm` with its `.frx` sidecar
  through the editor's own exporter; import updates a form's code and refuses to CREATE a form
  from source, answering `skipping-import` with the reason - a form's designer is not in the
  file. The refusal is the right behaviour and stays.

So the current contract is clean: the code half is fully served, the design half is
deliberately unreachable rather than half-served, and nothing on the surface lies about it.

## The decision already taken

Architecture section 6 chose the strategy and it holds: **our canvas, our inspector, every
mutation through the MSForms designer object model.** A form component exposes `Designer`, the
designer exposes controls, controls expose properties, and that model is what writes `.frm`
and `.frx` - so a form edited through it is byte-compatible with a form edited natively, and
stays editable by developers who do not have this tool.

Why the alternatives lose:

- **A hole in the cover for the native designer** puts a 1998 window inside the surface that
  exists to replace it, driven by entry points (project explorer double-click, the suppressed
  menus) that are hidden because they belong to the covered surface. Every form would pay the
  seam forever.
- **Writing `.frm`/`.frx` ourselves** makes this product a second authority on a binary format
  whose first authority is sitting in the process, already reachable, already the thing Excel
  trusts. [MS-OFORMS] is published, but a divergence between two writers corrupts workbooks;
  the object-model route makes byte-compatibility free.

Read-only-first is not a rival strategy; it is the first milestone of this one.

## What has to be proven, as experiments with instruments

The house method applies: the route ships before the feature, every claim is observable, and
the leak guarantee extends to every new wrapper on day one - a designer graph is COM all the
way down, and this codebase has paid four Excel crashes to learn what an undisposed wrapper
does.

1. **Reach and read.** From the shim: `VBComponent.Designer`, the controls, and per control
   its identity (Name, type/ProgID), geometry (Left/Top/Width/Height, in points), and the
   first ring of appearance (Caption, Font name/size/bold, ForeColor/BackColor, Enabled,
   Visible, TabIndex). Deliverable: a read route - `designer?module=` - answering the control
   tree as JSON, plus `New-FormFixture.ps1` building a workbook whose form holds one of every
   standard control. The route is the instrument every later spike reads through, and the
   com-leak suite grows its row the same day.
2. **Mutate and round-trip.** `Controls.Add` by ProgID, `Remove`, geometry and caption and
   name writes. The proof is the round trip: mutate, save, reopen, designer JSON identical -
   and sync-export before and after a read-only session, because the byte-diff of `.frm` and
   `.frx` is what answers whether reading the model dirties what it should not.
3. **Appearance fidelity.** What can be rendered honestly: fonts, OLE_COLOR to css, borders
   and special effects, and pictures - `IPictureDisp` extraction is the likely hard case. The
   standard is the architecture's own line: where a property is readable but not renderable,
   the canvas shows the control's real bounds and identity, never an approximation that lies.
4. **Containers.** Frame and MultiPage hold children, MultiPage holds Pages which hold
   controls, TabStrip is its own shape. Whether the controls collection is flat-with-parents
   or per-container decides the JSON tree's shape, and the tree must mirror the real
   hierarchy before anything renders it.
5. **Change detection.** Nothing announces a design change the way a code pane announces a
   revision - expect the publish tick's poll-and-change-key pattern again. The spike measures
   what a cheap change key costs on a form of fifty controls, against `publishUs`, before the
   tick carries it; hundreds of COM property reads per tick is exactly the class of cost C7
   and C8 taught us to measure rather than assume.
6. **Windows and modes.** Make a designer window visible from the shim and observe where it
   lands relative to the cover - today's probe only ever saw one born hidden. And the run
   story: `Show` a modal form and record what the surface does while the form owns the pump.
7. **Undo.** The expectation - to be verified, not assumed - is that the VBE's own Edit undo
   does not see object-model mutations. The architecture's answer is recorded property
   transactions in our layer; the spike proves the transaction log can invert every mutation
   class spike 2 established.

One more claim to convert from assumed to measured while at it: all of the above with "Trust
access to the VBA project object model" OFF, as everything else here runs. The shim's VBE root
comes from the add-in connection, not from `Application.VBE`, so the designer model should be
ungated the same way - one probe row makes it a fact.

## The page's half

The workspace assumes a tab is a monaco model: a group shows its active tab by swapping models
in its one editor. A design view is the first non-monaco tab content, and that is a real
piece of page architecture, not a widget.

Two shapes offer themselves: two tabs per form (code and design as separate identities), or
one tab with a code/design toggle. Lean toward **one tab, one identity, view mode as page
geography** - membership stays the host's (the component), while which face is showing is the
developer's, exactly the split the workspace already lives by. The host has no design pane in
its open list to echo, so a second identity would be page-invented and every host round trip
would have to special-case it. Decide finally at the canvas milestone, with the render spike
in hand.

The Properties panel is already the product's one property inspector, editing real component
state through the object model. Control selection extends that panel; a second inspector does
not grow.

And the loop that matters most day to day deserves its own name: **double-click a control,
get its event handler.** Stub the handler in the code-behind exactly as the native designer
would - same signature, same casing, same placement - and jump the caret there. It needs
spike 1 and the code surface that already works, nothing else; it may well ship before any
canvas does, and it alone converts forms from "edit the code blind" to "authorable".

## Milestones

- **M1 - observability.** The `designer` read route, the form fixture, the suites, and
  `capture` learning to shoot a designer window so parity claims have pictures. No UI.
  (Spikes 1, 4, 5, and the trust probe.)
- **M2 - the honest canvas.** Read-only design view in the workspace: real bounds, real
  captions, honest placeholders; selection; double-click writes the event stub. (Spikes 3
  and 6, and the one-tab decision.)
- **M3 - the inspector.** Selection flows into the Properties panel; property writes go
  through the model; the transaction log starts recording. (Spike 7.)
- **M4 - direct manipulation.** Move, resize, nudge, add from a toolbox, delete, align and
  distribute, undo over the transaction log. (Spike 2 in full.)
- **M5 - the finishing set.** Tab order, z-order, zoom, snapping and guides - and the
  suppressed menu entries return one by one, each unsuppressed in the same change that makes
  it true.

## Risks worth respecting

- **Pictures.** `IPictureDisp` to bytes may need an OLE round trip through a temporary
  stream; until it works, image-bearing controls render as honest bounds.
- **Third-party ActiveX.** Render bounds and identity, never guess an appearance; Additional
  Controls stays suppressed until add-by-ProgID is proven against at least one real
  third-party control.
- **COM volume.** Every property is a crossing; the designer's read path gets its own
  counter beside `publishUs` and `hostReadMs` before anyone asserts it is cheap.
- **The interop class of bug.** New interface, same rules: every wrapper counted, every
  release on the right thread, `com-leak.mjs` rows before features - the 16-byte VARIANT and
  the finalizer-thread FailFast were both found the expensive way.
- **Scope discipline.** The native designer keeps working untouched the whole time. Nothing
  unsuppresses, and no menu returns, until the thing behind it is true.
