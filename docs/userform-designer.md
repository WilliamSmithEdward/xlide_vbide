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

## The form as text: the markup layer

Decided 2026-08-13, at the owner's direction, and it reshapes the milestone order below: a
markup language of xlide's own that DESCRIBES a form - generate it from an existing design,
regenerate a design from it - sitting strictly ABOVE the real form as a higher layer of
abstraction. Nothing about it may disturb a developer without xlide: the MSForms model keeps
writing `.frm` and `.frx`, the workbook stores nothing new, and the markup exists only as a
live projection this product computes and applies.

Why it comes before the canvas: the workspace already IS a text editor. A form projected into
text is immediately editable with everything monaco has - diff, undo, search, copy a Frame
between forms - which makes the markup the first working design surface, with zero rendering
fidelity at stake. The canvas, when it lands, becomes a RENDERER of the same projection rather
than a rival editing model.

### The shape

```
' EntryForm, as xlide projects it. Edits here apply back through the designer model.
Form EntryForm "Quarter Entry" size 360x320
    Label NameLabel "Customer" at 12,14 size 66x16
    TextBox NameBox at 84,12 size 120x20
    Frame Options "Freight" at 12,112 size 92x66
        OptionButton PickGround "Ground" at 8,14 size 76x16
        OptionButton PickAir "Air" at 8,34 size 76x16
    MultiPage Wizard at 12,188 size 192x86
        Page Page1 "Page1"
            CheckBox Agree "Agreed" at 8,8 size 100x16
        Page Page2 "Page2"
    CommandButton OkButton "Start" at 262,250 size 72x24
        Font.Size = 12
        Font.Bold = True
```

The dialect is VBA's, deliberately: apostrophe comments, `True`/`False`, doubled quotes inside
strings, `&H8000000F&` where a colour is spelled. A header line is
`Type Name ["Caption"] [at left,top] [size width x height]` - identity and the universals
inline, because every visual control has them - and anything else is an indented
`Path = value` line, one level of dotting reaching a font. Indentation is containment: a
Frame's children sit under it, a MultiPage's Pages under it, a Page's controls under the Page.
Geometry is points, as the designer measures.

### The rules that make it honest

- **Projection, not source.** Generated FROM the designer walk, applied TO the model through
  the M1 primitives. The model writes the binary; byte-compatibility is inherited, not
  re-earned.
- **The control list is total; the property list is not.** Every control appears, so a line
  deleted is a control removed and a line added is a control created - that is what makes the
  document a design surface rather than a patch. Properties are the opposite: an unspoken
  property is NEVER touched on apply, so the narrow printed vocabulary cannot erase state it
  does not represent.
- **Apply is a diff, keyed by name.** Match by control name: present-only-in-markup adds,
  present-only-in-model removes, matched controls take their header geometry and property
  lines through `set`. A changed type or container is a remove-and-add, which is also the
  truth of what it means. Renaming in markup is therefore remove-plus-add of a fresh control -
  the refactoring rename stays the designer's own gesture.
- **Third-party controls stay honest.** They print under their raw type name; applying can
  move and set them, but creating one needs a `ProgId = "..."` line, because a name the
  toolbox table does not know cannot be conjured.
- **Deterministic printing.** Model order, canonical formatting, so a regeneration diffs
  cleanly against the last one and against source control if the markup is ever exported.
- **A parse error applies nothing.** The whole document parses first, errors carry line
  numbers and names, and a failed apply reports exactly which operations landed before the
  refusal - the M1 primitives' own add-or-nothing guarantees bound the blast radius until the
  transaction log (the inspector milestone) makes apply atomic.

### Where it lives for the developer

Expanding a form in the tree stops showing an empty row: it shows the CODE-BEHIND and the
FORM MARKUP as its two faces, each opening as a tab. The markup document is a page-side
monaco document with its own small language (tokens, folding, later completion for control
kinds and properties); its text is computed by the shim and refreshed the way module text is,
with the developer's unapplied edits holding the document the way unwritten code does. The
membership stays the shim's - a virtual document, since no native pane exists behind it - so
the host-owns-membership invariant survives.

## Milestones

- **M1 - observability.** The `designer` read route, the form fixture, the suites, and
  `capture` learning to shoot a designer window so parity claims have pictures. No UI.
  (Spikes 1, 4, 5, and the trust probe.)

  **Landed 2026-08-13, the day this document was written.** The `designer` route reads a
  form whole and mutates it (add by ProgID, remove through the owning collection, set with
  read-back); `New-FormFixture.ps1` builds FormFixture.xlsm - every standard control, a Frame
  with children, a control on a MultiPage page - from `form-plan.mjs`, the one declaration
  `designer-features.mjs` then verifies a read against, 87 checks green; the com-leak sweep
  grew a build-read-remove row; and the whole of it ran with project trust OFF, which
  converts that claim from assumed to measured. Spike 1 is answered and spike 4 is MOOT
  rather than answered: the walk recurses into containers and dedupes by name, so it is
  correct whether the runtime's collections are flat or hierarchical, and nothing yet needed
  to know which. Facts paid for on the way, kept here because they are documented nowhere
  else:

  - The extenders name themselves by INTERNAL INTERFACE - `IMdcText`, `IOptionFrame`,
    `ILabelControl` - and the route maps the fifteen standard ones to their toolbox names. A
    name outside the table passes through untouched.
  - `Controls.Add` produces controls with EMPTY captions; "Label1" is the native toolbox
    GESTURE's doing, not the model's.
  - The form's own Width and Height are not on the designer object. They live on the
    component's Properties collection - the native Properties window's source - and the
    route reads and writes them there.
  - A font's Size is VT_CY on the wire, and both generic variant readers in DispatchObject
    lacked the case: it read as null and printed as the literal text "VT_CY" until the raw
    scaled-long representation was handled.
  - Capture of a designer window is NOT done: the window is born hidden and spike 6 owns
    making one visible at all.

  **The language features closed the loop 2026-08-13, same day, across three upstream
  rounds** (xlide_vscode v3.6.0 and the two follow-ups, all wired here the hour they landed):
  a form's controls ride every seed as implicit members - name plus `MSForms.<kind>` - so
  completion answers a control's own type (per-control, not shared), diagnostics resolve the
  controls where the finding is made (the one-day shim-side filter retired with a note about
  what it over-swallowed), hover answers the control and its members with the MSForms
  signature, the call tip carries parameters, and `Me.` composes controls plus code plus the
  form surface.

  **The fourth round closed the colouring, 2026-08-13 evening** (xlide_vscode#20, landed
  v3.8.0 with a same-day audit fixing `Me.Hide`-without-meType and the With-block leading
  dot): a RESOLVED method call on a control paints `function` - the same `dcdcaa` yellow
  `Len` gets - while a property read and an unresolved member stay untouched. Three seams
  carried it and each is pinned at its own level: the engine spreads the analyzer's third
  collector (`engine/test/forms.mjs` holds the acceptance table headlessly, a gate and CI
  step so the unpinned analyzer checkout cannot move under it silently), the page's semantic
  legend gained `function` APPENDED so existing token indices keep their meaning (the page
  smoke test now asserts the legend and the theme rule together, because the provider drops
  unknown types silently), and the live paint is measured, not assumed: `AddItem`, `SetFocus`
  and `Hide` render mtk13 `rgb(220,220,170)`, byte-identical to `Len`, while `Value` stays
  identifier blue. The same drop merged `MSForms.Control` - the base every placed control
  extends - into every control class, so `NameBox.` now completes and hovers `SetFocus`,
  `Left`, `Visible` (in no per-control dump; three suite rows pin it), and the `_`-prefixed
  dispatch internals are filtered the way the VBE's own list hides them (`Me.` reads 70
  members now, not 85). Two more paid-for facts:

  - A form NAME burns for the whole session once used - added, removed, or refused - and a
    workbook that loads holding a name burns it the moment that form is removed. Fixture
    machinery uses fresh names per attempt and never removes a form it did not create.
  - Touching `VBComponent.Designer` MATERIALISES the designer window, a save while one
    exists makes the workbook RESTORE it on open, and the native Toolbox floats up beside
    any live designer. The walks put the window back down (`KeepDesignerDown`), which is
    also what keeps the Toolbox away until the canvas milestone wants it.
- **M2 - the form as text.** The markup language (model, parser, printer in Core, unit-tested
  without Excel), the generate and apply routes over the M1 primitives, and then the
  developer-facing half: the tree's form row expanding to code-behind and markup, the markup
  tab with its own language, edits applying back. The first working design surface.

  **The language and the routes landed 2026-08-13, the same day as M1.** `FormMarkup` in Core
  parses and prints the dialect with 20 unit tests behind it; `designer?format=markup`
  projects the live walk; `applyMarkup` diffs an edited document back - proven live by the
  suite: idempotent on the form's own text, a line added materialises a placed captioned
  control, re-applying the original removes it, and a document that does not parse applies
  nothing with its line named.

  **The designer TAB landed 2026-08-13, late evening, to the owner's shape** (his words, four
  messages: the markup lives inside the editor space, not a dock pane; a form opens a special
  editor tab that splits into the markup document and the visual representation, both visible
  at once; edits to the markup update the form; and it plays nice with the existing tab
  setup). What stands:

  - **The tab is host-tracked, not page-faked.** The pane list the strip mirrors gained a
    FACE: a code pane is the host's own window mirrored, a `design` tab is product state the
    shim carries (`_designerTabs`), because backing it with a native designer window would
    summon the Toolbox - `KeepDesignerDown` is the standing contract. That is the one
    designated exception to "the strip mirrors the host", documented at the field. It
    survives page reloads, publishes through `setModules` (`faces`, `activeFace`), and the
    api drives it with `pane?action=open&face=design` - the same method the tree's Open
    Designer item calls, so the api leaves the click's state (the mirror rule).
  - **One walk feeds both halves.** `formMarkup` carries the markup TEXT and the SPEC it was
    printed from (`FormDesignService.SpecOf`, one designer walk), so the document and the
    visual cannot describe two moments of the form. The visual is the honest canvas opened
    early: points scaled 4/3, children composed by NESTING inside their container's element
    (MSForms' own coordinate model), captions real, a foreign type drawn as bounds plus its
    name, never a guessed appearance. Page tabs render as headers with the first page's
    content showing.
  - **The page half**: `DocumentId` grew an optional face (two tabs, one module, distinct
    keys); the workspace mounts a per-form `DesignerView` over the group's editor and
    reparents it as the tab moves, so the markup editor's scroll and undo ride along; the
    view disposes when the host's list drops its tab - a move between groups never passes
    through that list, so a move cannot be mistaken for a close. The tree's form row offers
    Open Designer from its menu (opening the designer has no shorter gesture, which is this
    menu's own admission rule).
  - **The apply loop closed the same evening**: Ctrl+S in the markup half posts the document,
    the shim applies it through `FormDesignService.ApplyMarkup` - the machinery MOVED OUT of
    the `#if DEBUG` partial to product side, which is the unification the service header had
    promised, so the api's `applyMarkup` route and the tab's Ctrl+S are one operation by
    construction (the route is a wrapper now) - and the outcome comes back first
    (`formMarkupApplied`), then a fresh projection re-renders the canvas. The document
    adopts the canonical print as ONE UNDOABLE EDIT when the developer holds no unapplied
    edits; a refusal lands in a strip under the document with the host's own wording (the
    parse line, or "what landed first" for a stop partway) while the canvas shows what
    actually landed. The page owns the design tab's dirty dot - unapplied markup is page
    state, and the host's echo is barred from blinking it off. `act designerApply` and
    `act designerMarkup` drive the tab's own path from suites.

  Suite: 113 checks - the tab standing active and labelled `[Design]`, its close taking only
  the design face, a non-form refused, and the tab's own apply landing a control on the real
  form (independently confirmed by the designer read), removing it on re-apply, and refusing
  a bad document by line while touching nothing. The paint was verified by capture: markup
  document beside the rendered form, Frame and MultiPage nesting correct.

  **The markup language service, the owner's ask (2026-08-13 evening), in landing order:**

  1. *Highlighting and auto-indent* - a Monarch grammar on the page, standard token names so
     the existing themes colour it without edits, and enter-rules that indent under a
     container line. Page-local, no host involvement.
  2. *Linting, red squiggles* - **landed 2026-08-13, late**: `FormMarkup.Lint` is a
     TOLERANT PARSE with the strict parser as the one grammar - each refusal's line is
     blanked and the parse re-run, so the lint can never disagree with Parse, only continue
     past it - plus the semantic rows an apply would note and skip (a duplicate name at its
     second mention, a foreign type with no ProgId line; a stray Page arrives as the
     parser's own error). Six Core tests hold the collector; the wire is pure text both
     ways (`lintFormMarkup`/`formMarkupLint`, no designer touched, no window stirred); the
     view lints debounced as the developer types and draws the markers on the document.
     `act designerSetMarkup` and `act designerLint` drive it; suite rows pin the squiggles
     appearing at their lines, warnings apart from errors, and clearing on canonical text.
     The same evening the language moved to FOUR-space indentation (the owner's call):
     printer, parser, editor tab size, every literal and sample.
  3. *Completions and hover* - control types, `at`/`size` scaffolding, known property
     paths per control kind; needs the language service shape of (2) to answer from.

  **Form properties joined the markup 2026-08-13, late** (the owner's ask): the projection
  prints the form's own property lines - `BackColor`, `ForeColor` to start - read from the
  SAME source the native Properties window edits, which is what links the document to that
  panel: same property bag, same value. Printed only when NOT the default, because the
  dialect's standing rule is that an unspoken property is one an apply can never erase - a
  document without the line leaves a custom colour standing (suite-pinned from both sides).
  The projection is ONE path now: the route, the tab and the apply all print through
  `FormDesignService`, the route's own copy retired the day it drifted. And the first
  LIVENESS hooks landed with it: every designer mutation route re-projects an open tab, so
  an api `set` reaches the tab's document without re-activation - unless the developer
  holds unapplied edits, which a push never clobbers.

  **The canvas follows the typing - the draft preview, landed 2026-08-14** (the owner:
  "if I update in the markdown pane, it doesn't reflect in the xlide form designer"). The
  debounced lint round trip - already parsing every keystroke's text host-side - carries
  the parsed spec back beside the squiggles, and the canvas renders THAT while the
  document is dirty: dialect fields from the draft, display extras (fonts, colours,
  insets, tabs) worn from the last applied projection by name so the preview stays
  dressed, the unspoken-colour rule applied exactly as an apply would. A draft that stops
  parsing keeps the last good picture rather than blanking under a half-typed line; a
  document back at canonical puts the form's own picture back; a note strip and a dashed
  outline say DRAFT out loud. The form is untouched throughout - Ctrl+S remains the only
  apply - and there is still exactly ONE parser: the strict grammar, host-side, the
  apply's own. `designerCanvas` reads what the canvas shows, draft flag and placed
  controls both.

  **Ctrl+S is the product's save, and the canvas scrolls - 2026-08-15, the owner's pair of
  reports, and the key's REAL route took a second report to find.** Ctrl+S is a HOST
  accelerator: the editor takes the key before the page ever sees it and routes it through
  the session's Save - which is why the code editor's Ctrl+S works (the host flushes page
  edits, then saves) and why a page-side binding alone read as "still not working" from
  both halves. The designer now has the same shape the code editor has: the host's Save,
  finding a designer tab active, asks the PAGE to apply the tab's document - the document
  lives there, the host cannot flush it itself - and the page calls back for the raw save
  ("saveOnly", which skips the designer branch so the callback cannot loop). An OK apply
  is followed by the save; a refusal saves nothing (the file must not hold a form the
  developer was just told did not take their document); a clean document skips straight to
  the save. The page-side bindings stay as belts for focus states the accelerator misses.
  The canvas wheel is unconditional on the canvas half (which is also what makes it
  drivable: a synthesised WheelEvent never triggers native scrolling, so the scroll path
  had no driver before), and a click gives the canvas focus so it scrolls by keyboard too.
  A draft picture said so in a banner across the top of the canvas until 2026-08-15, when
  the owner retired it: the tab's unsaved dot already says it, the amber outline around the
  form says it in place, and a strip of canvas is a lot of room for a third voice.

  **The parity build, designed and next** (the owner's bar, 2026-08-13: what a user WITHOUT
  xlide sees on the real form surface is the truth the canvas answers to):

  - *Geometry by the model's own numbers* - **landed 2026-08-13, late**: the walk carries
    `InsideWidth`/`InsideHeight` and the canvas derives Frame and MultiPage client areas
    from them (side borders split the width difference, the remainder above is the caption
    strip); the guessed constants stay only as fallbacks for a model that will not answer.
    Measured immediately: the frame's real inset is 7.5px where the guess said 13.3.
  - *Appearance* - **landed with it**: per-control Font (name, size, bold, italic) and
    BackColor/ForeColor plus the form's own ride the same walk, OLE colours converted
    host-side through the live system palette (`GetSysColor`), so the canvas paints what
    THIS machine's real surface would - the form face went from the hardcoded classic-grey
    guess to the machine's actual `#f0f0f0` the moment the palette call landed. A Label's
    BackColor is deliberately NOT painted until the walk reads BackStyle: its default is
    transparent, and colouring what the real surface leaves clear is the wrong direction
    to approximate.
  - *Parity round 1, from the owner's side-by-side* - **landed 2026-08-13, late** (the
    owner ran the fixture form beside the canvas and named the gaps): the form's chrome is
    DERIVED now like the Frame's - the title bar sits inside the form rect and eats exactly
    `(Height - InsideHeight)` less the border split, where before it floated above and the
    full outer rect stood in for the client, sitting every control measurably off the
    running form - and it wears the RUNTIME's look on this machine (light bar, dark caption
    left, close glyph right) instead of the classic blue guess. A TabStrip's tabs ride the
    walk's row now (they are not controls, unlike a MultiPage's pages, so the strip drew as
    a bare box), and ScrollBar/SpinButton wear the arrow caps the runtime draws, axis from
    their own aspect. Two of the owner's observed gaps are NOT canvas defects and are
    recorded as the design surface's meaning: a checked Taxable and a filled combo exist at
    RUN time only (`UserForm_Initialize` sets them), and the canvas - like the native
    designer - shows design-time state.
  - *The dead property slot, the owner's screenshot's real finding* - **found and fixed
    2026-08-14**: the launched form wearing `UserForm1` was the current fixture telling
    the truth. A form-level property set through the api landed on the DESIGNER dispatch,
    which accepts the write and echoes it back on every read - route, markup, canvas all
    agreed with each other and with nothing real - while the form frame paints a different
    slot entirely: the component's Properties collection, the native Properties window's
    own. Measured by running the form and reading the real window's title (designer said
    "Quarter Entry", the running form and the design surface both said "UserForm1"; a
    Caption written through the bag came up on the runtime titlebar immediately).
    Form-level reads and writes now go BAG FIRST, designer dispatch only as fallback, in
    the product walk, the debug walk and the apply - so the markup document, the canvas,
    the native panel and the running form finally describe one form. The suite's run row
    doubles as the pin: the launched window must wear the caption the api set. The
    label-like "Quarter Entry" text inside the owner's run window remains unexplained by
    the current plan and is left as an open observation.
  - *Run is the form's, from a designer tab* - **landed with it** (the owner's ask): Run
    with a designer tab holding the active slot runs THE FORM, the editor's own F5 with a
    designer selected. The native designer window is made visible and focused as the aim;
    the editor POSTS the Run action and answers, reading the aim ~30ms later (measured
    2026-08-14 - the first cut put the window down synchronously after Execute, which
    un-aimed the posted action and degraded the run to the Macros dialog; the second cut
    put it down on the tick that saw the run START, which landed between run-start and the
    form window appearing and killed the launching form). So the window stands BEHIND the
    running form, exactly as the native editor leaves it, and goes down on the tick that
    sees the run OVER - or a 3s deadline when it never takes hold. A bonus fact from the
    same trace: a designer SURFACE is itself a child `ThunderDFrame` - the runtime's own
    class - which is why the `userform` route enumerates top-level windows only. The suite
    runs the fixture form, reads its runtime caption off the real window, closes it by its
    X's own message, and waits for the designer window and Toolbox to be down after. And
    the Toolbox is put down TOTALLY at the event layer now, a policy that took the owner's
    third AND fourth reports to finish (2026-08-14): the third fix matched the Toolbox by
    handle but still gated on the VBFloatingPalette class, and the owner's next screenshot
    plus an outside-the-process enumeration proved the standing window was the OTHER
    species - an Office "F3 MinFrame" tool-window frame, restored at SESSION BOOT by a
    workbook saved with designer state, which every earlier defense AND the enums that
    verified them filtered out by class, and which the object model's Visible sweep cannot
    even touch (route-end sweeps ran dozens of times while it stood). The rule now: a
    MinFrame show goes down unconditionally, because every native tool window this product
    replaces wears that frame and none of ours ever does; a VBFloatingPalette show goes
    down only when the object model's Type-10 handle says it IS the Toolbox, because the
    Watches and Locals ghosts share that class; and a Win32 sweep at hook-arm catches a
    Toolbox shown before anything watched. The old "designer surface" exemption is gone
    entirely - its class was the editor FRAME's, so the set it consulted had been empty
    since the day it landed.
  - *The parity probe* - materialise the native designer window (Visible on, capture by
    hwnd, Visible off, Toolbox down, NEVER saving while one stands - the restore trap),
    capture the canvas beside it, ship both images side by side; eyeball first, pixel-diff
    on control edges after. This is the canvas's definition-of-done row.
  - *The designer route refuses a from-disk form* - found 2026-08-13, late: the debug
    `designer` GET answered "has no designer to read" for a form loaded from disk and never
    touched, while the designer TAB read the same form fine seconds later through the
    identical `GetObject("Designer")` call. **Not reproducible at HEAD** (2026-08-14:
    first-touch route reads on fresh sessions answer the full form, before any tab). The
    likely-but-unproven variable is boot-time designer-restore state - the same restore
    that raises the Toolbox materialises the designer - which the refusing sessions lacked.
    The suite's FIRST row now reads the fixture's own from-disk form before anything
    touches it, standing exactly where the reproduction would begin.
  - *Rename-follows-tab* - **landed 2026-08-13, late**, and the parked finding is
    corrected: the live re-trace (a `waitForLog` blocking read armed BEFORE the rename,
    the move the first attempt's post-hoc log walk should have been) proved the product
    rename DOES cross `AdoptRename` - the adopt line fired within the second - and that
    on the unpatched build no second tab ever appears; the first attempt's double tab was
    its own broken patch, not an untraced path. So the fix is the one choke point the
    attempt assumed: `AdoptRename` re-keys `_designerTabs` and `_activeDesignerTab` in
    place, and re-keys `_lastNativeActive` so the publish does not read the shown module's
    new name as a native move and strip the tab of the active slot. The panel's `(Name)`
    branch stops closing the tab - following IS the adoption now. The page needed nothing:
    its reconciliation already swaps a re-keyed face atomically in the one publish that
    carries it. Suite rows rename there and back through both entrances (which also pinned
    that a renamed-away name is reusable in-session, unlike a removed one's). Unapplied
    markup edits do not survive the swap - the followed tab reopens canonical - and the
    tab may change strip position; both accepted until someone misses them.
  - *The panel targets a designer tab's form* - opening or activating a designer tab does
    not move `_propertiesTarget` today (probed 2026-08-13: the panel kept showing a prior
    component over an active designer tab), where the native designer click targets the
    form. One line in `OpenDesignerTab` when taken up, suite row beside it.
  - *Liveness beyond the funnel* - the panel's `editProperty` joins the refresh hooks, and
    native-side edits (only possible if the developer opens the native designer themselves)
    get a poll fingerprint gated on an open designer tab, with the COM-volume counter the
    risks table demands before it is called cheap.
- **M3 - the honest canvas.** A renderer of the same projection beside the markup: real
  bounds, real captions, honest placeholders; selection; double-click writes the event stub.
  (Spikes 3 and 6.) **First slice landed 2026-08-15** (the owner's go-ahead): a click
  SELECTS - a control by name, the form by its ground - dressing the selection in the
  native handles (an overlay beside the control, because every control box clips its
  overflow; M5 is where the handles learn to drag) and landing the markup caret on the
  selected thing's line, the two halves pointing at one thing. A double-click asks the
  HOST for the control's default event handler - Change for the value-bearing kinds,
  Click for the rest and the form, whose handlers answer to "UserForm" whatever the form
  is called - written into the code-behind through the product's own module write when
  absent, navigated to when standing, never duplicated. And opening a designer tab now
  targets the form in the Properties panel, the native designer's own selection (the nit
  queued 2026-08-13). `designerSelect`, `designerEventStub` and the grown `designerCanvas`
  drive all of it.
- **M4 - the inspector.** Selection flows into the Properties panel; property writes go
  through the model; the transaction log starts recording. (Spike 7.) **Bridgehead landed
  2026-08-15**: a canvas selection targets the panel at the CONTROL - curated rows the
  designer service can honestly round-trip (name, caption, geometry, state, colours,
  font), read tolerantly off the live control, edited through `SetControlProperty` so the
  panel, the api routes and the markup apply stay one write path, the open tab
  re-projecting per edit. Selecting the form's ground returns the panel to the component.
  The native panel enumerates the control's whole typelib; this one grows with the
  service, and the transaction log stays ahead - it opens with direct manipulation, where
  undo starts mattering. It opened the same day, below, and it turned out to be the
  markup document itself.
- **M5 - direct manipulation.** Click to select, drag to move, resize by handles, nudge,
  add from an xlide toolbox by true drag-and-drop, delete, align and distribute, undo over
  the transaction log - the native designer's full manipulation vocabulary, on our canvas.
  (Spike 2 in full. Confirmed as the road with the owner, 2026-08-13.) **Moving landed
  2026-08-15**: a press picks a control up, the pointer carries it, and the drop rewrites
  that control's line in the DOCUMENT - `at 84,38` becomes `at 108,50` - as one edit.
  Arrow keys nudge the selection a point at a time through the same commit.

  **The cursors took three tries with the owner, and the rule they settled on is worth
  keeping.** The HAND across the whole form face: the web's own "this responds to a press",
  which is exactly true, because every inch of the face selects something. The four-way
  MOVE only on the control that is already SELECTED - the one a press would actually pick
  up and carry. Worn by every control on hover, as it was for an hour, a four-way arrow is
  a claim about hovering rather than about dragging; worn by nothing, as the plain arrow
  that briefly replaced it, the canvas says nothing at all. A handle keeps its own resize
  cursor, aimed at a 7px target that earns the hint, and a drag in flight paints the whole
  canvas with its own gesture's cursor so passing over another control does not change what
  the hand is doing.

  **Resizing landed the same day**, and the handles stopped being decoration: each of the
  eight takes the pointer (the overlay around them still passes it through, so the
  boundary never eats a click), pulls its own edges, and wears the cursor for the pull it
  makes. A west or north handle moves the origin as well as the extent, which is what
  makes a corner feel like a corner rather than a move. Shift+arrow is the keyboard's
  resize where a bare arrow moves, the native designer's own pairing. The FORM resizes by
  its own frame, and its line takes a size and never a position. A pull past the opposite
  edge stops at a floor - four points for a control, twenty-four for a form, because a
  form that has lost its title bar is a mistake rather than a design - instead of
  inverting the box. It is one commit for every gesture: the same document write, the same
  undo element, the same draft preview, and a move that finds no `size` clause does not
  quietly add one.

  **Delete landed with them**, and it is the gesture that finishes the sentence the canvas
  started: the selected control's line leaves the document, and everything indented under
  it goes too - its properties, and a container's children - because a Frame whose header
  went but whose children stayed is not a document the parser would take. One undoable
  edit, like every other gesture, so a mistaken Delete costs one keystroke. The form keeps
  the control until Ctrl+S carries the removal through the apply's name-keyed diff.
  Selection lands back on the form, which is where the native designer leaves it and what
  returns the Properties panel to the component. The FORM itself cannot be deleted from
  its own canvas: removing a component is the tree's gesture, with the confirmation the
  product asks there.

  **The xlide toolbox landed 2026-08-15**, which is the thing the native Toolbox was
  suppressed FOR. A palette of the fourteen kinds the apply can add, docked in the tab
  rather than floating over the form, and a kind is dragged out of it onto the canvas: the
  pointer carries a ghost of the control's real size, the drop resolves the deepest
  CONTAINER under it - a Frame's client, a MultiPage's page, the form's own ground - and
  writes a line under that container, at that point, named the way the native toolbox
  names one (the kind plus the first free number) and sized the way MSForms sizes a
  control dropped rather than drawn. One undoable edit; the control is selected on
  arrival; Ctrl+S adds it to the form through the apply's own name-keyed diff, so a
  control born on the canvas is indistinguishable from one born through the api. A drop
  outside the form adds nothing rather than a control at the origin.

  Three notes from building it. The palette's icons are drawn FOR the palette: the first
  landing put a real canvas control in each button, on the theory that the two should not
  hold separate opinions about what a kind looks like, and at 16px every kind is the same
  grey rectangle, because what tells them apart on the canvas is their children. It is ONE
  ROW that scrolls, wearing the same edge arrows as the command strip and the tab strip
  (the owner's call): a second row of palette eats the canvas on a narrow tab. And the
  suite pins the palette against the ROUTE - the add route's own refusal names its kinds,
  so the two lists are compared rather than both being written down twice.

  One measured trap sits under all of this: **a gesture's origin comes from the DOCUMENT,
  never from the painted box.** The first resize read `element.offsetWidth`, which carries
  whatever border the renderer drew, so a form pulled twenty points wider grew twenty-one.
  The pointer path now starts where the keyboard path always did - at the numbers the line
  spells - and falls back to the picture only for a clause the line does not carry.

  And one measured latency: **a gesture does not wait out the typing debounce.** The canvas
  redraws from the host's own parse, which the lint round trip carries, and that request is
  debounced 350ms so a keystroke does not ask the same question mid-word. A drop and a
  delete were paying it: 347ms and 348ms from the act to the canvas, measured, and felt at
  once by the owner ("some delay when dragging an element onto the UI, or deleting"). A
  drag and a resize never showed it because they paint themselves as they go. Every commit
  now asks for its lint immediately - a gesture is finished the moment it happens - and the
  same two measurements read 4ms and 3ms.

  **The document IS the transaction log**, and that is the design decision this slice
  makes rather than the feature it ships. A canvas gesture writes the same text a hand
  would have typed, so there is no second model to keep in step: the draft preview shows
  the move at once, the dirty dot says it has not reached the form, Ctrl+S applies it
  through the one apply path, and Ctrl+Z takes a whole drag back from either half - from
  the canvas because the view routes undo to the document that both halves share. A move
  is a document edit, not a COM write, so the FORM does not move until the save; the
  Properties panel keeps reading the form's own truth meanwhile, which the tab's dot and
  the form's draft outline are there to explain. The base for a move is read from the
  DOCUMENT rather than from the picture, so the arithmetic stays right even when the canvas
  is a parse behind.

  **Undo has two rules, and both were learned the same hour.** A move opens its own undo
  element (a stack stop before the edit), so one Ctrl+Z gives back one gesture rather than
  every move since the tab opened. But the element stays OPEN afterwards on purpose, so
  the canonical print that lands after a save rides along with the move it echoes: without
  that, an undo after Ctrl+S first walks through a step that differs only by the machine's
  own rounding - the form reads back 50.000025 where the document said 50 - and looks like
  it did nothing. And the document's ARRIVAL is not on the stack at all: the first
  projection lands by `setValue`, which clears it. As an edit it left a "back to empty"
  step underneath everything, so one Ctrl+Z in a barely-touched tab blanked the whole
  document while the canvas kept showing the form - found by hand and by the owner within
  the same minute, with ten green drag rows standing. Two suite rows hold the line now:
  one gesture per undo, and twelve undos at the floor still leave the form's own text.

  Three rules the gesture keeps: a drag lands on whole points, because a hand-placed
  control wants round numbers (the 6-point grid the native designer snaps to is M6's); a
  drag cannot lose a control behind its parent's edge, so it stops at the corner - and
  dragging a control OUT of its Frame is reparenting, a later gesture, not a clamp
  failure; and a press below the drag threshold is still just a click, which is what keeps
  selection and moving one gesture. A line the view cannot rewrite - half-typed, a caption
  whose quote never closes - refuses the move rather than guessing, and a Page, which has
  no position of its own, offers no drag at all. `designerDrag` drives it with the real
  pointer sequence, aimed through the hit test.
- **M6 - the finishing set.** Tab order, z-order, zoom, snapping and guides - and the
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
