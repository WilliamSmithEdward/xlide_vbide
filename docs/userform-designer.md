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
  file. The refusal is the right behaviour and stays. *(It did not stay: import creates a form
  from its PAIR since 2026-08-16 - see the sync section below. The refusal was right about the
  code file alone and wrong about the two files together.)*

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
strings. A COLOUR is the one place it is not (2026-08-16, at the owner's direction, when the
Properties panel started spelling colours the same way): a plain colour is `#c0dcc0`, the
spelling every developer already has, and `&H8000000F&` stays for a SYSTEM colour, which is a
question about the machine rather than an RGB. Both spellings parse, one conversion in Core
serves the parser, the printer and the panel, and the model still stores the decimal. A header line is
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
    content showing - and a click on a tab opens any of them from 2026-08-16, under M5 below.
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
  3. *Completions, hover and the header hint* - **landed 2026-08-16** (the owner: "I'd like
     full hover intellisense / completions / hints etc. in the markup section").

     The vocabulary is MEASURED, and that is the whole design. A bare instance of each
     control's coclass was already being created for the defaults inventory; it now answers a
     second question while it is up - what each of those properties MEANS - read through the
     same `ITypeInfo` the Properties panel spells its enums with. So a completion offers a
     property MSForms on this machine actually has, a hover repeats the library's own sentence
     rather than one written here, and the two panels and the document cannot hold different
     ideas of what a Label is. `vocabulary` is the route, `FormMarkupVocabulary` shapes it, and
     the page asks once per session and holds it: no keystroke waits on a round trip.

     What it offers, by where the caret is. A line that can open a control offers the KINDS,
     filtered by what the container takes - under a MultiPage only `Page`, and a Page only
     under a MultiPage - each arriving as a whole scaffolded header with a free name, a
     caption, a position and MSForms' own drop size as tab stops. A line under a control also
     offers that control's PROPERTY names, and picking one opens the value list without a second
     keystroke. Past a header's name the CLAUSES are offered, never twice on one line and never
     inside a caption. On the value's side of an `=`, the property's enum members by the name a
     developer writes.

     **Hover answers what the line does not.** Over a KIND, what that class of control is; over a
     control's NAME, the declaration VBA would write - `OkButton As MSForms.CommandButton` - with
     the class described under it, which is the same sentence hovering that name in the CODE half
     already gets. Over a property path, its declared type, untouched value and members; over an
     enum member, the number behind it.

     What it does not do is read the line's own geometry back. The first cut did, and it was
     standing in the way of the text it was quoting (the owner: "the current hover information is
     superfluous... that's obvious from the markdown"). The position and the size are on screen
     under the pointer; the class of a `CommandButton` is not.

     **The coclass went the same way, 2026-08-16** (the owner: "the part on class rollover that
     says forms.optionbutton.1 says that on all of them... seems not helpful"). It was correct per
     kind - each card carried its own - and that was the problem: for the standard fifteen the
     ProgID is `Forms.` plus the word the pointer is already on plus `.1`, so every card spent a
     line restating its own heading. It stays only where it cannot be worked out, which is the
     third-party case it was there for, and the completion list drops it on the same rule.

     The fifteen class descriptions are the second table in this product whose WORDING is ours
     rather than measured, for the reason measured here: MSForms ships no help strings at all.
     The first is the system colour names, for the same kind of reason.

     The hint is the header's grammar with the clause the hand is on, followed by the clause
     KEYWORD standing rather than by counting words - `at` takes two numbers and a caption takes
     none, so word counting drifts immediately.

     **The page reads structure, and that is not a second grammar.** Which container a line sits
     in, whether it is a header or a property: positional, four spaces a level, suggestion-only.
     It can be wrong about what to OFFER and never about what is valid, because it refuses
     nothing - the squiggles keep coming from Core's tolerant parse, host-side, the apply's own.
     That is the line this feature had to stay on the right side of, and it is why the language
     moved into `formmarkuplang.ts` beside the grammar rather than into the view.

     Four things measured on the way, each of which changed the code.

     A font's Size is VT_CY - MSForms stores it as currency - and the printable filter did not
     have the case, so `FontSize` came back with no default at all.

     The DOTTED font paths were built and then deleted. The dialect spells `Font.Size = 12` and
     the apply reaches it, so the first cut walked each bare control's font object; it offered
     nothing, because an OLE font is a vtable interface whose getters declare their result as a
     retval PARAMETER and the property walk reads a parameter as "indexed, skip it". Chasing that
     was the wrong fix: a control's own extender already carries `FontName`, `FontSize`,
     `FontBold` and `FontItalic` flat, they measure cleanly, and an apply writes them - so the
     dotted spelling would have been a second way to say what the list already says. Two
     spellings of one property is a worse completion list, not a richer one.

     A bare MULTIPAGE throws out of its own type library rather than declining, which silenced
     the entire vocabulary until each half of the read was guarded per kind. It answers one value
     of its fifteen properties, and the other fourteen are offered without defaults - a property
     an unsited control refuses is still a property a developer can set.

     And MSForms carries NO help strings: every `doc` came back empty across all fifteen kinds,
     so a hover shows the type, the untouched value and the members, and the field stays for the
     libraries that do document themselves - a third-party control added by ProgID may.

     The FORM is the one kind with no coclass to instantiate bare, so its entry is described from
     a live form when the page names one, which is why the request carries a module at all.

     `vocabulary`, `act designerComplete`, `act designerHover` and `act designerHint` are the
     drive side; the suite compares what the page OFFERS against what the host MEASURED rather
     than against a list written down twice.

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
  - *And F5 runs the DOCUMENT, not the last save* - **landed 2026-08-16** (the owner: "F5 /
    Run should fully save the active form designer before launching", with two screenshots of
    a Hold button in two places after an undo). The document is the transaction log and the
    form catches up on a save, so a run that skipped the save launched whatever was saved
    last - the canvas and the running form disagreeing about a control that had just been
    moved back. Run over a designer tab now takes the same apply-and-save handshake Ctrl+S
    takes and rides its callback, so the window that opens is the document. A refused apply
    launches nothing and says why at the document, which is the right way round: the old
    behaviour would have run yesterday's form and said nothing at all.

    The intent travels WITH the request (`designerApplySave` carries `run`, and the page
    calls back through `saveOnlyThenRun`) rather than being remembered on the host, because
    a refused apply calls nothing back: a flag waiting for a callback that never comes would
    have fired on somebody's unrelated Ctrl+S minutes later.
  - *Snapping, one way or the other* - **landed 2026-08-16** (the owner: "can we add a
    toggleable snap to grid mode", then "can we add a snap to other objects realignment with
    guide lines, think the experience of moving objects in a powerpoint slide", then - having
    watched the two fight - "should we have it be either or (or neither)").

    So `designer.snap` is `grid`, `objects` or `off`, and never two at once. They collide when
    they disagree: a control near a neighbour's edge AND near a grid line has two right answers
    and takes whichever the code asked first, which reads as a designer that cannot make up its
    mind. `grid` ships, at six points, because that is the editor's own Align Controls to Grid.

    Pointer gestures snap - a drag, a resize by a handle, a drop out of the palette - and the
    KEYBOARD never does: an arrow moves one point whatever the mode, because the hand reaching
    for it has already decided that neither the grid nor the neighbours are where this control
    belongs, and a nudge that jumps six points is not a nudge. **Holding Alt escapes whichever
    mode is on**, for one gesture, read per pointer event so pressing or releasing it mid-drag
    changes what the next movement does.

    Two switches at the end of the toolbox row rather than a setting two dialogs away, because
    snapping is something a developer turns off for one awkward control and back on immediately.
    Each is a toggle of its own mode and the pair gives the either/or/neither; both write the
    SETTING, so the buttons, the dialog's row and `settings?designerSnap=` are three views of
    one fact.

    In `objects` mode the candidates are the edges and CENTRES of the siblings sharing a
    container, plus the container's own inside edges and middle - siblings only, because
    lining up with a control in another box is lining up with a coincidence. A guide is drawn
    where the gesture lands, inside that container, and comes down on release; the acts answer
    what they lined up with, which is the only way a probe can tell an alignment from an
    accident.

    **A container offers the edge it PAINTS.** A Frame's rectangle starts at the top of its
    caption band and its rule is four points lower, so a guide at the rectangle ran through the
    lettering and the control lining up with it looked aligned to the caption (the owner: "the
    button should snap to the frame's edge, not the label"). A developer means the line they
    can see.

    The grid is painted as a repeating background on the form's client - 3,180 cells on a
    360x320 form, and that many divs is a canvas that stutters - at a sixth of the form's ink,
    shifted back half a cell so a dot lands ON the coordinates the snap produces. Asked whether
    controls should sit exactly on the dots: yes, and they did not until that shift.

    **It also found a bug in the settings file, four settings old.** The record's properties
    were `init`, so the JSON source generator had to set every one of them in a single object
    initializer - passing `default` for each key the document did not name. A settings file
    naming only `format.indentSize` therefore produced FALSE for every boolean in the record.
    Nobody could see it while every key was in every file; the grid was the first setting added
    since, and it read back off, at a two-point spacing, on a machine whose file predated it by
    an hour. Settable properties let the generator construct first and assign only what is
    there. A row pins it: a key the file never mentions keeps its shipping default.
  - *A closed form is not yet an unloaded one* - measured while pinning the above. For about
    400ms after the close is posted, the component answers "no designer to read": the object
    is gone, not empty. `debugMode` is no guide (the log has it back at design while the form
    still stands) and neither is the running-forms list, which empties first. Anything reading
    a designer after a run waits for the designer itself.
  - *The parity probe* - **landed 2026-08-16** as `tools\harness\designer-parity.mjs`, and it
    found what it was built to find. Not the native designer window in the end but the RUNNING
    form, which is the surface a developer judges by: launched from the tab (so both surfaces
    hold the same document), photographed through `capture?window=form`, and compared landmark
    by landmark against the canvas's own DOM. Both sides reduce to points from the form's
    client origin, and the landmark is what each kind actually PAINTS at its top - a frame's
    rule, a tick box's glyph, the box itself for the rest - because the model's rectangle is
    the thing the two already agreed about.

    Three findings, all of them the same shape: **the canvas was drawing a container's
    rectangle where the model's rectangle is, and the runtime does not.** A Frame's caption
    band belongs to the control's own box, so its rule is about four points lower and its
    first child about nine - the canvas drew the rule at the box edge, which lines a frame up
    with the button beside it on screen and nowhere else (the owner's side-by-side). A
    MultiPage draws nothing at all above or beside its tabs, and the canvas drew a rectangle
    all the way round. And the model's `InsideHeight` reads about two and a half points short
    of the runtime's band for a frame, so the band comes from the caption's own line box now,
    which measures true.

    Before: the frame four points out, its children six, the multipage structurally wrong.
    After: every control's painted top edge within a point of the running form's, and most
    within a third of one. The probe prints rather than passing, the way the perf walk does.

    Its own three false starts are worth remembering, because each one produced a confident
    table of nonsense: it calibrated the client origin off a checkbox INSIDE a page (a
    container's child carries its parent's coordinates), it scanned a window that reached up
    into the control above and read that control's bottom as this one's top, and it aimed at a
    container's left edge where the column runs through the caption lettering rather than the
    rule.
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

  **A GROUP, and the arrange vocabulary, 2026-08-16** - the rest of M5's own sentence ("delete,
  align and distribute"). Everything on the canvas assumed one selection, so there was no Align
  Left, no Make Same Size and no Distribute, which is most of what the native Format menu carried.

  The selection is an ANCHOR and the rest: `selectedName` keeps its meaning - the one the handles
  dress, the one the Properties panel follows, the one an alignment lines the others up with -
  and everything else rides beside it. That is the native designer's primary control, and keeping
  it as the same field means every gesture written before groups existed still reads correctly.

  Ctrl+click gathers one, a rubber band over a container's own ground gathers several - what it
  TOUCHES rather than what it encloses, MSForms' own rule and the more forgiving one. A group is
  one container's business: a control from another box starts a new selection instead, because
  controls measured from different origins cannot be moved together or lined up in any way that
  means something on screen. The rest of the group wears a boundary and no handles, because only
  the anchor can be pulled.

  A drag or a nudge moves them all by the same delta, a Delete takes them all out, and each is ONE
  edit and one undo step - which is what the multi-line write exists for. Align, Same Size and
  Space are the same commit again: they compute each control's box and write them together.

  They live on the canvas's own CONTEXT MENU, and that is the answer to where the Format menu
  went. This product has no menu bar; the editor's own Format menu stays suppressed and must,
  because it would act on the native designer's selection rather than on ours. So the menu appears
  on a group and not on a single control - an item that cannot run is worse than no menu when the
  whole gesture is "do this to these".

  **A drag carries a control OUT of its container, 2026-08-16** - the reparenting M5 deferred
  ("dragging a control OUT of its Frame is reparenting, a later gesture, not a clamp failure").
  Until this, a drag stopped at the container's edge and moving a control from a Frame to the
  form could only be done by editing the markup by hand.

  It happens WHILE the drag runs rather than at the drop, and that is the whole implementation:
  every control box hides its overflow, so a control dragged past its parent's edge is simply
  clipped away and the developer is dragging something invisible. The control is carried into
  whatever container the pointer is over - the same hit test the toolbox drop uses - and the
  gesture is rebased as it goes: the origin becomes the box's position in the new container's
  coordinates and the press moves to here, which is exactly what the arithmetic measures from.
  The clamp then keeps it inside its NEW home, because the clamp reads the element's parent.

  Two guards, both of them the document's own logic rather than the picture's: a container
  cannot be carried into itself or into anything it holds, and the drop refuses a move whose
  destination is inside the block that is moving. Only a MOVE reparents - pulling a handle is a
  statement about size, not about belonging - and the dragged control is transparent to the
  pointer while it travels, or the hit test would answer "itself" and no drag could ever find
  where it was going.

  The drop rewrites the DOCUMENT: the whole block, re-indented a level under its new home and
  appended at the end of it, as one undoable edit. It is a whole-document splice rather than two
  ranges, because the two collide the moment the block being moved is the last one in the file.

  **A reparented control is a NEW control, and that is measured**: the apply's name-keyed diff
  calls a changed container a remove-and-add, which is the truth of what it means, so the control
  MSForms builds afterwards carries only what the markup printed. Moving a control out of a frame
  and saving, the fixture's option button came back with a fresh tab index (0 to 12) and the
  form's own font in place of the frame's. Everything the projection prints survives; everything
  it does not is born again at its default. That is the strongest argument yet for the sited
  baseline this document keeps as an open item - the more the document says, the less a reparent
  costs.

  **The containers opened, 2026-08-16.** A MultiPage drew its first page and nothing reached
  the second: page two's controls were in the document, invisible, and a drop always landed on
  page one whatever the tabs said. Now a click on a tab opens that page - its controls draw,
  the page is SELECTED (the Properties panel follows to it, the markup caret lands on its
  line), and the next drop lands on the page being looked at.

  What made it small is that **the page body IS the page**. It carries the page's identity now
  rather than being an anonymous box inside the MultiPage, so a press on it selects the page
  the way a press on the form's ground selects the form, a drop reads the page's name off the
  thing it landed on, and the selection has something to dress. It is dressed in an outline
  and no handles: a Page has neither a position nor a size of its own, and eight grips that
  can pull nothing are a promise the canvas would not keep.

  It runs both ways, which is the same rule as the markup caret's: **selecting anything inside
  a page opens that page.** A caret on a line in page two's block used to select nothing at
  all - the canvas had no element for a control it was not drawing, so the selection fell back
  to the form - and now it opens the page and dresses the control.

  **New Page and Delete Page sit on the tab strip's own context menu**, where the native
  designer keeps them, and both write the DOCUMENT as one undoable edit: New Page appends
  `Page PageN "PageN"` under the container and opens it, Delete Page takes the page's line and
  everything under it, and one Ctrl+Z gives a whole page back with its children. Delete Page
  acts on the page that is OPEN whether the right-click landed on a tab or on the empty end of
  the strip, because what a developer means by "this page" is the one on screen; a MultiPage
  the document has emptied draws as bare chrome and greys the item rather than offering to
  delete nothing.

  A TabStrip switches too and honestly shows nothing new, which is the whole truth about it:
  its tabs are an INDEX rather than containers, and the runtime draws the same controls under
  every one. ~~It gets no page menu, because its tabs are not in the dialect - they ride the
  walk as strings for painting, there is no line to add or remove, and every item would be a
  claim the product cannot keep.~~ **Superseded the same week:** the dialect learned `Tab` lines
  and the strip got its menu, three sub-sections down. The gap was real when this was written and
  the refusal to paper over it in a menu was right; what closed it was closing the gap.

  **Which made the tab's own appearance the whole of the feedback, and it was not enough** (the
  owner, 2026-08-16: "in the tab strip, there is no differentiation between tab 1 and tab 2").
  On a MultiPage a click changes the content, so a faint mark on the tab is a confirmation; on a
  TabStrip nothing else moves, and a 5% wash plus bold was all the difference there was. The open
  tab is drawn the way the runtime draws one now - raised into the body, its ground undarkened,
  the closed ones lower and darker - which is both clearer and closer to the surface a developer
  without xlide sees.

  **And the FIXTURE was telling a lie about the control** (the owner, the same evening: "there is
  a button drawn over it, but changing the tabs doesn't visually do anything"). The fixture's
  TabStrip was 86 points tall, which put its face under the OK button at 250, and it held nothing
  of its own - so it read as a broken tab control rather than as the one control on the form that
  needs code to mean anything. It is shorter than the MultiPage beside it now, it carries a label
  on its face captioned "Views_Change fills this", and the code-behind has the handler that makes
  the point. Nothing about the CONTROL was wrong; the demonstration was.

  **Then the dialect learned tabs, 2026-08-16** (the owner, looking at the document: "in the
  markdown, i dont see anything indented under the tab view"). The canvas drew two tabs and the
  strip's line stood with nothing under it, because a TabStrip's tabs rode the walk as strings for
  painting and were in no other way part of the projection.

  A `Tab` is a line now, shaped like a `Page`'s - identity and a caption, nothing else to say -
  and it goes the whole way: the walk emits one row per tab, the printer prints them under the
  strip, the parser takes them back (a TabStrip contains Tabs and nothing else; a Tab sits under a
  TabStrip and holds nothing), the apply adds and removes them through MSForms' own `Tabs.Add` and
  `Tabs.Remove`, and the strip's menu offers New Tab and Delete Tab exactly as a MultiPage's
  offers pages. Measured end to end: a tab added from the menu reaches the real form at Ctrl+S and
  leaves it again the same way.

  **And the two kinds with no coclass got their vocabulary.** A Page answered ZERO properties to
  completion and hover, and a Tab was not in the vocabulary at all - both for the same reason the
  Form was not: `ControlDefaults` measures a bare instance of a coclass, and neither has one. They
  are described from a LIVE one on the open form now, through the same walk the Form entry uses:
  24 properties for a Page, 5 for a Tab, measured rather than written down, and no form open means
  an empty list rather than an invented one.

  **The class hover leads with the type** (the owner: "can you add the class type to the class
  hover?"), so hovering `CommandButton` answers `MSForms.CommandButton` and then what the class is
  - the declaration a developer would write, which is the one thing about a kind that is not
  already on the line under the pointer.

  **Which page is open is VIEW state, and that is this canvas's one designated deviation from
  the native designer**, which writes the container's `Value` and dirties the form. Reaching
  page two must not rewrite the developer's form: the document is the transaction log, so a
  switch that landed in it would make LOOKING a change and Ctrl+S would carry it to the form.
  Navigation is not manipulation here, the way scrolling and selection are not. And the
  dialect has no honest line for it anyway - a MultiPage's measured vocabulary is fifteen
  properties and `Value` is not among them - so a form that should OPEN on a given page says
  so where every other unprinted property is said, the Properties panel.

  One bug paid for on the way, and it is a general one: **opening a tab redraws the canvas, so
  the element the event arrived on is detached the moment the handler acts on it.** The strip
  menu read which page was open off its own `strip` element after opening the right-clicked
  tab, got the previous picture's answer, and deleted the page the developer had just left.
  Anything reading the canvas after a redraw reads it again from the canvas.

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

  **Z-ORDER landed 2026-08-16, and the measurement is the design.** MSForms' ZOrder is a METHOD
  rather than a property, so nothing in this product could reach it; now Bring to Front and Send
  to Back sit on the canvas's own context menu, offered for one control as well as for a group.

  Two things were measured before any of it was built, and both changed the shape:

  - **The Controls collection is not in z-order.** Calling ZOrder front and back on a control
    leaves the walk's order exactly as it was. So the projection cannot see depth, the markup has
    no way to say it, and the canvas cannot draw it - it paints in collection order, which is
    creation order, and that is now a recorded parity gap rather than an unexamined one.
  - **ZOrder itself works, proved on the running form.** Two overlapping opaque labels, red under
    blue: 3,669 red pixels against 8,040 blue before, and exactly the reverse after. That
    photograph is the only instrument that can see this feature, and the suite takes it.

  So depth is the ONE canvas gesture that writes the model rather than the document, and it is a
  designated deviation with its reason: the dialect cannot express it, so there is nothing to
  write. It behaves like a Properties panel edit - at once, no Ctrl+S - and the developer sees it
  when the form runs.

  **ZOOM landed 2026-08-16**, at the end of the palette row: a menu of the usual percentages and a
  Fit, on a button wearing the percentage it is drawing at. Anything asked for is clamped to 25%
  at the bottom and 400% at the top; Fit never goes past 200%, because a small form blown up to
  fill a wide tab is not what Fit means.

  The implementation is one decision. The picture is a CSS transform on the form itself, so
  everything INSIDE it - every control's left and top, every guide, every snap - stays in the
  form's own coordinates and knows nothing about the scale. Only the places where screen pixels
  cross into points were touched, and they all go through two functions now. What that buys is
  the property worth having: a drag of twelve points moves a control twelve points at 200% as it
  does at 100%, which the suite pins by dragging at a zoom and reading the DOCUMENT.

  The form rides a stage sized to the scaled dimensions, because a transform does not change
  layout: without it the scroll bars would still describe a form at 100% and half of a zoomed one
  would be unreachable. Fit measures the canvas as it stands, takes off the margin the stage
  carries, and picks the smaller ratio - 78% for the fixture's form in a short tab.

  It is VIEW state, not a setting, for the reason the open page is: it is about looking rather
  than about behaviour, and a big form wants Fit where the one beside it wants 100%.

  **TAB ORDER landed with it**, as a dialog of xlide's own on the canvas's context menu. The
  native one (View > Tab Order, menu 469) stays suppressed for the reason every suppression here
  has: it would act on the native designer's selection rather than on ours.

  Why a dialog when the Properties panel already carries TabIndex and can already write it:
  because that is exactly the part nobody can use. A tab order is a SEQUENCE, and a sequence is
  edited by seeing all of it at once - typing 4 into one control's box while MSForms silently
  renumbers the other eleven is a puzzle rather than an edit. The dialog lists one container's
  controls in tab order with Move Up and Move Down, and each move is a single TabIndex write
  through the same `SetControlProperty` the panel and the api use: the moved control takes the
  index it is going to and MSForms pushes the other one along, which is the native dialog's own
  behaviour and the reason the page never computes a new order.

  Two things it gets right that are easy to get wrong. ONE CONTAINER at a time, because tab order
  is per container in MSForms - a Frame's children have their own 0..n and Tab walks into the
  frame and out again - so selecting a control inside a Frame opens the Frame's list, not the
  form's. And a control with NO tab stop is left out entirely: an Image cannot take focus and
  MSForms answers null rather than a number for it, so listing it would put a row in the sequence
  that Tab never visits and number every row below it wrong.

  The projection carries `tabIndex` for this - display truth, like the fonts and the client areas
  beside it, not a line the dialect prints.

  **PICTURES landed 2026-08-16**, which closes the oldest open risk in this document: M3 named
  `IPictureDisp` extraction as the likely hard case, the risks table below guessed it might need
  an OLE round trip through a temporary stream, and until now an Image control drew as a crossed
  box whatever it held, a form with a picture painted without it, and the Properties panel dropped
  the row entirely.

  It needs no stream and no temporary file. The way in is the one part of an IPictureDisp that IS
  on its dispatch interface - `Handle`, a GDI HBITMAP or HICON - and from there it stops being COM
  and becomes Win32: `GetDIBits` for a bitmap, `DrawIconEx` onto a DIB section for an icon (an
  icon is a colour bitmap and a mask, which `GetDIBits` cannot read at all), and a BMP is
  fifty-four bytes of header around the pixels. Top-down 32-bit, so there is no row-padding
  arithmetic to get wrong, and one repair that matters: GDI leaves the alpha byte zero for every
  bitmap that has none of its own, and a browser given a 32-bit BMP whose alpha is zero throughout
  may draw nothing at all - so a picture with no alpha ANYWHERE is made opaque, and one with real
  alpha keeps every byte of it.

  **The write needed two loaders, and finding that out is the measurement of this slice.**
  `OleLoadPicturePath` is the obvious call and it answered `0x80004005` for a PNG - OLE's picture
  loader reads what OLE has always read (BMP, GIF, JPEG, ICO, CUR, WMF, EMF) and PNG is not on the
  list. So the refusal falls through to GDI+, which decodes anything this machine has a codec for,
  and the bitmap it hands back is wrapped with `OleCreatePictureIndirect`. The ORDER is the design:
  OLE first, because it keeps an icon an icon and a metafile a drawing, where GDI+ would flatten
  both to pixels.

  A flattened picture has one honest question in it - a PNG's transparent parts have to become
  some colour, because an OLE picture is a bitmap and a bitmap has no alpha - and the answer is
  **the control's own BackColor**, which is the colour that makes them disappear. Read at the
  write site, resolved through the live system palette like every other colour here.

  On the canvas the placement is exact rather than approximated, because MSForms' two placement
  families map onto CSS without a remainder. A SURFACE picture (the form, an Image, a Frame's
  client, a Page's body) is a background: size mode 0 clips, 1 stretches, 3 zooms, alignment is a
  background-position and tiling is a repeat. A CAPTION picture (a button, a Label, a check box)
  is an `<img>` in a flex line, and `fmPicturePosition`'s twelve values ARE four directions by
  three cross alignments - so all twelve are drawn, not five, plus the thirteenth (Center) as a
  background behind the caption. An `img` rather than a background for those, because a picture
  beside a caption draws at its natural size and an `img` is the one element that knows what that
  is.

  The panel's row has no text box, which is the point rather than an omission: MSForms keeps the
  pixels and forgets the file they came from, so a field showing a path would show a path that is
  not true a moment later. It says what it HOLDS in the native designer's own words - `(None)`,
  `(Bitmap)`, `(Icon)`, `(Metafile)` - shows a thumbnail over a transparency chequerboard, and
  offers Browse and Clear. Browse raises the machine's own file dialog host-side, because a page
  cannot hand back a path; that is the one place the two roads differ, and they meet again at the
  write, which is the ordinary property write the api makes.

  **The DIALECT says nothing about any of this, deliberately.** A picture is binary in the form's
  `.frx` and MSForms does not remember where it came from, so a printed path would be a lie and a
  printed bitmap is not a document. It rides the projection with the fonts and the colours - and
  that is where the defect was: the DRAFT preview dresses a parsed document in the applied
  projection's display extras by name, and the wardrobe carried fonts, colours, insides and tabs
  but not pictures. So every image on the form blanked the moment a drag made the document dirty,
  and came back at Ctrl+S. Measured, not reported: the model kept the picture and the canvas lost
  it. One field.

  **Cost, measured rather than assumed**, because a 256x256 logo is 350KB of base64 and this rides
  every projection: a full gesture loop - drag, document rewrite, re-projection, canvas agreeing -
  runs 5ms median with no pictures, 7ms with 700KB of them. Two milliseconds. No cache, therefore,
  and no downscaling: both would be machinery bought for a cost that is not there.

  **And the layout bug the work uncovered, which was making gestures land on the wrong half.**
  The canvas's scroll box carried its 24px padding on the SCROLL PORT, and a padded port cannot
  shrink below its own padding: on a short tab the port stayed 48px tall inside a half only 37px
  high. Its parent clipped the paint, so the canvas looked right - and every gesture aimed by
  coordinates went wrong, because `scrollIntoView` and `scrollToPoint` measure the port's own
  rect and would happily place a pointer twenty pixels below the clip, where the markup editor
  is. A toolbox drop there answered "nothing landed", which is how it was found; the refusal
  says what the pointer actually hit now, and that one sentence is what turned a shrug into a
  diagnosis. The gap is a margin on the form now, which belongs to the content and scrolls with
  it. In the same pass the stacked split stopped starving the canvas: the document's half is
  capped at what leaves the form 120px, with a 60px floor of its own, so a short tab cramps both
  halves rather than making one of them unusable.

## The panel speaks the developer's language

**Landed 2026-08-15** (the owner, looking at a form's rows: "can we get the actual enums
loaded? instead of the int representation"). The panel had been showing what the object model
hands over - `BorderStyle 0`, `Cycle 0`, `MousePointer 0`, `BackColor -2147483633` - and every
one of those is a name in the language the developer writes.

The names come from where the Object Browser's come from: the object's own `ITypeInfo`. For
each property getter the return type is followed, through an alias if it is one, and where it
lands on an enum that enum's members are read with their values. So a row shows
`fmCycleAllForms`, offers the members it could hold, and takes a member name, a hex literal or
the raw number on the way back in - `PropertyTypes.Spell` and `Unspell` are the one pair of
translations, and both panels use them: the component's rows and the designer's control rows,
because they are one panel as far as the developer is concerned.

Three decisions worth keeping:

- **Nothing is guessed from a property's NAME.** `fmBorderStyle` for `BorderStyle` is a
  convention until the day a control names one differently, and a panel that renames a value
  wrongly is worse than one that shows the number. `StartUpPosition` is a plain Integer in
  MSForms however much it looks like an enum, so it keeps its number, and a suite row pins
  that it does.
- **The row stays typeable, and its caret is always there.** The values are offered beside a
  field rather than instead of it (the owner: "the user should be able to type into a type if
  it's selected"), so the list drops down and the field still takes a number or a name nobody
  listed. The first landing used a `<datalist>`, which looks like that from the outside and is
  not: it filters its list by what the field already holds, so a row showing
  `fmScrollBarsBoth` offered exactly one choice - itself - and its caret only appeared on
  hover. It is a real combobox now: the caret is always visible and shaped like the arrow on
  the True/False rows beside it, the list shows every member, and typing filters it only while
  the developer is typing something new. **A colour row gets a swatch** that paints the value
  the host resolved - system colours included, because only the host knows what the machine
  calls a button face today - and opens a picker of xlide's own (below).
- **A refusal is the host SAYING so**, not the row failing to echo the request. The
  `editProperty` act compared the two, which was right only while every value had one
  spelling; a written `0` that reads back `fmCycleAllForms` had it reporting an honest write
  as refused. It watches the status line now, and answers with what the row holds.

Read once per class and cached: a loaded type does not change, and the panel republishes on
every selection.

**And a document component too, the same evening** (the owner: "please update enums for
workbook / worksheet properties"). A worksheet's values live in EXCEL's library rather than the
editor's, and a worksheet has no designer to point at one, so those rows had gone on reading
`Visible -1` and `EnableSelection 0` while every form in the product spoke in names. The panel
goes to the HOST object for them now - the sheet or the workbook itself, matched by code name
through the same trust-free application route the unsaved dots use - and the rows read
`xlSheetVisible`, `xlNoRestrictions`, `xlUpdateLinksUserSetting`, offering their members and
taking a member name back.

The obvious route was wrong and looked right: a VBE `Property` answers `Object` with its own
VALUE, not with the thing it belongs to, so asking the first property in a worksheet's
collection handed back Excel's Application - and the panel learned the enums of the wrong class
entirely, quietly, because a wrong library still answers.

**And the colour rows read like colours, 2026-08-16** (the owner: "for color pickers in
properties, can you use #f0f0f0 format?"). A colour now reads `#rrggbb`, the spelling every
developer already has, instead of the `&Hbbggrr&` a form's binary speaks.

A SYSTEM colour does not, and that is the design decision in this slice rather than the feature.
`&H8000000F&` is not a colour at all but a question - what does this machine call a button face -
and spelling it `#f0f0f0` would answer that question permanently, freezing today's theme into the
form. So those rows read the NAME (`Button Face`), the swatch paints what the machine answers
now, and picking one writes the question back. The write takes all four spellings: `#rrggbb`, a
system name, `&Hbbggrr&`, or the raw number.

**And the rows that said `[object]` say something now, or nothing** (the owner, the same evening:
"what about properties that say [object]?"). A form's panel was listing 51 rows, six of which
read `[object]` and did nothing: `Font`, `Picture`, `MouseIcon`, `Controls`, `ActiveControl`,
`Selected`. An object is not a value, and a row that says so is a row asking to be ignored.

A FONT is the one object this panel can serve, and it serves it as its parts - `Font.Name`,
`Font.Size`, `Font.Bold`, `Font.Italic` - which is what the CONTROL rows already did, so the two
panels now build those rows from one place instead of one having them and the other not. The rest
go: `Controls`, `ActiveControl` and `Selected` are runtime state the native panel does not show
either. `Picture` and `MouseIcon` waited on the picture pipeline this document kept as an open
risk, and since 2026-08-16 they are rows of their own - see PICTURES above. The font comes from
the designer rather than from the VBE property wrapper around it, which
hands back a font this side cannot read through - measured, after the first cut showed no font
rows at all.

Two more filters came with it, both MEASURED rather than listed. A leading underscore is the
library's own business (`_Font_Reserved` was on screen), which is the rule the Object Browser and
the defaults walk already keep. And a member the library marks hidden or restricted goes -
`DesignMode` did. What survives is 42 rows.

**Writable stopped being a guess in the same change.** The panel decided it from the VARIANT that
came back, so `CanPaste` - a Boolean with no setter - drew as an editable row. It comes from the
library's own PUT now, and everything else is writable because the type says so. The code name is
the one exception, deliberately: `(Name)` is the VBE's rename gesture rather than a designer
property, and the designer's library has no say over it.

**And then those rows left entirely, 2026-08-16** (the owner, on why three of them read False and
grey: "if theyre jot settable; dint suow them"). A panel is for setting properties. `CanPaste`,
`CanUndo` and `CanRedo` are questions about the editing SESSION rather than about the form - what
the clipboard holds, what the undo stack holds - which read False on a form nobody has touched and
cannot be written at all; `InsideWidth` and `InsideHeight` are measurements the canvas reads for
its parity and a developer cannot type into. Every row a UserForm's panel shows now can be set.

**And the font rows became pickers, 2026-08-16** (the owner: "should font be a drop down selector
for all properties?"). `Font.Name` had stayed a text box while every enum row beside it offered its
members, which asked the developer to spell a face exactly right or get nothing.

The list is MEASURED, like the system colours and for the same reason: no list written down here
could be true on another machine. GDI's own family enumeration answers it - the call every native
font picker makes - once per session, sorted, with the `@`-prefixed vertical variants of CJK
families left out the way every native list leaves them. 274 faces on this machine.

`Font.Size` gets the ramp every office application offers, and that one IS written down, because a
point size is a number rather than a capability and there is nothing to ask. Both rows stay
TYPEABLE around their lists, which is the panel's standing rule: MSForms stores a face as a plain
string, and a form written on another machine may name a font this one has never had - offering
the list must not turn that value into an error.

The same list reaches the DOCUMENT, because it is the same measurement: `FontName` completes to
this machine's faces in the markup, one walk feeding both surfaces so the panel and the document
cannot hold different ideas of what fonts exist. That took one thing the completion had not needed
before - a face is a STRING in the dialect, so what a suggestion REPLACES matters as much as what
it inserts. Accepting one where the developer has already typed `"Tah` takes the quotes with it
(`FontName = "Tahoma"`), where the word-shaped token range would have nested a second string inside
the first. `designerComplete` answers `replaces` now for every suggestion, which is what let that
be pinned rather than eyeballed.

The picker is xlide's own for the same reason it is not `<input type="color">`: that opens
Windows' colour dialog, in another visual language, in the middle of a surface built to replace
exactly that - and it has nowhere to put a question. This one has a generated palette (a grey row
and eight hues down seven lightnesses, written as a ramp so the grid cannot carry a typo) and a
System half fed by `GetSysColor`. Only the wording of the thirty system names is written down,
because no call hands back a display name for `COLOR_BTNFACE`; every value beside them is
measured. `act colourPicker` drives it, `setSystemColours` carries it to the page at load, and
the suite picks from what the palette actually offers rather than from a hex written twice.

**A control's changed COLOURS ride the document, 2026-08-16** (the owner, after the panel work:
a control's changed properties are invisible in the document). Until this, the projection carried
identity, containment, geometry and caption and nothing else, so a colour set in the Properties
panel lived only in the object model: absent from the text, absent from the draft preview, outside
the document's undo, and lost when a Frame was copied between forms.

What "changed" means took two baselines and a measurement to get right, and the measurement is the
part worth keeping. The obvious baseline is the defaults inventory - a bare instance of the same
coclass - and it is WRONG for a control that has been sited:

| | bare instance | freshly added to a form |
| --- | --- | --- |
| `Frame.SpecialEffect` | 0 | 3 |
| `ToggleButton.BackColor` | -2147483643 | -2147483633 |
| any control's font | MS Sans Serif 8.25 | Tahoma 8 (a Frame: 8.34) |

MSForms initialises a control differently when it joins a form, so a bare comparison prints
choices nobody made: the first cut put a font line under every control on the fixture and a
`SpecialEffect` under the Frame. The form's own values are not a baseline either - a Label
inherits the form's button face, but a TextBox, a ComboBox and a ListBox are born with the WINDOW
colours and keep them on any form, so comparing against the form printed a `BackColor` under every
entry control instead.

So a colour is the developer's only when it matches NEITHER: not what the form passes down, and
not what the kind is born with. That is exact, and it is what ships - two lines appear when two
colours are set, and both leave when the colours go back.

**The rest of the properties wait on a truthful sited baseline.** Two ways looked honest at the
time: probe a control of each kind onto the form and take it away again, which dirties a workbook
nobody asked to dirty, or read MSForms' own answer out of a `.frm` export, at the cost of exporting
during normal use. The owner chose a third that neither writes nor exports - read the storage Excel
has already saved - and it is measured further down, in "The sited baseline".

**The design rides SYNC as text, 2026-08-16** (the owner: add the markup to the import and export
path). A UserForm exported as code alone cannot be put back - Excel writes every control into a
binary `.frx` and the text merely names it - so a form in source control has always been half a
form: the code diffable, the design a blob. A form now exports as three files:

```text
EntryForm.cls    the VBE's own export, code and the form's own properties
EntryForm.frx    the controls, in MSForms' binary
EntryForm.form   the design as xlide's markup, the same projection the tab edits
```

The `.form` is a row of its own in the plan, so a developer sees it, ticks it and diffs it like
any other file. On import it goes through the markup's own name-keyed diff - the apply Ctrl+S
makes - so an edit made in a text file reaches the real control: the suite moves `OkButton` to
40,40 by rewriting a line on disk and then reads 40,40 off the form. A `.form` whose form is not
in the project is skipped, saying to add the form first, because the markup can build controls on
a form and cannot conjure the form itself.

**A form is exported BY THE VBE now**, rather than assembled from spliced text, and that fixed a
defect the change uncovered: the exporter names the sidecar in an `OleObjectBlob` line, and a
spliced file named the TEMPORARY path the header had been read from - a fresh GUID on every
export, pointing at a file that was already deleted. The exported form now names `EntryForm.frx`,
which is sitting beside it. The encoding compromise the sync service otherwise exists to avoid is
accepted here deliberately: a `.frm` and its `.frx` must agree byte for byte, and only the
exporter can promise that.

**Two planners, and only one of them is ours.** Sync's default planner is the companion editor's,
shared through the engine, because both products write into the same folders. It has never heard
of a form's markup and should not, so the design rows are built in Core and added to whichever
plan came back. The `.frm` extension for a form's CODE file is the opposite case - a shared
convention - so it belongs upstream rather than as a local patch, and a form's code still exports
as `.cls` here until that lands.

**AND IMPORT CREATES THE FORM, 2026-08-16.** The refusal that stood here - "a UserForm's designer
is not in this file" - was only ever true of the code file ALONE. The pair the VBE's own exporter
writes is exactly what its importer reads, so the form comes back whole: nineteen controls, both
pictures, and the code, into a workbook that had none. That is the last thing that stood between
this product and a form living in source control.

Which file wins was the open question, and the answer is the one this document already argued
for: **the sidecar is authoritative on a create.** It carries everything the markup does not
print, so the create is the binary and the `.form` applies on top - where the dialect's own rule
keeps it safe, because an unspoken property is never touched.

The foot-gun is real and it is GUARDED rather than accepted: the markup's control list is TOTAL,
so a `.form` out of date with its `.frx` would prune controls the binary brought in. So the design
row beside a create is offered UNTICKED, with the reason in its warning. Pruning stays possible -
it is sometimes exactly what a developer wants, which is why the row is offered at all - and it is
now a decision rather than a surprise. Ticked deliberately, the apply reports what it did:
`+0 -0, 84 set` when the two files agree, which is the normal case.

Three things the implementation had to get right, each of them measured rather than assumed:

- **Content decides what a file is, not the extension.** The shared planner writes a form's code
  as `.cls` where this product writes `.frm` (xlide_vscode#21) and the bytes are identical, so a
  file beginning `VERSION 5.00` with the MSForms coclass in its first `Begin` is a form whatever
  it is called. That un-blocked this slice: waiting for the upstream naming fix was never
  necessary, because the naming was never the fact that mattered.
- **The import stages a COPY.** The editor's importer decides what to make from the EXTENSION, so
  a `.cls` imported as it stands becomes a class module holding a form's header - not a form, and
  not undoable into one. The pair is copied to a temporary folder as `Name.frm` beside the
  sidecar under exactly the name the `OleObjectBlob` line spells, which a copy can promise and a
  developer's folder cannot.
- **A refusal carries the editor's own log.** `Import` answers "Errors during load. Refer to
  'C:\...\Name.log' for details" - and that log sits in the staging folder this side is about to
  delete, so the reason used to die with it. It is read and quoted in the failure now. What it
  said, the first time: a pair whose `Begin` line and `VB_Name` disagree is refused. Renaming a
  form in source control is three edits to the header, not two.

## What a hunt across the surface found

**2026-08-16.** Every gesture and every route driven at its edges - an empty form, a zero-sized
control, a control nine thousand points off the form, a negative position, a caption holding the
dialect's own punctuation, ten undos on a document nobody edited, a resize past the opposite
edge, zoom at 10% and 500%, an empty document, hovers and completions past the end of the text.
Most of it held: the refusals said what they meant, the floors held, the quoted caption
round-tripped, and a delete reached the canvas in 12-20ms rather than paying the typing debounce.
Seven things did not.

**Closing a designer tab with unapplied edits lost them, silently.** The worst of the seven,
because it is data loss and it was invisible: the close was unconditional, and the host's own
comment said the page asked its own question - which it did not. So a developer who moved three
controls and pressed Ctrl+W lost the three moves with nothing said. The question is the PAGE's,
because the state is: unapplied markup lives in the view, and nothing host-side can know to hold
the close. It is the same modal a module's unsaved text raises, with the same three buttons, and
Save means what Ctrl+S means on that tab - the document applied to the form, then the workbook.

**A removed form left its designer tab standing**, and what stood in it was an overlay reading
"EntryForm2 has no designer" (the owner: "sometimes i see an overlay that says entryform2 has no
designer"). Two causes, both fixed. The component is still in the collection for a moment after a
Remove, so a projection in that window finds a component whose Designer will not answer, which is
indistinguishable from the first-touch flake this document already records - so the walk asks
ONCE MORE before believing it, and a form that answers the same way twice has its tab collected.
And the removal itself now collects the tab rather than leaving the next request to find out.

**Selecting a control on a page that was not open was refused.** The act asked the CANVAS whether
the control was there, and a control on a closed page is real and undrawn - so the refusal meant
for a deleted control caught a living one, against a rule this document states two sections up.
It asks the PROJECTION now.

**The Properties panel kept a control that had been removed** - name, geometry, font and all -
because a removal re-projected the tab and nothing told the panel. The panel's own publish
already knew what to do when its target is gone; nothing was asking it.

**An empty document squiggled twice.** The tolerant pass retires a refused line and parses again;
retiring a line that is already empty makes no progress, so the same finding arrived twice. No
progress now means stop.

**A name MSForms will not take answered `error 800a9c6c` and nothing else** - measured by adding
a control called `_Leading`. A control's name is a VBA identifier, and that rule is now a squiggle
on the line before anything is written, with the same sentence on the model's refusal behind it.

**A folder passed as a picture said "there is no file at ..."**, which is true and unhelpful when
the folder is sitting right there. It says it is a folder.

## The sited baseline: the path chosen, and how far it got

**The owner chose it 2026-08-16**: read the storage Excel has ALREADY saved, out of band. No
workbook is touched, nothing is exported during normal use, and the answer is exactly what Excel
wrote rather than an approximation of it.

What makes the path work at all is that **MSForms records only NON-DEFAULT properties**, as a
PropMask bitfield. So the mask alone answers the question this whole item asks - which properties
did the developer change - and the VALUES can still come from the live object model. Decoding
values is unnecessary; decoding the mask is the job.

**The format is documented, which retired the whole problem.** The spike stopped at the per-site
record layout and called it days of reverse engineering. It is [MS-OFORMS], published by Microsoft
under the Open Specifications programme, and it is on the open web: the record is
`Version(2) cbSite(2) PropMask(4) DataBlock ExtraDataBlock`, so the mask sits four bytes into
every site and the two mystery bytes before each name were `cbSite`. Reading the spec took an
evening and replaced every guess with a citation.

**The reader is `tools\harness\saved-design.mjs`** and it goes the whole way: the `.xlsm` is a ZIP,
`xl/vbaProject.bin` is a compound file, `/<Form>/f` is a FormControl carrying the site array,
`/<Form>/o` holds the controls' own blocks in site order, and a container gets its own storage
named `i` plus its site ID - `Options` has ID 6 and lives in `i06`, `Page2` has ID 13 and lives in
`i10/i13`. No export, no COM, no Excel.

**Four independent things agree, which is what makes it believable.** Every site record consumes
exactly its own `cbSite`. The sum of every `ObjectStreamSize` equals the `o` stream to the byte -
394,080 on the fixture. All fifteen `ClsidCacheIndex` values name the kind the fixture actually
put there. And the masks decode to what the fixture was built with: a Label to `Caption`, the OK
button to `Caption, PicturePosition, Picture`, the Image to `PictureSizeMode, Picture`, the
ScrollBar and SpinButton to `Orientation`.

**A SET BIT DOES NOT MEAN THE DEVELOPER CHANGED IT**, and that is the finding that decides how
this gets used. The mask says a property differs from the **file format** default. Where a control
KIND is born with something other than that default, the bit is set on controls nobody touched.
Measured: every control on the fixture carries `FontName`, because the form is Tahoma and the
file's default is MS Sans Serif - the bytes say so plainly, `mask 0x35` and the word `Tahoma`
sitting in the block. And every CheckBox, OptionButton and ToggleButton carries `BackColor` and
`ForeColor`, though `form-plan.mjs` sets nothing but a `Caption` on any of them.

So the baseline does not replace the walk's comparison against a bare coclass. **It narrows it.**
The mask is the short list of properties that could possibly be non-default, read once per form
and cached; the walk asks only those and compares as it already does, which is what filters the
file-format noise back out. Reading fifty properties of every control on every projection is the
cost that comparison was avoiding, and this is how it stops having to.

**One thing the mask cannot say.** `VariousPropertyBits` is a single bit over a packed field
holding Enabled, Locked, Visible, AutoSize, WordWrap and more. Set means one of them changed, not
which. Telling them apart means decoding the DataBlock rather than the mask, so it is reported as
itself and those properties keep the comparison they have.

**THE PROJECTION ASKS THE NARROWED LIST, 2026-08-17, and that closes this item.** The reader is
`Core.Forms.SavedDesign` over `Core.Forms.CompoundFile` - in Core rather than the shim because it
reads a FILE and has no COM in it, which is what gives it tests that run in 72ms with no Excel on
the machine. The walk carries a dotted control path (`Options.PickGround`) because the storage
nests the way the form does, and each name the mask offers is looked up in the kind's vocabulary -
which has already dropped identity, geometry, caption, the runtime-only members and anything
object-valued - read off the live control, and printed only if it ALSO differs from what the kind
is born with. Measured on a form built for the purpose:

```text
Form ProjProbeB "UserForm1" size 240x180
    TextBox Box at 8,8 size 80x18
        MaxLength = 12
        SpecialEffect = 0
    CommandButton Go "Go" at 8,40 size 60x20
        MousePointer = 3
```

Seven names are taken from the mask and then deliberately dropped: `BackColor` and `ForeColor`
have a better comparison already (against what the form passes down as well as the bare kind),
the fonts are inherited the same way and would otherwise print under every control of a form that
is not MS Sans Serif, a picture is binary and rides its own face, and the packed fields name many
properties with one bit and cannot say which.

**What it costs.** A markup print is **6.8ms** against the 8ms this walk cost before, so with the
baseline cached it costs nothing measurable. A save invalidates the cache by write time, and the
first print after one is **11.4ms** - about 4.6ms to re-read the workbook's storage, paid once per
save rather than once per projection. An enum rides as its number, because the dialect has four
value kinds and none of them is an enum name; the panel spells the member, the document carries
what round-trips.

A form that has never been saved has no baseline in the file and keeps the bare-coclass answer,
which is honest and stays.

## What a sweep for dead and convoluted code found

**2026-08-16**, after the hunt and the walk. Fast growth leaves fallbacks nothing reaches, and the
designer grew fast: the view was past four thousand lines.

**A fallback that could never be reached.** A MultiPage's tab band took `FRAME_INSET_TOP` - an
honest ten-point inset for a model that declines to answer - and then clamped the result to a
twenty-pixel floor two lines later. Ten is less than twenty, so the fallback branch and the
measured branch had one outcome and the code read as though they had two. The constant is gone
and the floor is the starting value.

**A load-bearing branch disguised as a fallback.** `roomOfContainer` had three: the form, the
container's own client, and a `closest(".dc-page-body")` off a name lookup. The third was not
redundant, which is the part worth knowing - since a page body started carrying its own name, the
SECOND branch finds the body and then searches inside it for a client, finds nothing, and the
third quietly does every page's work. Two cases now, naming which is which: a Frame's client is
inside it, a Page IS its client.

**A module that was hiding in the view.** The picture painters - the two placement families, the
alignment table, the layer type - hold no view state at all: they take a picture and an element.
That is a boundary, so they are `formpicture.ts` now, beside `colourpicker.ts`,
`formmarkuplang.ts` and `taborderdialog.ts`. A hundred lines out of the view, and the ones that
were easiest to reason about on their own.

**A dead overload from this session's own first cut.** `DataUriOf` took an owner and a property
name; every caller ended up holding the picture object instead, because it wants the kind as well
as the pixels. The convenience overload was never called from outside its own file.

**And a tolerance that did not scale.** The z-order row compares pixel counts before and after,
within 200 - a rule written when the labels painted about 8,000 pixels. They paint 18,000 now, so
a 1.5% difference failed a check meant to allow 2.5%. It compares within a few PERCENT now, which
is what "the counts swapped" actually means when the instrument is a photograph.

What the sweep deliberately did NOT do: strip the `export` off some fifty types and constants
that are only used inside their own file. It is over-export rather than dead code, the spec files
reach for several of them, and churning fifty declarations to save nothing is how a sweep becomes
a risk.

## Liveness beyond the funnel

**Landed 2026-08-16.** Every designer mutation this product makes re-projects the open tab: an api
`set`, a Properties panel edit and a canvas gesture all funnel through one place, which is what
keeps the document and the visual current without anyone re-activating the tab. What goes round
all of it is an edit made OUTSIDE - in the native designer underneath, or by a sync import, which
applies markup straight at the form.

Nothing announces one. A form has no revision counter the way a code pane does, and MSForms raises
no event for a control being moved. So it is ASKED rather than heard, and the whole design is in
making the question cheap enough and rare enough to be free.

**RARE: it rides the window events that already fire.** The pane tracker refreshes on a window
appearing or going and explicitly not on one moving - a frame resize fires thousands of moves and
a control dragged inside a form fires none, because MSForms draws its controls windowless. Those
same appear-and-go events bracket a native designer session: the window is shown, edited, and
hidden or destroyed, so the closing bracket is where an outside edit is caught. A floor of half a
second between checks collapses the bursts (a tooltip appearing and dying is two events, a menu
opening several).

**CHEAP: the key is each control's name and its four bounds, and nothing else.** Measured on the
fixture form: **49 wrappers and 3ms** against a full projection's 120 and 6ms. So a check that
finds nothing - which is almost all of them - costs 40% of the walk it avoids, and a check that
finds something pays for the projection it was right to make.

What it catches is what a hand does in a designer: add, remove, rename, move, resize. What it does
NOT catch is a property changed without any of those - a colour typed into the native Properties
window - and that is a stated limit rather than an oversight, because the alternative is reading
every property of every control twice a second.

The ledger is level across it: 1,225 wrappers over 25 checks, all given back.

`designer?action=liveness` is the read side, and it exists because everything else in this surface
funnels through a re-projection: without it no probe could drive the path a native edit takes. It
runs the check with the floor stood down and answers which tabs it re-projected. The suite uses a
SYNC IMPORT as its outside edit, which is a real one a developer can make - and the same shape as
the native designer's.

## What it costs, measured

The owner felt a lag in two gestures and asked for a walk of the whole surface
(2026-08-15). `tools\harness\designer-perf.mjs` is that walk, kept: every interaction timed
five times from the act to the moment the surface has ANSWERED - the canvas redrawn, the
panel following, the form itself carrying it - against the eighteen-control fixture form.

**Walked again 2026-08-16**, after this session's slices, against the fixture as it now stands:
nineteen controls, two of them wearing pictures. Every number below is that run rather than
August's, and the walk grew the gestures that landed since - the container tabs, the group, the
depth, the zoom, the liveness check and a picture load.

| interaction | median | worst |
| --- | --- | --- |
| designer route | 7ms | 8ms |
| markup route | 10ms | 11ms |
| open the designer tab | 14ms | 14ms |
| select a control, panel follows | 8ms | 12ms |
| nudge by an arrow key | 10ms | 11ms |
| drag a control | 10ms | 12ms |
| resize by a handle | 11ms | 12ms |
| toolbox drop | 19ms | 23ms |
| delete | 13ms | 20ms |
| open a page on a MultiPage | 9ms | 11ms |
| the tab strip's menu | 3ms | 4ms |
| marquee over the form's ground | 8ms | 9ms |
| align the group | 5ms | 8ms |
| bring to front | 7ms | 9ms |
| zoom the canvas | 4ms | 4ms |
| the liveness check | 3ms | 5ms |
| load a picture onto a control | 6ms | 8ms |
| typed edit to draft (debounced) | 358ms | 360ms |
| Ctrl+S: apply and save | 54ms | 59ms |

Everything a hand does is single-digit or low-double-digit milliseconds, and the two three-digit
numbers are both meant. The typed edit is the DEBOUNCE, paid on purpose: a keystroke waits 350ms
before the canvas follows, where a gesture does not wait at all. The save is Excel writing the
workbook, and it has more than halved since August (128ms then, 54ms now) - the host-thread
marshal samples read 15-16ms with occasional 47-62ms, which is also the floor under every number
above: one hop to the host thread costs about a frame, and the designer's own work disappears
underneath it.

The COM side, which the risks table asks for by name: a `designer` read costs **120 wrappers and
6ms**, the markup print **115 and 8ms**, the liveness key **63 and 6ms** - a little under half a
projection, which is the point of it - and an idle tick **0 wrappers**. The ledger stayed level
across 50,361 wrappers.

And what the PICTURES put on the wire: **683KB of base64 per projection** for the fixture's two,
which costs 2ms on a full gesture loop (measured against the same loop with none). That is why
there is no cache and no downscaling - both would be machinery bought for a cost that is not
there.

**One measurement was lying and is fixed.** The typed-edit line read 6-7ms for four of its five
rounds, because the canvas was already previewing a draft when the round started, so the wait it
was timing was already over. Every round starts from a clean document now, and the line says what
it should have said all along: 358ms, the debounce, by design.

**And the rest of the surface, since the ask was the whole of it** - `tools\harness\surface-perf.mjs`,
the same shape, against DebugFixture.xlsm:

| interaction | median | worst |
| --- | --- | --- |
| open a module tab | 70ms | 83ms |
| close a module tab | 37ms | 47ms |
| switch between two tabs | 18ms | 19ms |
| read a module's text | 2ms | 3ms |
| completions where a hand asks | 16ms | 17ms |
| hover | 16ms | 17ms |
| the project tree | 13ms | 13ms |
| a module-scope search | 6ms | 7ms |
| activate, and the panel follows | 20ms | 34ms |
| an idle snapshot | 3ms | 4ms |

Nothing here reads three digits either. The highest is opening a module tab at 70ms, which is the
native pane being created, its text published and monaco taking a model - consistent with the
host-thread hops it makes at 15-16ms each, and the same order as a tab open in any editor. The
PANEL costs about two milliseconds of the twenty beside it: the activate and the panel following
it measure 20ms where the switch alone measures 18, which is a fair price for walking a
component's whole property bag through its type library. Completions and hover are 16-17ms with
the analyzer in another process, and the page's own main thread stalled once, for 63ms, at
start-up - one stall in a whole session of driving.

Two things came out of the earlier walk rather than out of the clock. The drop and the delete were
paying that same typing debounce (347ms and 348ms, now 19ms and 13ms - see M5 above). And timing a
surface means touching every part of it, which is how the handles were found to be
unreachable wherever a neighbouring control overlapped them: a TextBox whose right edge met
a ScrollBar could not be resized from that side at all, because the selection overlay is a
sibling of the control it dresses and the neighbour painted over it. The overlay is above
the controls now, with a suite row on the adjacency that found it.

## Risks worth respecting

- ~~**Pictures.**~~ **Closed 2026-08-16.** No stream and no temporary file: the picture's GDI
  handle plus `GetDIBits`, or `DrawIconEx` onto a DIB section for an icon. The write took two
  loaders rather than one, because OLE's own refuses PNG. What still renders as honest bounds is
  a METAFILE - a drawing rather than pixels, with nothing to read out of it here - and any
  picture past four megapixels, which is a photograph somebody pasted rather than a form's
  decoration. Both answer null and the canvas draws the box, which is the architecture's rule.
  Cost measured: 700KB of base64 on the wire adds 2ms to a full gesture loop.
- **Third-party ActiveX.** Render bounds and identity, never guess an appearance; Additional
  Controls stays suppressed until add-by-ProgID is proven against at least one real
  third-party control. **Still suppressed, and now for a measured reason (2026-08-16).** Twelve
  real ProgIDs were tried against the fixture form and every one was refused, in two distinct
  ways. `MSComctlLib.ProgressBar.2`, `MSComCtl2.DTPicker.2`, `MSCAL.Calendar.7` and
  `MSForms.HTML:Text.1` answered `Invalid class string` - not registered on this machine at all.
  `MSComctlLib.TreeCtrl.2`, `MSComctlLib.ListViewCtrl.2`, `MSComctlLib.Slider.2`,
  `MSComctlLib.ProgCtrl.2`, `Shell.Explorer.2` and `RefEdit.Ctrl` all RESOLVE and all answered
  `TRUST_E_SUBJECT_NOT_TRUSTED`. That is Office refusing rather than MSForms or this product:
  `HKCU\Software\Microsoft\Office\Common\Security\DisableAllActiveX` is 1, the Trust Center's
  "Disable all controls without notification", and while it is on nothing in a form can create a
  non-MSForms control. So the condition this risk sets cannot be met on this machine without
  changing a security setting that belongs to whoever owns the machine - **and the owner's answer,
  given the measurement, is that third-party controls are out of scope for now** (2026-08-16).
  Additional Controls stays suppressed by DECISION rather than by blocker, and the toolbox offers
  the fifteen standard kinds it can create. The dialect keeps what it already had: a foreign
  control prints under its raw type name and can be moved and set, because a form that arrives
  holding one must not become unreadable.

  What the slice DID land is the refusal saying where the refusal came from: "Office is refusing
  to create ActiveX controls at all, which is the Trust Center's ActiveX Settings rather than
  anything about this control", and "no control of that ProgID is registered on this machine" for
  the other. Without that the message reads as a defect in xlide and the next hour goes into the
  wrong place. The suite's row asserts whichever truth the machine it runs on has: a control that
  arrives must walk under its own type, a control refused must be refused in those words.
- **COM volume.** Every property is a crossing; the designer's read path gets its own
  counter beside `publishUs` and `hostReadMs` before anyone asserts it is cheap.
  **Re-measured 2026-08-16** on the nineteen-control fixture, two of its controls wearing
  pictures: a `designer` read costs **120 wrappers and 6ms**, the markup print **115 and 8ms**,
  the liveness key **63 and 6ms**, an idle tick **0**. The ledger stayed level across 50,361
  wrappers. The August figures below were taken on the eighteen-control form before pictures and
  before this session's slices, and are kept because they are the baseline those were measured
  against.

  **Measured 2026-08-15** by `tools\harness\designer-perf.mjs`, on the eighteen-control
  fixture form: a full projection costs **185 COM wrappers and 8ms**, the markup print
  152 and 6ms, an idle tick 0 and 2ms - about ten wrappers a control, which is the
  identity, the geometry, the colours and the font, and no surprises hiding in it. The
  ledger stayed level across the whole walk (`givenBack == disposed`). So the read path
  is cheap, and now it is cheap with a number.
- **The interop class of bug.** New interface, same rules: every wrapper counted, every
  release on the right thread, `com-leak.mjs` rows before features - the 16-byte VARIANT and
  the finalizer-thread FailFast were both found the expensive way.
- **Scope discipline.** The native designer keeps working untouched the whole time. Nothing
  unsuppresses, and no menu returns, until the thing behind it is true.
