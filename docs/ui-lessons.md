# Building this surface: what the UI taught us

The editor surface is a web page living inside a native host that was not designed to have
one. Most of what follows was learned the expensive way — a symptom that looked like one
thing and was another. This is the UI companion to [lessons.md](lessons.md), which carries
the host-side findings; where a lesson has a number there, it is cited.

The rule under all of it: **the page is not alone**. A native editor is underneath, a COM
object model is beside it, and both move things while the page is drawing. Every reflex a
web developer has about owning the DOM has to be re-earned here.

## The host moves under the page

**A rebuild can destroy the element a pointer is pressing.** The tab strip is rebuilt
whenever the set of open modules changes, and the host echoes that set freely — a focus
click alone produces one. A `click` needs its press and release on the same LIVE element, so
a press that triggers an echo means the click never happens: the close box read as dead
until a second try. The fix is to arm at `pointerdown`, fire at `pointerup`, and capture the
press identity as DATA rather than as an element reference, so a rebuilt twin still
satisfies the release. This bug had two lives — the first was a stale drag flag — and both
looked like "the X sometimes doesn't work".

**Rebuild only when something drawn has changed.** The cheapest defence against the above is
not rebuilding: every strip keeps a render key covering everything it draws (identity,
order, active, badge counts, dirty flags), and an echo that changes nothing rebuilds
nothing.

**A pointer stream can end in `pointercancel`, and often does.** The host steals focus
mid-gesture on all sorts of occasions — showing a module, following a navigation, its own
window management. A drag that only handles `pointerup` leaves its suppression flag raised
forever, and the next click is swallowed. Every drag here handles `pointercancel`, and the
window-level listeners are what keep a fast drag alive after it leaves the element.

**A drag must also end when the WINDOW does.** Alt-tab, a screenshot tool, the host taking
focus — none of them produce a pointer event. Without `blur`, `visibilitychange`, and
Escape handlers, the dim and the drop compass outlive the gesture and the surface sits
looking permanently mid-drag.

## One page-wide service, several editors

**`editor.addCommand` is not scoped to the editor you called it on.** Every standalone
Monaco editor on a page shares ONE keybinding service; the `when` clause is the only scoping
mechanism there is. Splitting into editor groups produced two identical Backspace rules,
both matched, the later registration won everywhere, and its handler deleted in ITS editor —
so Backspace in the group being typed in silently edited the other group. A context key
created per editor (`editor.createContextKey`), included in that editor's when-clauses, is
the fix. Anything bound per editor needs it. (lessons.md 34)

**The shape of the failure is the diagnosis.** The live page answered `prevented: true,
handled: false`: the key WAS claimed and the command DID run. A dead key would have been
`prevented: false`. That single distinction turned "Backspace is broken" into "the binding
points at the wrong editor" without a single guess.

**Language features answer for a model, not an editor.** Completion, hover, and signature
providers are registered per LANGUAGE and receive whichever model is being asked about. With
several models live, each provider must check the model it was handed against the one the
host's engine is mirroring, or a background group's text gets answers computed for the
active module's offsets.

## Models, not editors, are where the state lives

**Keep a model per open document and the rest becomes easy.** Attaching content listeners to
the MODEL rather than the editor means an edit is attributed to the document it changed,
whichever editor made it — which is what makes several editors possible at all. Undo stacks,
markers, decorations, and view state all hang off the model, so tab switching costs nothing
and a background module's squiggles keep updating.

**Decorations that describe host state belong on the model too.** The stopped line and the
breakpoint dots follow the host-active document; hung on an editor they would be wrong the
moment its model changed, and invisible in whichever other group showed that module.

**Two workbooks can hold a Module1.** Every identity — model URI, tab key, baseline key, view
state key — is the (workbook, module) PAIR. A name-only key is a latent corruption waiting
for the day both are open: on the host side, a line diff computed against the wrong module's
baseline writes a merge of two unrelated files.

**Monaco created in a detached container reports nonsense.** An editor built into a element
that is not yet in the document measures a few pixels and its automatic layout does not
recover on its own. One explicit `layout()` after attaching gives its observer truth to
track from.

## CSS in a page you do not fully own

**Generic class names collide with the bundle's.** A `.row` on a split container inherited
`align-items: baseline` from an unrelated rule in the bundled stylesheet and collapsed every
cell to its tab strip's height. The layout looked broken in a way that read as a flex bug in
our own code. Class names for structural things are prefixed now (`split-row`, not `row`),
and a structural property that MUST hold is stated explicitly rather than left to inherit.

The diagnosis is worth keeping too, because it generalises: walk `document.styleSheets`,
keep every rule whose selector the element matches, and read what each one says about the
property in question. Computed style tells you the answer; the rule list tells you who
decided it. That loop is now the debug api's `inspect?rules=1`, so the next collision costs
one request rather than an hour.

**A flex child does not grow unless told.** Moving the workspace from grid to flex left the
editor area with `grid-row` properties that meant nothing and no `flex`, so it sized to its
content — which, for a Monaco container, is nothing — and the bottom dock swallowed the
window. Dead properties after a layout change are worth grepping for; they fail silently.

**Elements that swap must share a cell.** The editor area and the empty-workspace view are
mutually exclusive and must occupy the same space, or swapping them moves everything else.

## Drag and drop that a person can aim

**Guessing intent from pointer position does not survive real geometry.** The first drop
implementation picked a zone from where the pointer sat in a region. Over a wide short
region — a bottom panel — "near the left edge" and "just left of centre" are a few pixels
apart in fractional terms, and the nearest-edge rule sent drops to the wrong section. Worse,
an if-chain that tests `x` before `y` sends a point near the top edge but left of centre to
the LEFT, which is exactly what it looks like from the outside: broken.

**So make the target explicit.** A compass of five zones over the region, and the pointer
must come to the zone it means. It is more motion, and it is the motion every studio asks
for, because it converts a guess into an aim. A release off every zone drops nothing, which
is also honest.

**Hit-test the compass by geometry, not by pointer events.** The dragged element holds the
pointer capture, so nothing else receives pointer events at all. The zones are rectangles
compared against the pointer position.

**The preview must describe the actual outcome.** Dropping on an editor edge where a section
already exists JOINS that section — so the preview outlines the section, not a half of the
editor the drop would never touch. Creating something that does not exist yet gets a dashed
edge, because the shape is a proposal rather than a place. A heavy translucent slab over the
editor reads as "your editor has been replaced", which is alarming for a gesture that has
not happened yet; a light wash with a definite edge says the same thing calmly.

**Offer only what the drop will honour.** A group's only tab cannot split against its own
group: the tab would leave, the group would dissolve, and the result is the same picture one
splitter wider — so the code refused it, correctly, while the compass went on showing the
zone. A lit zone that does nothing is indistinguishable from a bug, and it was reported as
one. The zones offered are computed per region: over a pane's own group, the centre is where
it already is, and the edges only exist if something would remain behind.

**Reordering wants the strip itself as the feedback.** Moving the dragged tab in the DOM as
the pointer crosses its neighbours' midpoints is clearer than any overlay, because the
answer to "where will this land" is the strip already showing it. Move the element rather
than re-rendering: the pointer capture keeping the drag alive belongs to that element and
dies with it.

## Persistence and identity of arrangement

**Membership is the host's; geography is the developer's.** Which modules are open, and
which panes exist, are answered by the object model and by the shell. Where each one sits,
which group holds it, and how big everything is are the developer's arrangement, and must
survive every host echo. Conflating them produces a layout that resets itself whenever the
host says anything.

**Layout is page-local state, not product settings.** It lives in `localStorage` beside the
splitter positions rather than in the settings file the host owns: it describes one screen's
arrangement, not a preference about behaviour.

**A stored layout must tolerate a changed product.** Panes that no longer exist are dropped;
panes the stored layout has never heard of are placed somewhere sensible rather than
vanishing. A layout that can strand a pane is a layout that loses one on the next release.

**A closable thing needs a route back.** Every pane has an X, so every pane needs a menu that
lists it — and one pane (the explorer) may not be closed at all, because with every tab shut
it is the only way back to a module. The menu says so rather than silently refusing.

## Talking to a host that reloads you

**A reloaded page has heard nothing.** The surface's "what changed" caches — last tab list,
last language facts, last chrome state — exist to spare a page that already has the picture.
A reload produces a second `ready` on the same session, and those caches then describe a
conversation the new page was never part of: republishing compared, matched, and sent
nothing, so the page came back with its models and no tabs. Every such memo must be
invalidated when the client restarts. (lessons.md 35)

**The traffic tap is how you see it.** The bug above was invisible in the page and invisible
in the log; the message tap showed every republish EXCEPT `setModules`, which named the
cache in one line.

## Probing a live surface

**A synthetic `KeyboardEvent` is enough to exercise a keybinding.** Monaco resolves bindings
from the event; only text INSERTION needs a trusted event. So a probe can press keys without
a focused window — but it must place the caret somewhere the key can do work. Backspace at
the start of an empty line is correctly a no-op, and reading that as a fault sent a probe
hunting a bug that was not there. A check that inserts its own text first asserts the key,
not the fixture.

**Synthetic pointer drags work if every handler on the path is synchronous.** They are, here,
deliberately — which also means the debug door can drive a whole drag in one request.

**Wait for the condition, not for a duration.** Probes that slept a guess after a split
raced the thing they were watching: a freshly created editor gets its input element a beat
after its model. Poll the actual predicate.

**The page's content policy exempts the browser's own evaluation, not its callbacks.** A
script injected by the host may call `eval` synchronously — the debug door's own script
runner relies on it — but the same call from inside a `setTimeout` callback is refused
("unsafe-eval is not an allowed source"). A polling waiter written the obvious way therefore
never ran its predicate and reported every condition unmet. Compile the string to a function
ONCE, during the synchronous evaluation, and call the function afterwards.

**A probe that rearranges persistent UI state must put it back.** This suite drags panes
around to test the drop compass, and the arrangement lives in `localStorage`. Clearing the
key at the end is not enough: the page still holds the rearranged layout in memory and
writes it back on the next render, so the developer opens the editor to find a pane docked
somewhere they never put it — reported, reasonably, as a bug. Clear AND reload, then assert
the default came back.

**Order matters in a suite that drives one live host.** A check that RESTARTS the page
belongs last: everything before it depends on host state built up in sequence, and a reload
in the middle left a later check watching an editor whose caret had been re-established
underneath it.

**Assert what the check is testing, not the weather around it.** A check that the log's wait
returns when a line ARRIVES watched for a line that a save only produces while modules are
open. That was true early in a fresh session and false after two suites had run against the
same host, so the check passed all day and then failed on a run where nothing had changed
but the order. Watching for the line the command itself writes tests the same thing and
depends on nothing.

**Look at it, do not only measure it.** Most of this surface was built by reading numbers —
rects, class lists, computed styles — because a screenshot of a whole editor frame is a big
picture in which a 54-pixel drop zone is invisible. Cropping the capture to one element
(`capture?selector=`) changes that, and the first thing it caught was itself: the crop landed
on the toolbar when asked for a pane header, because the surface's overlay window is not the
document area it is a child of — the surface is taller, since it draws the menu bar and
toolbar too. A picture of the wrong thing is at least obviously wrong; a number from the
wrong thing looks like data.

**Verify against the demo transport first.** The page runs its own loopback host in a plain
browser, with two documents and live tabs. Nearly every layout finding above was reproduced
there in seconds, then confirmed live — the browser tells you what the code does, and the
host tells you what the host does to it.
