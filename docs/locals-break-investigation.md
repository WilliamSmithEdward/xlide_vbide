# Locals in break mode: the investigation and the fix

The developer's report, 2026-08-05 evening: sitting at a breakpoint in a live session, the
title bar says `[break]`, the code pane highlights the stopped statement - and the Locals
panel says "Not stopped. Variables appear here in break mode." This document records how
that was run down and what it turned out to be, because the answer retired a day-old
architectural conclusion and explains a crash storm that had never been root-caused.

The short version: the buffer every accessibility property read was written into was
declared eight bytes too small, sized for 32-bit Windows in a 64-bit process. Release
builds happened to survive it; the Debug builds that Smart App Control forced the dev loop
onto that morning did not. One `Size = 16` became `Size = 24`, and everything downstream -
the "unsupported configuration" theory, the crashes, the abandoned feature - turned out to
be that one number wearing different costumes.

## The system under investigation

Locals is fed by the ghost palette (lesson 29): the native Locals window is floated through
the object model, made layered at alpha zero, parked off the virtual screen, and read by
handle through UI Automation. The editor faithfully feeds any window with a paintable
surface, so the invisible ghost tracks every break and step while the themed panel renders
what the reader reads. The reader parses each list item's accessible name - "Expression
counter Value 1 Type Long" - back into columns. This shipped and passed its standing probe
(`Test-GhostLocalsPanel.ps1`, three rows at entry, values tracking steps) on 2026-08-04.

On 2026-08-05 it died, and stayed dead through twenty-four sessions.

## The evidence trail

### 1. The log said the reader was running and getting nothing

The live session's log, during the reported break:

```text
17:57:38.071 [verb] [host] locals: skipped 4 unreadable element(s)
17:57:38.071 [verb] [host] page <- {"type":"setLocals","context":null,"rows":[]}
17:57:38.071 [info] [host] locals: 0 row(s) at (no context)
17:57:38.221 [verb] [host] debug: mode 1 (break)
```

Break mode was detected fine (the code pane's highlight comes from `VBProject.Mode` and
`GetSelection`, a different path). The reader ran every poll tick and every payload-bearing
element "faulted natively" when its properties were fetched. At v0.1.0 the first such fault
latched the reader off for the whole session; the day's working copy had already replaced
the latch with per-element isolation and backoff - which made the failure visible instead
of silent, and proved resilience alone could not put rows on the panel. No session log on
the machine contained a single non-zero read all day.

The panel's message was a second, smaller defect: the page rendered "Not stopped" for any
empty locals set, because the message shape could not distinguish "not stopped" from
"stopped but nothing readable". Two bugs, one screenshot.

### 2. Out-of-process reads worked perfectly - so the ghost feed was healthy

The decisive split: while the in-process reader was dying, a PowerShell UI Automation
client (a different process) dumped the same ghost window mid-break:

```text
[0] ControlType.List      name=Locals
[2] ControlType.Pane      name=VBAProject.BreakProbe.BreakHere
[4] ControlType.ListItem  name=Expression BreakProbe Value  Type BreakProbe/BreakProbe
[5] ControlType.ListItem  name=Expression counter Value 1 Type Long
[6] ControlType.ListItem  name=Expression label Value "alpha" Type String
```

Everything was there, in exactly the parse format, at the same moment the in-process read
died at its first property fetch. The window, the float, the alpha-zero feed - all healthy.
The fault lived somewhere between the shim and the same data. (This dump also answered the
old "context reads null" nit in passing: the context box is a bare PANE to the
accessibility tree, not the Edit control the reader was filtering for.)

### 3. Raw MSAA walks the window but the rows are not in it

If in-process UIA was cursed, perhaps the older, simpler interface would serve: MSAA
(`AccessibleObjectFromWindow` + `IAccessible`), the VBE's native accessibility surface,
called same-thread with no client machinery at all. A full MSAA tree walk of the ghost
succeeded - title bars, scrollbars, the context box's text, the Call Stack button - and
contained NO variable rows anywhere. The grid windows expose only chrome through MSAA.

That is a load-bearing negative: the rows are served by a native UI Automation provider
only. The UIA channel is the only channel. No MSAA rewrite, no text scraping, no
alternative transport exists in-process or out. Whatever was wrong with the UIA reads had
to be fixed, not routed around.

### 4. The thread hypothesis: right shape, not the crash

The reader ran on the host thread - which IS the editor's thread, meaning the UIA client
was calling into a provider served by its own thread, a configuration the accessibility
framework has never supported. Out-of-process clients work because their requests are
serviced while the editor's thread pumps messages, and during a break it pumps constantly;
that is what makes the debugger interactive. So the readers moved to a dedicated MTA
thread (`GhostReaderThread`): the host thread only asks and looks, never waits.

Built, deployed, probed: the reads STILL skipped four unreadable elements - now tagged
`[t7]`. The thread was the right shape (and stayed), but the fault followed the reads
across threads. Something more fundamental was wrong.

### 5. The one number

With "which thread" and "which API" eliminated, what remained was the call itself:

```csharp
[StructLayout(LayoutKind.Sequential, Size = 16)]   // "a variant is sixteen bytes on x64"
internal struct UiVariant { ... }

[PreserveSig] int GetCurrentPropertyValue(int propertyId, out UiVariant value);
```

Sixteen bytes is the x86 VARIANT. On x64 a VARIANT is TWENTY-FOUR bytes: eight of type tag
and reserved words, then a sixteen-byte data union whose widest member is a record's two
pointers. The callee initialises the whole variant it is handed. Every property read
therefore wrote eight bytes past the buffer, into whatever the caller kept beside it on
the stack.

That explains every observed shape of the failure at once:

- **Why it worked on 2026-08-04 and died on 2026-08-05.** Release stack layouts happened
  to leave dead space in the overhang. Smart App Control began blocking fresh Release
  binaries mid-morning on the 5th, the dev loop switched to Debug publishes, Debug laid
  out a live slot where the overwrite lands, and every session from then on failed.
- **Why the symptoms were protean.** "Element 6's Name, persistently" in one session,
  "skipped 4" in another, a harmless caught fault here, Excel dying outright there - stack
  layout roulette, re-rolled by every code change, inlining decision, and thread.
- **Why the exception named nothing.** A `NullReferenceException` with no meaningful
  frames is memory corruption wearing an innocent exception's clothes: the overwrite zeroes
  a neighbouring slot, and whatever dereferences it next takes the blame.
- **Why out-of-process always worked.** The PowerShell client is .NET's own UIA wrapper,
  which allocates the real twenty-four bytes.
- **Why the Immediate mirror never broke.** Its reader uses the TextPattern's `GetText` -
  a BSTR out-parameter, no variant anywhere on the path.

Both structs (`UiVariant` and the placeholder `ComVariantBlock`, which is also written
through out-parameters) are `Size = 24` now.

## What shipped

1. **The root-cause fix**: 24-byte variants in `Interop/UiAutomation.cs`.
2. **The reading thread** (`Editor/GhostReaderThread.cs`): both readers live on one MTA
   thread; the host requests reads and consumes published snapshots, never blocks
   (bounded 500ms join on dispose), and clears readings at break exit so a new break
   starts empty rather than showing the previous break's values.
3. **Armor**: per-element fault isolation, stage-named first-failure logging, five-second
   backoff with recovery lines - a fault costs one element or one read, never the session.
4. **Honesty**: `setLocals` carries a `stopped` flag end to end (shim record, surface,
   bridge, shell, demo). A break with nothing readable says "No variables to show." - the
   panel can no longer claim "Not stopped" while the editor sits at a breakpoint.
5. **The context strip**: the reader accepts a dotted pane name (the context box is a
   pane, not an edit) and normalises empty to null so the strip hides rather than showing
   blank. In-process the pane's name currently reads empty even though outside clients see
   the full path, so the strip stays hidden for now - recorded as a known limit.

## Verification

Quiet-break probe (nothing external touches the editor while it sits stopped):

```text
18:41:16.752 [verb] [host] debug: mode 1 (break)
18:41:16.753 [info] [host] locals: 0 row(s) at (no context)     <- honest empty, first tick
18:41:17.757 [info] [host] locals: 3 row(s) at (no context)     <- rows, one tick later
18:41:24.763 [verb] [host] debug: mode 1 (break)                <- break still alive
```

`Test-GhostLocalsPanel.ps1` (breaks, steps twice, resets): PASS - three distinct row
pushes whose payloads track the steps (counter 1 to 2, label alpha to beta), first rows
170ms after break entry, one clean clear at exit. The build's test gate passed 72/72.

## The bigger meaning

The crash storm of 2026-08-05 - which forced the reset to v0.1.0 and was never
root-caused - is this same bug. The day's experiments (same-thread UIA, LegacyIAccessible,
an MTA worker that "crashed Excel on the first read") all called the same undersized-buffer
read path, and their differing fates were layout roulette, not evidence about the
provider. The conclusion drawn from them - that the provider is unsafe from anywhere
inside the process and only an out-of-process helper could ever read it - is retired. The
in-process reader works, and no second process is needed.

The day's tooling (the CDP door, the xlide api and its routes) lives on branch
`post-v010-experiments`. Re-landing it starts by rebasing onto this fix, then goes commit
by commit with the developer testing between steps.

## Traps recorded for future probes

- Never dump a ghost palette's UIA from outside the process while the in-process reader
  is alive: measured twice, the collision can RESET the project mid-break (panes
  destroyed, mode back to design). Quiet-window probes that only read the shim log are
  the safe shape.
- Step Into (command 188) is enabled in design mode too - a "wait for break" loop gated
  on its enablement does not wait.
- Breaks entered by external command (OnTime) update on the one-second poll; the editor's
  own run/step commands arm the fast watch and update in under a quarter second.
- The Watch panel shares the whole pipeline, but its row parse is still unverified against
  a real watch - adding one needs the native Add Watch dialog, so the developer's first
  watch is the test.
