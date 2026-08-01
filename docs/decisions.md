# Decisions

Choices that would be expensive to reverse, with the reasoning that produced them. Superseding a
decision means editing its entry, not contradicting it silently elsewhere.

## 1. The add-in is a native server, compiled ahead of time

Status: decided and proven.

A managed add-in normally loads a runtime into the host. That runtime is shared with whatever else
the host has already loaded, it costs start-up time on a path the user is waiting on, and it
enlarges the failure surface inside a process we do not own.

Compiling ahead of time to a native library removes all of that. The measured result is a 1.91 MB
library that the editor loads directly, with no runtime deployed beside it and none loaded into the
host.

The cost is real and accepted: no reflection-heavy libraries, no runtime code generation, no
framework that assumes a runtime, and a C++ toolchain required to build. Every one of those is a
constraint on us rather than on the user.

## 2. Analysis runs outside the host process

Status: decided.

The editor is single threaded and owns the user interface thread. Analysis of a large project is
measured in seconds on a first pass, and a previous generation of this analyzer needed two separate
rounds of optimisation before large real-world modules were usable at all. Running that work in the
host means either blocking the thread the user is typing on or maintaining a marshalling discipline
across every analysis path forever.

A separate process cannot block the host, cannot leak memory into it, and cannot crash it. It is
also reusable: the same engine already serves an editor extension, so analyzer improvements land in
both products.

The cost is a protocol boundary and process lifetime management.

## 3. The engine is reused rather than rewritten

Status: decided.

The analyzer exists, is validated against the real compiler across a corpus of accepted and rejected
cases, carries a rule set with recorded evidence per rule, and already meets per-keystroke latency
budgets on modules of tens of thousands of lines. Rewriting it in another language would take months
and start from behind on correctness.

It ships as a self-contained executable so the user needs nothing installed.

Reversing this later is cheap by construction: the protocol is the contract, so an engine in another
language can replace it without the add-in changing. Rewriting is a performance decision to make
against measurements, not in advance.

## 4. User interface is web-based, hosted in the editor's own windows

Status: decided, hosting in progress.

The editor can site an ActiveX control inside a native docked tool window. That is the documented
extension point and the only way to get a first-class docked panel rather than a floating window
that does not belong to the editor.

What goes inside that control is our choice, and a browser surface is the strongest one available.
It renders out of process, so interface crashes cannot take the host down. It gives a modern layout
and styling system, real accessibility, and a component ecosystem. Most importantly it makes the
editor surface possible at all, because a full editing component can be hosted in it.

The alternative, drawing everything with native controls or custom painting, makes every feature a
separate project. That is the trap that has historically limited what tools of this kind can offer.

## 5. Calls into the editor use late binding at the control plane

Status: decided.

Early binding to an automation interface requires the declared member order to match the type
library exactly. A mismatch does not raise an error; it calls the wrong function through the wrong
signature. Across host versions that is a memory-corruption bug waiting for a user we cannot debug.

Control-plane calls happen once per user action or editor event, so dispatch overhead is not
measurable. Paths that run per keystroke may use early binding, chosen against a measurement, with
the member order taken from the type library rather than from documentation.

## 6. Registration has one source of truth

Status: decided and tested.

Registration decides whether the add-in loads at all, and a wrong key produces silence rather than
an error. There are three consumers: the development script, the installer, and the tests. Any two
of them drifting apart produces a bug that reproduces only on a machine nobody is debugging on.

All three derive from one type. A test asserts the installer authoring matches it in both
directions, so an entry cannot be added to one without the other.

## 7. Install is per user, from a single executable we wrote

Status: decided. Replaces an earlier choice of a packaged installer format.

Writing class registration under the user hive needs no administrator rights, which removes the
single largest obstacle to someone trying the product. It is also the correct scope rather than a
reduced one, because the editor resolves class registration through that hive.

The installer is an ordinary program of ours, compiled ahead of time into one executable that
carries the product inside it. Three things follow that a packaging format does not give:

The registry layout has one definition, used by the product, the tests, and the installer. A
packaging format needs its own copy of that layout in its own language, which then has to be kept
in agreement by a test. Sharing the code removes the class of bug instead of detecting it.

Installation is verifiable the same way everything else is, by running it. There is no separate
toolchain to install before the installer can be built, which also keeps continuous integration
simple.

Self-update, which the product needs because nothing else will do it, is the same code path as
install rather than a second mechanism.

The trade accepted: no packaged-format deployment for administrators who require one, and no
built-in transactional rollback. Per-user installation of a development tool is the case that
matters, and the installation is small enough that its failure modes are enumerable. A packaged
format can be added later around the same payload if a real need appears.

## 8. Releases are signed

Status: decided, not yet implemented. Blocks public release; does not block development.

An unsigned installer triggers a Windows Security warning naming the executable as unpublishable
("we can't confirm who published xlide-setup.exe"), observed on this machine during the first
install round trip. Developers are the audience most likely to take that warning seriously, and
the uninstaller re-launches itself from the temporary folder, which is a pattern reputation
systems watch closely.

The plan is Authenticode signing of the installer, the shim, and the engine executable in the
release pipeline. Azure Trusted Signing is the current fit: subscription-based, no certificate
files to protect, and it accrues SmartScreen reputation. Local development builds stay unsigned;
nothing in the product may behave differently based on whether it is signed.

## 9. The integration harness owns its host instance

Status: decided and proven.

A harness that attaches to whatever host is already running will eventually drive, and then
terminate, the developer's own session. The harness therefore starts its own instance, confirms by
process identity that it is driving what it started, refuses to run when a host is already open
unless told otherwise, and terminates only the identity it recorded.

It also launches the host as a process rather than creating it through automation, because a host
created through automation does not load add-ins at all, and it restores state the host rewrites
after a failed load. Both are recorded in `lessons.md` with the evidence.
