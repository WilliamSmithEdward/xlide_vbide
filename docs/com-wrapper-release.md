# The line that never released anything

**2026-08-08.** For the whole life of this product, every COM wrapper it took kept its reference
and waited for the finalizer thread. The release was one line, it looked correct, three separate
investigations read past it, and it had never done anything at all.

This is the write-up because the defect killed Excel at least five times across two days, was
blamed on three different libraries, survived two fixes that were each real and neither sufficient,
and was invisible to a leak suite built specifically to catch it.

---

## 1. The line

```csharp
public static void GiveBackWrapper(object? wrapper)
{
    if (wrapper is null) { return; }

    (wrapper as IDisposable)?.Dispose();     // <- never released anything
    Interlocked.Increment(ref _wrappersGivenBack);
}
```

`ComRuntime.Wrappers` is a `StrategyBasedComWrappers`, and the object it hands back from
`GetOrCreateObjectForComInstance` is a **`ComObject`**.

`ComObject` does not implement `IDisposable`.

So the cast produced null, `?.` swallowed it, and the line was a no-op. The counter on the next
line incremented regardless.

The correct call is `ComObject.FinalRelease()`.

## 2. Why it ends the host

A wrapper built with `CreateObjectFlags.UniqueInstance` takes **its own reference** on top of the
one the caller holds, and is never cached, so nothing else will ever give it back. Left
unreleased, it is released when the collector eventually reaches it - on the **finalizer thread**.

The editor's objects are apartment-threaded. Releasing one from a thread that is not its apartment
is an access violation, and in an ahead-of-time compiled image that exception cannot be thrown:
the runtime calls `FailFast`, and the host process ends. It also corrupts COM's own bookkeeping on
the way out, which is why the same defect was reported against `ntdll` twice and `VBE7.DLL` twice
as readily as against this library.

The stack is always the same and always says nothing about the cause:

```
Marshal.Release
  -> FreeThreadedStrategy.Release
    -> ComObject.Finalize
      -> __Finalizer.DrainQueue
```

By the time it runs, whatever created the wrapper is minutes gone.

## 3. Why every instrument said it was fine

This is the part worth keeping.

**The counter agreed with itself.** `taken: 1895`, `givenBack: 1882`, `live: 13`, stable across
every operation, every session, all day. Correct arithmetic over a quantity that was not the one
that mattered: it counted the **intention** to release, because the increment sat next to the
disposal rather than being conditional on it.

That failure mode had already been diagnosed once in this codebase, in the previous version of the
same counter, and the fix was to make taking and giving back "the only two doors, each doing its
own counting" so that a caller "cannot dispose without counting, or count without disposing." That
made the two operations inseparable. It did not make either of them *work*.

**The leak sweep passed 37 rows.** `com-leak.mjs` runs 380 operations over every read route and
every state-changing route paired with its undo, and asserts that the live count returns to its
resting level and that handles do not grow. All of it true. All of it downstream of the same
broken number.

**Handle counts were flat.** They would be: the underlying COM object's refcount was still held,
so no handle was released and none needed to be. A leak that holds a reference looks exactly like
a healthy program to a handle counter.

**A forced finalizer drain survived, repeatedly.** After startup, after each of eight operation
groups, after the full sweep, after writing a 64,802-line module, and after five rounds of the
churn pattern that had killed it twice. Every one clean. The wrappers being drained were mostly
over objects that tolerated the release; the crash needed a particular object, on a particular
thread, at a particular moment.

**Two earlier fixes were real and neither was enough.** Both removed places where a wrapper was
created that nobody needed:

- `DispatchObject.Dispose` never called `GiveBackWrapper` at all (2026-08-07). Restoring the
  defect on purpose showed a single `project()` call leaking 441 wrappers.
- `VariantToString` asked the runtime to build a wrapper for an interface-valued variant, first in
  its fallback (2026-08-07) and again in its default branch, which the first fix did not cover
  because `VT_BYREF | VT_DISPATCH` is not `VT_DISPATCH` (2026-08-08).

Both lowered the rate of a crash whose mechanism was untouched. That is why the crashes got rarer
and never stopped, and why each fix looked plausible at the time.

## 4. What actually found it

Counting the two things separately.

```csharp
if (wrapper is IDisposable disposable) { disposable.Dispose(); released = true; }
else if (wrapper is ComObject comObject) { comObject.FinalRelease(); released = true; }

if (released) { Interlocked.Increment(ref _wrappersDisposed); }
```

and then asking:

```
taken 1895   givenBack 1882   disposed 0   live 13
```

`disposed: 0`. Nothing else needed to be said.

The route to that question was a different instrument built the same evening: `drainfinalizers`,
which forces a collection and waits for the finalizers so a leaked wrapper fails **now**, at the
operation that leaked it, instead of minutes later. It did not find this defect - every drain
survived - but it ruled out the entire operation surface fast enough to make it clear the fault
was not in *which* wrappers were taken. That left only *what happens when they are given back*.

## 5. Before and after, measured

Same session shape, two builds:

| | before | after |
| --- | --- | --- |
| wrappers given back | 1,882 | 1,882 |
| wrappers actually released | **0** | **1,882** |
| the churn pattern that killed Excel twice | died | survives, 3 rounds, ~20,000 wrappers |
| leak sweep | 36 rows pass (meaninglessly) | 37 rows pass, gap 0 |

## 6. The guard

`com-leak.mjs` now asserts that **given back and released are the same number**. They are counted
separately at the source, so they can only diverge if a wrapper is counted home without its
reference being released - which is the defect, exactly.

This is the check that would have caught it on day one, and none of the other 36 rows could have.

## 7. What to take from it

**An instrument that agrees with the code is not evidence.** Every check here was written against
what the code was understood to do. `(wrapper as IDisposable)?.Dispose()` *reads* as a disposal,
so the counter beside it was written to count disposals, so the suite was written to trust the
counter. Three layers of instrumentation, all derived from one unverified assumption, all
confirming each other.

**Count effects, not intentions.** The increment must be conditional on the release having
happened. If those two facts can be counted separately, count them separately and compare - the
comparison is free and it is the only thing that can catch this class of defect.

**`?.` on a cast is a silent branch.** `(x as T)?.M()` does nothing when `x` is not a `T`, says
nothing, and is indistinguishable from success at every call site. In a release path it is a leak
with no symptom. Prefer the explicit `is` with an `else` that complains - the `else` here now logs
loudly, and if a third kind of wrapper ever appears it will say so before it kills anything.

**A fix that lowers the rate of a crash is not a fix.** Both earlier fixes were correct, both were
verified by restoring the defect and watching a check fail, and both left the mechanism intact.
"It happens less now" should have been the signal to keep going rather than to close it.

**The oldest line is the least suspected.** This one predates every investigation that read past
it, which is precisely what made it invisible: it was part of the furniture, and each hunt started
from "what changed recently".

---

Related: [lessons.md](lessons.md) entries 36, 40, 49 and 51 are the earlier stages of this hunt,
kept because the wrong turns are the useful part. [driving-excel.md](driving-excel.md) covers the
COM wrapper counters and `drainfinalizers` as tools.
