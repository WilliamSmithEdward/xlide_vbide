# Telling two of the same thing apart

**2026-08-08.** Two unsaved workbooks are both called `VBAProject`. That single fact produced a
defect at every layer of this product that names a workbook, and fixing each one exposed the next.
Five bugs, one cause, found five times in an afternoon.

This is the strategy that came out of it, written down because the same shape will arrive again
the moment two of anything can share a name: two modules, two procedures, two references, two
Excel instances.

---

## 1. An identity and a label are different things

The whole family of bugs comes from one substitution: using the name a human reads as the key a
program looks things up by.

| | identity | label |
| --- | --- | --- |
| answers | which one is it | which one am I looking at |
| must be | unique, stable for the lifetime that matters | recognisable, short, honest |
| may be | ugly | ambiguous, if you say so |
| compared by | machines | nobody |
| changes when | the thing is replaced | anything about presentation |

A workbook's identity here is its file path, lowercased. Its label is the file name. Those are
close enough to each other that for a year nothing noticed they were two ideas, and then a
workbook arrived with neither.

**An unsaved workbook has no file name.** Ask for one and the VBE raises rather than answering
empty, so the code fell back to the project's own name, and every new workbook's project is called
`VBAProject`. Two of them and the fallback returns the same string twice.

## 2. What made it dangerous was the keying

The label leaking into a lookup key is the part that corrupts rather than confuses.

```
liveKey(projectId, moduleName)
```

With both workbooks answering `VBAProject`, one workbook's `Sheet1` **was** the other's to every
map in the engine and the shim. The engine seeded one over the other. The pass compared the wrong
workbook's sources and skipped analysis it should have run. Findings computed against one
workbook's code decorated the other's.

None of that shows as an error. It shows as an editor that is quietly wrong about somebody's code.

**The rule this gives:** a key must never be derived from something whose job is to be read. If
the only available name is one a human chose, or a default the platform hands out, it is a label
and needs an identity found elsewhere.

## 3. Where to find an identity when the obvious one is missing

In order of preference, and this order matters:

1. **Something the platform guarantees unique and stable.** A file path. A GUID the object was
   created with. Use it even when it is ugly, because nothing reads it.
2. **The object's own COM identity**, its canonical `IUnknown`. This is what COM means by
   identity: two references to one object always answer the same pointer, two different objects
   never do. Query for `IID_IUnknown` specifically; a QueryInterface for any other interface may
   hand back a different pointer for the same object, and comparing those is a bug that looks like
   it works.
3. **A synthetic id you assign and hold.** A counter, stored beside the thing. Correct, but you
   now own its lifetime, and getting that wrong is a leak or a stale key.
4. **Never** the type name, the default name, or anything the user can change.

We used (2). An unsaved project is `vbaproject#24710821458`. It is unique among the projects alive
at once and stable for exactly as long as the project is, which is the lifetime the id has to
cover. It changes if the workbook is closed and reopened, and that is right: that is a different
project.

**Verify stability before trusting it.** Read the identity three times in a row and compare. Ours
was stable; had it not been, every map keyed by it would have leaked an entry per read.

## 4. If you synthesise a label, you have to teach every surface

The tree numbers a shared name: `VBAProject 01`, `VBAProject 02`. A name that is already unique is
left alone, because `Book1.xlsm 01` is noise.

That numbering broke clicking **within minutes**, and then broke four more things, each in a
different way:

| what broke | why |
| --- | --- |
| resolving a click | no project answers to `VBAProject 01`, so every lookup fell through to "the shown project" and the click landed on whatever was already open |
| the tab strip | it labelled tabs with the RAW name while the tree used the numbered one, so the strip published `projects: ["VBAProject"]` beside `activeProject: "VBAProject 01"`, nothing matched, and activation silently did nothing |
| the reverse lookup | id to display returned the raw id, so an unsaved project would have shown `vbaproject#24710821458` to a developer |
| the tree selection | rows were matched by NAME alone, so clicking one `ThisWorkbook` lit every workbook's `ThisWorkbook` |
| the api | `projects` still reported both as `VBAProject`, so nothing driving the api could name the one it meant |

Every one of those is the same mistake: a second place that decided for itself what a workbook is
called.

**The rule:** a synthesised label must be produced in ONE place, stored against the identity, and
read from there by everything else. Ours lives in `_projectNames`, id to shown name, built where
the tree is built. `DisplayFromProjectId` and `FindProjectByDisplayName` are the only two doors,
and they are inverses of each other.

**The tell that you have missed one:** two surfaces in the same window disagree. The strip saying
`VBAProject` beside `activeProject: VBAProject 01` was visible in the message tap the whole time
and named the bug exactly.

## 5. A pair identifies a child, not a name

A module is not `Sheet1`. It is `Sheet1` **in** a workbook. The tree highlighted by name alone, so
one click lit the same-named row in every workbook.

Anywhere a child is looked up, selected, compared or highlighted, the key is the pair. This is
easy to get right at the point where the parent is obvious and easy to get wrong three files away
where only the name was passed along.

**Look for functions that take a bare child name.** Each one is a place that will pick the first
match and be right most of the time.

## 6. How to test it

Ambiguity cases are not reached by accident, which is why all five of these were reported by a
person rather than caught by a check.

- **Make two of the thing, and make them ambiguous on purpose.** Two unsaved workbooks, not two
  saved ones with different names.
- **Drive each one independently and assert which one answered.** Not "a module opened" but "the
  module in THAT workbook opened". Ours needed `shownProject` in the assertion, and without it the
  test passed while every click went to the wrong place.
- **Give the harness a way to set the case up.** A case that cannot be set up will be reported by
  a user. `Workbooks.Add` through the Immediate window makes two unsaved workbooks in one line.
- **Use unique names per run in the test itself.** The freshness suite reused two fixed module
  names and inherited the previous run's answers under them. A suite must not depend on the thing
  it is testing being right.

## 7. The short version

- A label is not an identity, and the moment they differ you need both.
- Never key by anything whose purpose is to be read.
- Prefer an identity the platform guarantees; COM's canonical `IUnknown` is one, and it is
  available whenever you hold the object.
- Synthesise labels in one place, store them against the identity, and read them from there.
- A child is identified by a pair, never by its own name.
- Two surfaces disagreeing about a name is the symptom; look for the second place that decided.

---

Related: [lessons.md](lessons.md) entry 58 is the account of the five bugs as they were found, and
entry 57 is the second Excel instance, which is the same shape one level up: one shared resource
where there should have been one per process.
