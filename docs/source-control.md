# Source Control

A git repository behind the folder a project's modules are exported to, and a ninth dockable pane
that works it. **Designed 2026-09-08; the sections at the end record what shipped against it.**

Rubberduck's source control panel was its most requested feature and it never shipped working.
The pieces this product already has are the ones every earlier attempt lacked: an export that
round-trips a module faithfully, attribute header included; a form's design as text; a change log
that already models rounds and restore; and a Changes pane whose comparison renderer is shared with
the import dialog. Git turns the session-scoped log into something durable, branchable and shared.

The dependency is git.exe. The product does not bundle it: a MinGit has no credential manager, so it
could commit but never push, and bundling git plus a credential manager doubles the installer and
makes xlide responsible for shipping git's security updates. Git for Windows installs per user
without administrator rights, which is how xlide installs. Its Git Credential Manager is the whole
login story: the first push or pull opens the host's own sign-in in a browser, the token lands in
Windows Credential Manager, and xlide never sees a credential. The GitHub CLI was weighed and set
aside as the base, because it carries no git of its own and speaks only GitHub; it remains a
natural optional layer for Publish and pull requests later.

## What it looks like

A Source Control pane beside Tests and Changes. Its rows are the project's modules, live against
the branch head: modified, added, deleted, renamed. Tick rows, type a message, press Commit. The
message box comes pre-filled with the labels of the change-log rounds since the last commit, which
is where an agent's work becomes a reviewable commit. A row's click shows the comparison in the same
side-by-side renderer the Changes pane uses. Below the rows, the history: commits newest first, each
opening to the modules it touched, each module to its comparison against the live text, with Open
this version and Restore. A Blame toggle paints author, date and short hash at the end of every
committed line of the active module, and marks lines changed since the last commit as uncommitted.
The divider between the rows and the comparison drags, or takes the arrow keys: a third of the
pane until it is moved, then the width it was left at, kept across reloads, between a floor the
rows can still be read at and a ceiling that leaves the message box and the comparison their room.

The branch is a select. Changing it is a checkout followed by an import into the project, refused
while the workbook is dirty, the way git refuses a checkout over uncommitted work. Its last entry,
New branch..., asks for a name and cuts one from the head, staying on the same commit, so nothing
is imported and a dirty workbook is no bar. After a fetch, the branches only a remote has are
listed under the remote's name; picking one checks it out as a local branch of that name tracking
it. Fetch and Push are buttons that run git and report what it said. Pull is a checkout's shape:
refused while the workbook is dirty, then git, then an import of what it brought into every open
project in the repository - it has to import, because a save's true-up export writes the project's
text over the folder, and a pull left unimported would be undone by the next save. A pull that
stops on a conflict is the conflicted state, with Abort.

The branch head's row in the history carries Undo, the pane's amend: the branch goes back one
commit, the commit's changes return to the rows, and its message returns to the box to be edited
and committed again. Nothing is deleted - the folder keeps the text. It is greyed with the reason
when this pane cannot undo the head: the first commit, a merge, a commit that did not touch this
folder, or one the upstream already holds, which undoing here would only make the branch diverge
from.

## The three copies

A module exists in three places and they stay in this order: the editor is never behind the
workbook, and the workbook is never behind the repository.

1. The editor's text, flushed to the module when typing pauses.
2. The workbook's module, saved to the .xlsm by a save.
3. The file in the folder, committed to the repository.

- **The rows are live.** The pane compares the editor's current text with what HEAD holds, read
  with `git cat-file`, using the same comparison the sync dialog draws with. Computing the rows
  writes nothing.
- **Saving exports.** A save made through this editor writes the folder through the existing
  export path in true-up mode, so the folder always equals the last saved workbook and a developer
  using VS Code or the command line on the folder sees a normal working tree. A save made from
  Excel's own window is not seen; the folder catches up at the next save here or at Commit.
- **Commit is save, then export, then commit.** Commit applies pending annotations, saves the
  workbook if it is dirty, exports, and commits the ticked rows' files with `git commit --only`,
  so files somebody staged outside are left alone.
- **Staging is per module.** No hunks. A module is the unit VBA compiles, and half a module
  staged is a commit that cannot build.

## Rows

The rows compare CODE: the attribute header a file carries and a form's design are compared at
commit time by the export, not on every refresh, because a header costs one temporary VBE export
per class module to read and a design costs a markup projection per form. An attribute-only change
therefore shows as clean in the pane until the next save or commit writes it, and the docs say so.

A rename is paired from the change log: a module the log knows under an earlier name that is
missing from the project now, beside a module the log knows arrived under the new name, is one
renamed row rather than a deletion and an addition. Git pairs the files itself at commit time.

A second section, Folder, lists modules whose file in the folder differs from the live text. That
is the outside-change signal: a checkout, a pull, an edit in another editor. It offers Import
(folder to project, through the existing sync import, which keeps tabs and the caret and applies
annotations afterwards) and Export (project to folder). Nothing imports silently, ever. The host
watches the folder and the repository's HEAD and refs, and taps the pane with a stamp; the pane
re-reads a moment later if it is showing. The watcher is armed by the last step of a status
round, which checks that the project still names the folder, because a forget can land while
the round is out with git, and a watcher armed for a forgotten folder held it, and stamped the
page for it, until the workbook closed.

## Where the repository lives

One folder per project, remembered per project in sync.json as `Repository`, a field of its own.
The sync dialog's `Folder` is not reused: every apply and every folder pick re-point it, so one
export to a scratch folder would have moved the repository. The pane offers the sync folder as its
first suggestion and otherwise a folder beside the workbook named after it.

The folder is flat, because the export is flat. The repository root is whatever
`git rev-parse --show-toplevel` answers from the folder: a folder inside an existing repository
joins that repository at its relative path, so a workbook and its add-in can share one repository
under two subfolders. With no repository above it, Initialize runs `git init` in the folder on a
branch named `main` - GitHub's default, where a repository made here is most likely headed, unless
a configured `init.defaultBranch` names another - sets `core.autocrlf` false so a module
round-trips byte for byte, and writes a `.gitignore` for the export's lock and partial files. An
existing repository's configuration is never touched.

Two open projects never share a folder; the pane refuses to point a second project at a folder
another open project already uses, because a true-up export of one would delete the other's
modules. Two projects can share a repository. Branch and remotes are the repository's, so a
checkout imports into every open project that lives in it and refuses if any of them is dirty,
naming which.

An unsaved workbook has no path and no identity that survives a restart, so the pane says to save
the workbook first.

## Blame

`git blame --porcelain HEAD -- <file>` on the committed file, never on the working copy. The
file's lines are mapped to the editor's lines through the same diff that draws the rows: an equal
pair maps a committed line to a live line, and every live line without a pair is uncommitted. That
one mechanism accounts for the attribute header, for attribute lines inside procedures that the
file carries and the module does not, and for local edits. The page paints the hint as injected
end-of-line text on the model, so it survives tab switches, and refetches when the model's version
moves.

## History and the past-version tab

History is `git log` over the folder's path, newest first, with the files each commit touched.
Open this version puts the module's text at that commit in a read-only editor tab titled
`Module1 @ abc1234`. The tab is host-listed with the face `history:<short>`, the way a form's
designer tab is host-listed with the face `design`, so the strip, the `ui` route and the host agree
about it and it survives every republish. Restore writes the text through the ordinary module
write, so the Changes pane records it as a round labelled with the commit, and it can itself be
restored away. A form's restore is its code; its design is not restored, and the outcome says so.

## Remotes

A remote is attached from the pane. The ready state's first line is the remote: with none, an
input for the URL and Add remote, which runs `git remote add origin`; with one, its URL and
Change, which shows the same input filled in and re-points it with `set-url`. Fetch, Pull and
Push are greyed until a remote is attached, because without one they could only be refused. The
repository on the other end is made first, on GitHub or as a bare folder, since gh is not bundled
(decision 18); the flow is create it empty, paste its URL, push.

Fetch, pull and push run git with terminal prompts disabled, so nothing can hang waiting for a
password in a hidden console; a missing credential helper fails in words. A first push on a
branch with no upstream sets it, to origin when there is one and else to the only remote,
because a bare `git push` stops there to ask for `--set-upstream`, which a button cannot answer. A pull that conflicts
leaves the folder conflicted: the pane says so, blocks Import until the folder is clean, and offers
Abort. Resolving conflicts inside the VBE is a later milestone.

## What it deliberately does not do

- No hunks, no merge or rebase inside the VBE, no history rewriting, no submodules.
- The workbook binary is not committed. The repository holds source.
- A form's design is compared and written at commit, not in the live rows.
- Saves from Excel's own window are not exported until the next save through this editor.

## The wire contract

One brain, two doors: the pane's request and the api's `scm` route reach the same three steps -
gather on the host thread, work off it, apply and answer on it - so an api action leaves the state
the pane's button would. The route runs pool-side and crosses to the host thread only for the
gather and the apply, because git can take seconds and the host lane's budget is three.

### Requests

The pane posts `{type: "scm", id, ...args}` and receives `{type: "scmResult", id, json}`; the host
taps it with `{type: "scmStamp", stamp}`. The route is `GET|POST scm?action=...`. Every value is a
string. `project=` resolves a display name, a full path or an identity, and absent means the shown
project. `by=` attributes writes, as everywhere.

| action | arguments | answers |
| --- | --- | --- |
| absent, `status` | | ScmStatusReply |
| `settings` | `folder=` | the status, after remembering the folder |
| `forget` | | the status, after forgetting it |
| `browse` | | the status, after the host's folder chooser (blocks; the pane's, not a harness's) |
| `init` | | the status, after `git init` on `main` (a configured `init.defaultBranch` wins) |
| `identity` | `name=`, `email=` | the status, after writing the repository's identity |
| `remote` | `url=`, `remote=` (origin when absent) | the status, after `git remote add`, or `set-url` when the name exists |
| `commit` | `message=`; body: module names, one per line; empty means every row | ScmCommitReply |
| `undo` | | ScmUndoReply: the head taken back one commit with `git reset --mixed HEAD~1`, its `subject` and whole `message`, and the status; refused with the status's `undoBlocked` reason |
| `branch` | `name=` | the status, on the new branch cut from the head; git's words when the name is refused or taken |
| `export` | | the status, after writing the folder |
| `import` | body: module names, empty means every Folder row | ScmImportReply |
| `log` | `limit=`, `module=` | ScmLogReply |
| `diff` | `module=`, `ref=` (HEAD when absent) | ScmDiffReply, live against the ref |
| `show` | `module=`, `ref=` | ScmTextReply: the text at the ref and its rows against live |
| `open` | `module=`, `ref=` | the status; the past-version tab is open |
| `restore` | `module=`, `ref=` | ScmRestoreReply |
| `blame` | `module=` | ScmBlameReply |
| `checkout` | `ref=`: a local branch, a commit, or a remote's branch as `origin/name` (bare `name` when one remote has it), checked out as a local branch tracking it | the status, after checkout and import |
| `fetch`, `push` | | the status, with git's words in `detail` |
| `pull` | | the status after `git pull --no-rebase` and the import it brings; refused over a dirty workbook like `checkout`; a conflict answers the `conflicted` status rather than an error |
| `abort` | | the status, after `git merge --abort` |

### Replies

Every reply carries `detail` first. Refusals are `{error}` and the harness client throws on them.

- `ScmStatusReply`: `detail`, `project` (display), `projectId`, `state` (`noGit`, `noProject`,
  `unsaved`, `noFolder`, `noRepository`, `noIdentity`, `conflicted`, `ready`), `folder`,
  `repository` (the root, or empty), `gitVersion`, `branch`, `head` (the branch's commit in
  full; empty on an unborn branch), `upstream`, `remote` and `remoteUrl` (origin, else the only
  remote; empty with none), `ahead`, `behind`, `dirty` (the workbook), `identity` `{name,
  email}` or null, `rows[]` of `{module, kind, file, status, from}` with status
  `modified|added|deleted|renamed`, `outside[]` of `{module, kind, file, status}` with status
  `folderNewer|missingInFolder|missingInProject`, `branches[]` of `{name, current, upstream,
  remote}` - the local branches, then those only a remote has, `remote` naming it and empty
  for a local one - `conflicts[]` of file names, `lastCommit` `{hash, short, author, when,
  subject}` or null, `undoBlocked` (empty when the head can be undone from here, else why not),
  `suggestedMessage`, `covers`.
- `ScmCommitReply`: `detail`, `hash`, `short`, `committed[]` (module names), `skipped[]` of
  `{module, why}`, then the same status fields.
- `ScmUndoReply`: `detail`, `hash` and `short` of the commit taken back, its `subject` and
  whole `message` (subject and body, for the box), then the status under `status`.
- `ScmImportReply`: `detail`, `imported[]`, `skipped[]` of `{module, why}`, then the status.
- `ScmLogReply`: `detail`, `commits[]` of `{hash, short, author, email, when, subject, body,
  files[]}` where files are `{module, file, status}`.
- `ScmDiffReply`: `detail`, `module`, `ref`, `rows[]` of the sync diff row shape.
- `ScmTextReply`: `detail`, `module`, `ref`, `text`, `rows[]`.
- `ScmRestoreReply`: `detail`, `module`, `ref`, `did` (`written|added|unchanged|skipped|failed`),
  `why`.
- `ScmBlameReply`: `detail`, `module`, `head`, `lines[]` of `{line, hash, short, author, when,
  summary}` for committed lines, `uncommitted[]` of live line numbers.

The `status` and `covers` sentences are the pane's own words and the route's, from one constant.

### The page acts

`act("scmPane", {press})` presses `refresh|commit|undo|export|import|fetch|pull|push|init|abort|blame`
and `remote`, the remote line's button; `{file}` points the pane at another open file through the
select's own change; `{tick, on}` ticks a row; `{message}` types the message; `{module}` opens a
row's comparison; `{commit}` opens a commit; `{branch}` picks a branch through the select, a
remote's as `origin/name`; `{create}` makes one through the select's New branch... entry and the
card it raises; `{url}` types the remote URL, unfolding the input when a remote is attached;
`{width}` puts the divider where a drag to that many pixels would leave it. `ui.scm` is the
pane's state: `project`, `state`, `branch`, `branches` (a remote's with the remote in front),
`head`, `undoable`, `undoBlocked`, `remote`, `remoteUrl`, `rows`, `outside`, `commits`,
`showing`, `message`, `busy`, `behind`, `listWidth`.
`act("blame", {which: toggle|on|off})` drives the editor layer and `ui.blame` reads it.

## Threading and the host

- COM is read on the host thread only: the live modules' code through `ProjectReader.ReadSource`,
  the workbook's saved flag, the display names. What crosses to the pool thread is strings.
- The save and the dirty check are the host's. Excel answers them from `Workbooks` and Word
  from `Documents`, with the same members; Access has neither, a database being no document,
  so the project's own `Saved` flag says whether its modules need saving and the editor's own
  Save is what saves them. Before that, Access read as never dirty and a checkout imported over
  an unsaved edit that Excel refuses to touch ([lessons.md](lessons.md) finding 80). Word's
  branch is proven on its own fixture, which is also where a Word project turned out to name
  Word's save-time temp file as its own after every save; the document the checks look up by
  name comes through the one resolver every reader of a project's file uses (finding 81).
- git.exe runs on the pool thread through one runner: hidden window, both streams drained, UTF-8,
  `GIT_TERMINAL_PROMPT=0`, `--no-pager`, `-c core.quotepath=false`, a deadline per command (20
  seconds; 120 for the remote commands), killed with its tree on timeout. Every invocation is
  logged with its arguments, exit code and elapsed time.
- git.exe is found on PATH, then in Git for Windows' three standard locations. Absent, the state is
  `noGit` and the pane's empty state names the download; nothing else in the product changes.
  Found as Git for Windows' `cmd\git.exe` (or `bin\git.exe`), which is a launcher that starts
  the real `mingw64\bin\git.exe` as a second process on every call, the real binary is started
  directly, with the launcher's PATH additions and `MSYSTEM` set the same way so ssh and the
  credential manager are found exactly as before. The launcher was 7 of every 18 milliseconds a
  git call cost (2026-09-09).
- The repository root is remembered per folder once `rev-parse` has answered it, trusted while
  its `.git` is still there and none has appeared in the folder itself; the identity is one
  `config --get-regexp` read only by the actions that need it, the remotes are read only when
  asked for, and a commit's parents ride the log format (`%P`), so the head's undo question costs
  no process of its own.
- Writes into the project go through `WriteModule` with the change log's restore bracket, so every
  checkout, import and restore is a round. Exports go through `ModuleSyncService.Apply` under the
  folder lock with the built-in planner, so the files are what the sync dialog writes.
- The inside door refuses the route: it arrives on the host thread, and git would run there.

## Surfaces

- The pane, `ui/editor/src/scmpane.ts`; the blame layer, `ui/editor/src/blame.ts`; the
  past-version tab, a `history:<short>` face beside the designer's `design`.
- The host, `src/Xlide.Vbe.Shim/AddIn/AddInSession.Scm.cs`, `src/Xlide.Vbe.Shim/Scm/`; the pure
  parts, `src/Xlide.Vbe.Core/Scm/`, unit-tested against captured git output.
- The api, `scm`, documented in [xlide-api.md](xlide-api.md) and
  [driving-excel.md](driving-excel.md).
- The suite, `tools/harness/scm.mjs` against `ScmFixture.xlsm` (`tools/New-ScmFixture.ps1`),
  which initialises a repository in a temporary folder, commits, edits, branches, imports, restores
  and blames, and puts the fixture back.

## What shipped

**Shipped 2026-09-08**, as designed above, with these things learned on the way:

- **A remote is attached from the pane** (later the same day): the flow from Initialize to a
  first push no longer leaves the VBE for `git remote add`. The status carries the remote and its
  URL, the remote verbs are greyed until one is attached, and the same verb re-points a wrong
  paste with `set-url`.
- **Initialize names the first branch `main`** (2026-09-09). git's own default is still `master`,
  and the first live test pushed a `master` at a GitHub whose default is `main`. A configured
  `init.defaultBranch` is the developer's and wins; a git without `--initial-branch` (before
  2.28) falls back to its own name.
- **The divider drags** (the same day). The two halves were a fixed grid, a third and two
  thirds, and a long module name or a wide comparison had no way to ask for more. It is the
  Changes pane's rail divider now, with the floor, the ceiling and the arrow keys, and the width
  is kept in the page's storage beside the dock sizes.
- **Undo, New branch..., and the remote's branches** (the same day, from the same live test).
  Undo is `git reset --mixed HEAD~1` rather than the soft reset VS Code runs: this pane never
  shows the index, so a soft reset would leave the undone commit staged where only an outside
  `git status` could see it, and the next `--only` commit of one row would leave the rest of it
  there. The reasons Undo is greyed ride the status as `undoBlocked`, read with every status, so
  the button explains itself instead of refusing when pressed. A remote's branch is offered
  with the remote in front and checked out with an explicit `--track`, because a bare
  `git checkout origin/name` detaches and the bare name's do-what-I-mean fails as soon as two
  remotes hold it.
- **A refresh is 121 ms of git, from 305** (the same day, [lessons.md](lessons.md) finding 76).
  A status was ten processes and the history read after it five, four of them the same fixed
  overhead; Git for Windows' launcher was a third of every one. The binary is started directly,
  the root is remembered, the identity is one read for the actions that need it, the remotes are
  read when asked for, the parents ride the log, and the gather reads one module for the actions
  about one module. The suite holds the counts.
- **The branch select acts on a pick, not on every keystroke.** A closed select fires change
  on each arrow key and typed letter, and each was a checkout with an import; a walk now settles
  on Enter or on the select losing focus, and Escape puts it back. A ref, name, email, URL or
  remote beginning with a dash is refused before git could read it as an option, and an undone
  commit's open comparison and unfolded state go with it.
- **Pull imports what it brings** (the same day, [lessons.md](lessons.md) finding 79). It ran
  git and reported, and the button had promised the import from the start; a pull left
  unimported was undone by the next save's true-up export and the reversal offered as a commit.
  It refuses over a dirty workbook the way a checkout does, merges divergent branches
  (`--no-rebase`, since git otherwise refuses a divergent pull until told how), and answers a
  conflict as the conflicted state. The suite pulls a colleague's commit from a clone of its
  bare remote, and a conflicting one.
- **The rows compare without copying** (the same day, [lessons.md](lessons.md) finding 77).
  Each status compared every module with the branch head and with its file by building four
  normalised copies of each text; `ModuleSync.SameCode` walks both texts in place instead. An
  upstream whose remote branch is gone reports `upstreamGone`, so zero ahead is not read as
  "already pushed" and the head's Undo stays live. And the status, diff and blame log their
  phases, which is how the measuring was done.

- **An import writes one line terminator fewer than the file carries.** The export appends a
  newline to a module's text, as a text file should, and the applier wrote the file's whole body
  back on import, so every module grew an empty last line on each round trip through the folder
  and a checkout's blame reported that line as uncommitted. `ModuleSync.BodyForImport` is the
  faithful inverse of the export now, and the sync dialog's import took the same fix.
- **Blame turns itself off when the host refuses it**, and says why in the pane's notice: with
  no repository, no git, or a module nothing has committed, a button left pressed over nothing
  painted would claim a state the host had just declined. `ui.blame.refused` carries the words.
- **A comparison stays on screen across a refresh.** Every read ends in a redraw; the pane keeps
  the rows it last drew and asks for them again, rather than replacing them with "Comparing...".
- **A stamp during an action is not dropped**, and a hidden pane runs no git: the re-read a stamp
  asks for waits for the action to finish, and a pane behind another tab re-reads when it comes
  forward.
- **The past-version tab is a host-listed face** `history:<short>`, so the strip, the `ui` route
  and the host agree about it, and `act("closeActive")` on one now reports the close.
- **git's output reaches the parsers whole.** Blame's date spells a zero offset `Z` as `%aI`
  does; a control character stored inside a commit message cannot begin a record or end a body;
  a 2,100-line untouched module no longer reads as uncommitted from its first line.
- **The harness learned two things about this machine.** The fixture driver's blank-workbook Excel
  is released wrapper by wrapper before `Quit`, because a wrapper the collector finalised after
  the server had gone made DCOM start a fresh hidden Excel that the launcher's sweep then refused
  to close. And the headless page probe reads the DevTools endpoint from the profile's
  `DevToolsActivePort` file and closes the browser over the protocol, because the msedge.exe a
  probe starts can hand off to a second process and exit at once, taking its stderr with it and
  leaving the browser running under a dead parent.

Not shipped, on purpose: Publish to GitHub, pull requests, conflict resolution inside the VBE, and
the form design in the live rows. Each is a later milestone the design already names.
