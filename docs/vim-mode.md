# Vim mode

Vim's modes and keys in xlide's code editors. Off until you turn it on.

Requested by u/CptBadAss2016 on Reddit
([#173](https://github.com/WilliamSmithEdward/xlide_vbide/issues/173)).

## Turning it on

Open Settings (Tools > Options, or the wrench menu) and tick **Vim mode**. It takes effect at
once in every open code tab, and the status bar starts with `-- NORMAL --`. Untick it and the
caret, and every key, go back to how they were.

The setting is `editor.vimMode` in `%LOCALAPPDATA%\xlide_vbide\settings.json`.

## What works

The keymap is [monaco-vim](https://github.com/brijeshb42/monaco-vim) 0.4.4, which is CodeMirror's
Vim emulation with an adapter for the Monaco editor. Its coverage is what you get. The table
lists what the keymap binds; the live suite presses a sample of each row, named under
[How it is tested](#how-it-is-tested), and the rest are bound and untried here.

| | |
| --- | --- |
| Modes | normal, insert, visual (`v`), visual line (`V`), visual block (Ctrl+Q), replace (`R`) |
| Motions | `h j k l`, `w b e`, `0 ^ $`, `gg G`, `f t F T ; ,`, `{ }`, `%`, `H M L`, counts |
| Operators | `d c y`, `> <`, `g~ gu gU`, with motions, counts and doubled forms (`dd`, `yy`, `>>`) |
| Text objects | `iw aw`, `i( a(`, `i" a"` and the other bracket and quote pairs |
| Editing | `x`, `r`, `J`, `o O`, `i a I A`, `p P`, `~`, `u`, `.` |
| Registers | named registers (`"ayy`, `"ap`), the numbered and small-delete registers |
| Marks | `ma`, `'a`, `` `a `` |
| Macros | `qa` ... `q`, `@a`, `@@` |
| Search | `/` and `?` with `n N`, `*` and `#` |
| Commands | `:w`, `:q`, `:wq`, `:x`, `:{line}`, `:s/old/new/`, `:%s/old/new/g`, `:sort`, `:noh`, `:undo`, `:redo`, `:map` |

`/` takes a JavaScript regular expression, as the prompt says.

Vim mode is on the code editors only. The Immediate line, the text half of the form designer and
a past version opened from Source Control keep the plain keys.

## The status bar

The left end of the status bar is Vim's line:

- the mode, as `-- NORMAL --`, `-- INSERT --`, `-- VISUAL LINE --` and so on;
- the `:` or `/` prompt while one is open, and `(recording)[a]` while a macro records;
- what Vim last said, such as `Not an editor command ":foo"`, for five seconds;
- the keys of a command you have started and not finished, such as `2d`.

With the editor split, it shows the mode of the group you last typed in. Each group keeps its own
mode. Registers are shared, so a yank in one group can be put in the other.

## Where it differs from Vim

**Ctrl shortcuts stay the editor's.** Ctrl+C copies, Ctrl+V pastes, Ctrl+X cuts, Ctrl+A selects
all, Ctrl+F opens Find, Ctrl+Y cuts the line, Ctrl+Z undoes. Vim binds most of those to other
things, and here it does not get them. There are four exceptions:

| Key | Does | When |
| --- | --- | --- |
| Ctrl+[ | Escape | every mode |
| Ctrl+Q | starts a block selection | outside insert mode |
| Ctrl+D | scrolls down half a page | outside insert mode |
| Ctrl+U | scrolls up half a page | outside insert mode |

Ctrl+Q is Vim's own stand-in for Ctrl+V on systems where Ctrl+V pastes.

**The VBE's keys stay the VBE's.** F5 runs, F8 steps, F9 toggles a breakpoint, F2 opens the
Object Browser, Ctrl+S saves, Ctrl+G opens Immediate, Ctrl+W closes the tab. Excel's host takes
these before the editor sees them, in either mode.

**Redo is not Ctrl+R.** Ctrl+R opens the Project Explorer, as it does in the native editor. Use
`:redo` or Ctrl+Shift+Z.

**`u` is the editor's Undo.** One press undoes one step of the same history Ctrl+Z walks. Vim
undoes a whole insert at once; here a typed line and the recasing xlide applies to it afterwards
are two steps, so it can take two presses. In exchange `u` undoes a rename across every module
it touched, which Vim's own undo could not.

**Yank and put do not use the Windows clipboard.** `y` and `p` use Vim's registers, which live
only in the editor. Ctrl+C and Ctrl+V use the clipboard. There is no `"+` register.

**Saving saves the project.** `:w` runs the host's Save, which writes the whole workbook, document
or database. `:q` closes the tab and asks about unsaved changes when there are some. `:wq` and `:x`
save and then close without asking.

**Completion still works.** The completion list opens in insert mode as it does without Vim, and
Escape closes it as well as leaving the mode. A selection made another way, such as Ctrl+A, is a
selection the editor's own commands act on; the status bar goes on saying `-- NORMAL --`.

## What is not there

- No vimrc. `:map`, `:nmap`, `:imap` and `:vmap` work for the session and are gone when Excel
  closes.
- No `:e`, buffers, windows or tabs commands. Tabs and splits are the editor's own.
- No Ctrl+O and Ctrl+I jump list, no Ctrl+A and Ctrl+X on numbers, no Ctrl+W in insert mode. The
  editor or the host has those keys.
- No system clipboard register.

## Keyboards with AltGr

On layouts where `[`, `{`, `@` or `\` need AltGr, the browser reports the key as Ctrl+Alt. Vim
mode gives Vim the character that was typed, so `{` is still "back a paragraph". This path is
covered by a unit test and a synthetic key event, and has not been tried on a physical AltGr
keyboard.

## How it is built

| File | Holds |
| --- | --- |
| `ui/editor/src/vimmode.ts` | which editors get an adapter, the status bar line, what `:w`, `:q` and `u` mean, and the guard described below |
| `ui/editor/src/vimkeys.ts` | the key policy, as a pure function |
| `ui/editor/build.mjs` | sends monaco-vim to its ESM build and its two Monaco imports to the page's own Monaco, and fails the build if a second copy of Monaco is read |
| `src/Xlide.Vbe.Core/Editor/ProductSettings.cs` | the `editor.vimMode` setting |

One thing needed a guard. The keymap records what goes into the document during insert mode so
that `.` can type it again. It cannot tell your keys from the edits xlide makes around them: the
`End If` that Smart Enter writes, and the recasing of `dim` to `Dim` a moment after the line
rests. With those recorded, repeating `dim delta as long` produced `Dim delta As longL`. Changes
xlide makes are kept from the keymap's record, so `.` types what you typed and xlide treats the
repeat as it treated the original.

## How it is tested

- `ui/editor/test/vimkeys.mjs` pins the key policy: which Ctrl chords Vim gets in which mode, and
  the AltGr case.
- `tools/harness/vim.mjs` runs against a real Excel and presses real keys through the page's
  DevTools protocol. It checks that normal-mode keys move the caret and type nothing in the
  editor or the project, that an edit and its undo reach the project's module, the repeat above,
  the Ctrl keys, the `/` and `:` prompts, a macro, `:w`, `:q` and `:wq`, the status bar's
  contrast in both themes, and that turning the setting off restores plain typing. It is part of
  the release gate. The keys it presses: `j w $ gg G 3G`, `dd yy p u`, `ciw`, `o`, `.`, `x`,
  `v e`, `V`, `>> <<`, `qa` `q` `@a`, Ctrl+Q, Ctrl+[, Ctrl+D, Ctrl+U, Ctrl+A, `/` with `n`,
  `:4`, `:w`, `:q`, `:wq`. Replace mode, a mark and a put across a split were tried by hand
  once and are not in the suite.
- The xlide api reports the mode as `ui.vim` and takes the setting as `settings?vimMode=true`.
  [xlide-api.md](xlide-api.md) has the fields.

Not covered by a test: the function keys with Vim mode on. They never reach the page, so the
mode should not affect them, and nobody has pressed them to check.
