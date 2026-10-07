# Settings and layout

XLIDE saves preferences for the current Windows user in `%LOCALAPPDATA%\xlide_vbide`.
Excel, Word, Access, and PowerPoint load the same preferences. Changes are saved as they
are made, so closing an Office application is not required to retain a layout.

- `settings.json` holds editor behavior, Explorer choices, analyzer rule overrides, and
  the agent API switch. Existing settings files keep their format and location.
- `sync.json` holds each project's import/export choices and source-control folder.
- `ui-state.json` holds pane visibility, active pane tabs, dock sides, tab order, split
  proportions, pane sizes, the Source Control divider, editor group geometry, and the
  editor and Object Browser windows' normal bounds and maximized state.

Editor group geometry is shared across hosts. Restored empty groups are available for
the modules opened in the new session; source text and document identities belong to
the Office project. Minimized windows do not replace the saved normal bounds. Restored
windows fit the nearest available monitor's work area.

All three files use one persistence service. It merges changed fields under a process
lock and atomically replaces the document, retaining unrelated choices written by
another Office session and preserving unknown settings from newer versions. A save
failure is logged; a failed pane-layout save also displays a notice in the editor.

Browser profiles remain isolated by process. Their local storage is a reload cache and
a migration source for older builds; the per-user files carry preferences across
processes. A fresh session prefers the shared file over an old profile, including when
Windows reuses a process ID; a reload within the session retains its current cache.
A browser-profile cleanup does not remove these files.

The focused live check is `tools\harness\Test-UserPreferences.ps1`. It customizes the
visible layout and verifies it in fresh Excel, Word, Access, and PowerPoint processes,
then checks maximized windows across Excel and Word. It restores the user's files and
closes only its own hosts. The release gate runs it against an isolated Debug preference
root, along with the other live suites, so test layouts cannot change the user's layout.
