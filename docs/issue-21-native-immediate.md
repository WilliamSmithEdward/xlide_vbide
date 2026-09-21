# Issue 21: native Immediate evaluation

The issue was reopened on 2026-09-21. The user authorized the narrow input exception recorded in
decisions.md, section 16. Paused evaluation now uses the native Immediate window; design-mode
evaluation retains its existing scratch procedure.

## Finding

The hidden native Immediate window accepts WM_CHAR text followed by WM_KEYDOWN/WM_KEYUP for
Enter. WM_CHAR for Enter alone did not execute the line. The existing ImmediateReader can read
the resulting transcript while the native window stays hidden. Evaluation neither changes the
clipboard nor sends global keyboard input.

The isolated Excel DebugFixture regression suite exercises:

- Arithmetic reads the paused procedure's locals.
- VBA function calls and string concatenation use the paused scope.
- Repeated expressions return results each time.
- Object members such as ThisWorkbook.Name work.
- Division by zero and malformed syntax report native error dialogs; acknowledging those
  dialogs preserves the debug session and allows another expression to succeed.
- Stepping changes the next expression's result from 2 to 3, proving that the live local is read.
- Runner's source remains unchanged and no scratch evaluation module is created.
- Debug.Print from the Immediate input and from resumed VBA reaches the panel once.
- UI submissions recover from errors without relying on the debug API's dialog rescue.
- Empty results, multiline strings and printed spaces survive output capture.
- Long native history remains readable beyond the provider's 1024-character GetText limit.
- Debug.Print still arrives when the native window drops its oldest history lines.

Run against an isolated debug fixture, passing the PID printed by Start-Excel.ps1:

```powershell
tools/harness/Start-Excel.ps1 -Workbook artifacts/fixtures/DebugFixture.xlsm -Separate
node tools/harness/native-immediate.mjs <pid>
```

The suite also accepts XLIDE_PID and is part of verify.ps1's live DebugFixture group.
Live validation on Excel passed 97 native Immediate checks and the existing 24 Immediate/Watch
checks. Repeating the suites in one session also passed. The repository gate includes 664 core
unit tests, along with the page, engine and native publish checks.

## Implementation constraints

The session identifies the native window by its localized object-model caption. Evaluation
checks its process and debugger mode, serializes submissions, and never falls back to project
edits while paused. A unique command comment separates the echoed input from the output; the
poller is suppressed during evaluation and its baseline advances before the result is displayed.
Pending Debug.Print output is delivered first.

The accessibility provider caps individual text reads at 1024 characters even when a larger
length is requested. Reading 512-character ranges avoids that cap. The original code pane must
be shown after evaluation, because selecting the native Immediate caret changes the target of
VBE commands. Calling into the project object model between typing and Enter also disturbed
native execution during testing; the mode recheck happens before positioning and typing.

New error notices are captured and acknowledged off the host thread. This scope is shared by UI
and API evaluations and never issues Reset. The API's scratch recovery is disabled for a
request that began in a paused session. Its completion can safely arrive after an HTTP timeout.

The native working buffer is replaced for each paused command, after pending output is delivered
to XLIDE. The visible Immediate history is retained. This avoids stale native accessibility
positions once VBE's history has rolled over; moving to the end of an accumulated native history
proved unreliable under repeated evaluations.

Commands must be one line, without control characters, and fit the native line limit including
the internal comment. If the command cannot be verified before Enter, it is not executed. If
its output cannot be recovered in full, the evaluator reports failure without rerunning it.

Microsoft documents that Immediate statements in break mode run in the current procedure's
scope: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/immediate-window
