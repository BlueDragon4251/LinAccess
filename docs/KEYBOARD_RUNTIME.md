# Native Access keyboard and event subset

This is executable C++20/Qt behavior, not complete Access 2003 keyboard or UI
parity. The pre-change regression reproduced Escape closing the native form via
QDialog instead of cancelling edits. Tests use actual Qt key events and independently
inspect SQLite state. No Windows Access 2003 event or rendering oracle is available.

## Implemented native document keys

| Input | Native effect |
|---|---|
| Escape in a form | Restore the pending focused field to its typed focus-entry value; a further Escape discards remaining record edits |
| Escape in the report layout preview | Close that preview through its normal lifecycle |
| Shift+Enter, including keypad Enter | Save the current document record, honoring validation/cancellation |
| F9 | Recalculate displayed calculated controls without reloading records |
| Shift+F9 | Requery the current native form source with normal save/validation handling |
| F2 in a line editor | Toggle whole-field selection and editing at the end of the text |
| Enter in a line editor | Commit pending controls and advance focus; do not activate an unrelated default dialog button |
| Ctrl+W / Ctrl+F4 | Close the containing document window, honoring Unload cancellation |

Undo retains NULL versus empty text, restores duplicate bindings to the same field,
and can discard invalid numeric text before conversion. Uncommitted new records
are not inserted by Escape. Form.Undo discards the pending record; Form.Recalc
recomputes displayed expressions without a source reload or save. These are real
native form methods. Report invocation is explicitly unsupported.

This does not implement every Access Undo event, record conflict, edit/navigation
mode, TabIndex, option, keyboard layout, menu accelerator or designer shortcut.
Other controls retain their existing native Qt behavior; complete Access focus,
selection and UI parity still requires reference tests.

## VBA key events

SaveAsText OnKeyDown, OnKeyPress, OnKeyUp and form KeyPreview are retained.
Native keyboard delivery invokes only bound event procedures with actual ByRef
KeyCode/KeyAscii arguments and Shift bit masks (Shift=1, Control=2, Alt=4).
KeyPreview chooses form-before-control versus control-before-form delivery for each
stage. Embedded documents own their events; the parent's KeyPreview does not
intercept the child's input. Cancelling KeyCode suppresses subsequent/default key
processing. KeyAscii can cancel or replace a supported character before Qt inserts
it. Handler failures stop default processing and appear in native diagnostics.

The bounded key map covers the documented common navigation keys, ASCII letters/
digits, function keys and common punctuation; it is not full Windows virtual-key,
keyboard-layout, IME or locale emulation. KeyPress conversion/replacement currently
supports ASCII; native Unicode text without a KeyPress handler remains supported.
Non-ASCII bound KeyPress conversion, virtual-key replacement, macro/function event
bindings and complete Ctrl/Alt character-event behavior remain explicitly limited.
VBA vbKey constants and acShiftMask/acCtrlMask/acAltMask are available for the
implemented native event paths. No desktop-global input injection is introduced.

`tests/keyboard_commands_tests.cpp` checks saved SQL values, cancelled saves/closes,
NULL restoration, aliases, invalid numeric input, uncommitted new records, F2/focus,
record-count distinction between Recalc/Requery, real method execution and key-event
SQL traces. KeyPreview order changes the actually inserted character. KeyCode and
KeyAscii cancellation, typed modifier masks and a handler raising error 11 are
checked independently of default Qt behavior.

## Original source boundaries, not invented bindings

All original artifacts remain hash-verified and unchanged. The three existing
KeyPress procedures in modules
`6bde4373d4584a1626a6a7f5edbb14d072f3f01083bd797faf9fee914f04f7f8` and
`6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8`
are separately tested as explicit original calls. Only the latter procedure has a
matching bound control in the exported layout. The former layout binds two renamed
controls whose handler procedures are absent. Native dispatch records those as
ProcedureMissing and does not substitute the older similarly named procedures.
These checks do not establish a fully working original sale workflow.

The complete original login-layout probe also reproduces a source-column boundary:
its combo declares ColumnCount=5, while its unchanged handler requests Column(5).
The native handler fails with error 5 at source line 34, before authentication
comparison. The regression uses deliberately synthetic non-credential input and
never reads or logs original passwords. It does not claim successful login.
Microsoft's current [Column documentation](https://learn.microsoft.com/en-us/office/vba/api/access.combobox.column)
limits the zero-based index to ColumnCount-1; an actual Access 2003 comparison is
still needed to classify the supplied original application's behavior. Neither
layout/source bytes nor column limits are silently altered to make this pass.

Validation: Release build, 18/18 CTests (49.12 s), GUI-disabled Release core 1/1
(0.37 s), fresh .deb executable Build ID/payload/permissions, and four unchanged
original fixture SHA-256 checks passed. Remote distribution CI and Access 2003
reference behavior remain unverified.


## Native form date/time insertion keys (2026-11-03)

Ctrl+semicolon inserts the current local ISO date; Ctrl+Shift+semicolon/colon
inserts the current local HH:mm:ss into the focused native form TextBox,
replacing its selection or inserting at the caret. Existing control buffers,
Change events, BeforeUpdate cancellation and record saving are retained.
Read-only/disabled/calculated/unsupported bound contexts and capacity overflow
refuse mutation. The existing layout target passes 1/1 (13.58 s), exercising
actual keys, Change, cancelled save, SQL Date/time persistence and read-only
refusal. App builds. Regional Access date masks and datasheet keys remain open
in this increment; full-suite/package baseline remains 8dc6afa.


## Datasheet clock keys and keyed cell restoration (2026-11-03)

Forms and native editable datasheet text cells share one Qt clock-key mapping.
The table view starts editing the selected cell; an already open delegate editor
preserves the caret/selection. Ctrl+S retains existing validation/manual-submit
behavior and persists the actual inserted value. Read-only/non-text editors do
not mutate. ISO date/whole-second time formatting remains bounded.
Saving a table now restores the selected existing record/column by its complete
primary key after model reset, including when sort order moves that record.
Search is bounded to one million fetched rows; absent/NULL/new-generated keys
and complete large-table/no-key selection semantics remain open.
Main-window regression passes 1/1 (1.65 s), proving real date/time keys, persisted
SQL values, partial replacement and keyed restoration as sorting moves the same
record twice. The shared-helper form layout target also passed in the preceding
integration run. App builds; existing targets only. Last full suite/package/hash
baseline remains 8dc6afa; no complete datasheet/regional-format parity claim.


## Datasheet previous-record value key (2026-11-03)

Ctrl+apostrophe in the native table view copies the same column from the
immediately preceding displayed model row into the current manual-submit draft.
Checked Currency/date/Boolean conversion and independent SQL NULL preserve the
supported field types. Existing drafts are retained; SQL is not submitted.
First-row/no-current, hidden source/target, generated/read-only and binary
contexts do not mutate. Active text-editor insertion/caret behavior, filtered
multiuser record identity and complete Access keyboard parity remain open.
The main-window test verifies actual key dispatch and unchanged SQL before Save;
the Currency target verifies high exact Currency, Date, True stored as -1, NULL,
AutoNumber/BLOB refusal and actual persistence. A focused two-target run passes
2/2 (1.95 s); final typed case passes 1/1 (0.40 s). No new targets.
Full suite/package baseline remains source 24854bf (48/48, 86.21 s), privately
published as 33981e26e528c0d708ba2a2f3be265aedf3a01ea with identical source tree
caca8e5ef7c1de87a10390a3b14d7272a700c895. This key increment is focused-verified.


### Previous-record key inside a native text-cell editor

The native QLineEdit delegate also accepts Ctrl+apostrophe: it replaces the
current field buffer with the checked prior-row value and refreshes the editor.
This command intentionally replaces an unfinished text entry. Read-only/binary/
generated and hidden prior-row contexts retain their guards. The value remains
in the model's manual-submit cache until explicit Save.

Untouched editors now close without writing an empty string over SQL NULL,
including normal Qt delegate submission and the explicit finishEditing path.
Explicit user typing followed by deletion still requests empty text for a text
field. A shared editor-change check distinguishes these cases. The existing
Currency target proves editor key dispatch, NULL retention and explicit empty
text through actual SQL persistence; main-window/Currency checks pass 2/2
(2.22 s). Full suite/package baseline remains 24854bf, as above.


## Native datasheet F2 and Shift+F2 (2026-11-03)

F2 opens an editable table cell with a caret at the end. Pressing F2 in the
native line editor validates its field into the manual-submit cache and returns
to table navigation; invalid input retains the editor and its text. No SQL is
submitted. The focused F2 increment passed the existing main-window/Currency
targets 2/2 (1.88 s), including actual arrow navigation and invalid Currency.
Microsoft documents the [F2 editing/navigation distinction](https://support.microsoft.com/en-au/accessibility/access/keyboard-shortcuts-for-access);
that documentation does not establish complete Access 2003 parity.

Shift+F2 opens a real native text Zoom dialog for a supported table cell or its
active text editor. OK returns a field draft, Cancel preserves the prior value.
Read-only view contexts display text without write controls; unsupported binary
values are refused. An untouched dialog retains original line endings and NULL
semantics. Editing in the multiline dialog uses Qt paragraph/line normalization.
The dialog guards destroyed owners and invalidated model indexes across its
nested event loop. Native form TextBox support was added in the following increment; all Access
Zoom formatting options remain open. The final main-window Zoom integration passes 1/1 (1.56 s), exercising Cancel,
multiline values, unchanged CRLF and view/editor mode transitions. Offscreen
synthetic modal closure explicitly reactivates the main window in the test.
The rebuilt complete desktop suite passes 48/48 (86.44 s), Debian contents/
permissions/Build-ID/pinned SQLite checks pass and four original hashes match.
GUI-free core remains checked at 1/1 (2.63 s); subsequent edits are Qt-only.


## Native form TextBox Zoom and DoCmd.RunCommand (2026-11-03)

Shift+F2 also opens the shared text Zoom dialog for a native form TextBox.
The result uses the existing edit buffer, checked writable-control path, Change
binding and deferred BeforeUpdate/save. Opening the child dialog suppresses only
that field's automatic editingFinished commit while it owns the Zoom dialog;
it does not set the form-wide loading flag across a nested event loop. Cancel
and readonly display preserve the prior buffer; disabled/locked or unsupported
contexts retain their guards. Changed control state is checked again on return.
Native execution lifetime protects a form closed during the nested loop.

DoCmd.RunCommand acCmdZoomBox uses the same native operation through the active
document resolver. The [AcCommand enumeration](https://learn.microsoft.com/en-us/office/vba/api/access.accommand)
defines acCmdZoomBox as 179. An unsuitable active control returns error 2046;
this is not simulated SendKeys or a silent successful command. ComboBox/memo-specific Windows formatting and complete Zoom parity remain open.
Application.RunCommand now forwards the same implemented command subset, with
required/named Command argument binding. Unknown options remain error 3251;
bare unqualified RunCommand and the complete AcCommand set remain open.

The existing layout target passes 1/1 (14.08 s), including readonly Zoom, actual
Change dispatch, cancelled BeforeUpdate, later SQL persistence, interpreted
DoCmd execution and unavailable-control error 2046. Full suite/package baseline
remains local 2fcaf98 / private API ab5e6cd (48/48, 86.44 s); this form/runtime
increment is focused-verified. The GUI-free rebuild/core passes 1/1 (1.53 s).
After Application.RunCommand forwarding was added, the final existing layout
target passes 1/1 (14.90 s), including actual interpreted named-argument
Application execution and unavailable-control error propagation.


## Native field navigation and retained invalid input (2026-11-03)

Ctrl+End completes bounded query fetching before moving to the final field of
the last record. It refuses the supported cell budget instead of silently
selecting the end of Qt's initial fetch batch. Ctrl+Home retains normal native
first-record/first-field navigation. Active edits are validated first.
Tab, Backtab/Shift+Tab, Return and keypad Enter now validate into the field cache
before Qt closes a native delegate editor. Rejected Currency/type input stays
visible and focused. Valid input follows Qt field navigation while SQL remains
unsubmitted. Shared native editor commit logic also serves explicit Finish/F2.
Mouse/focus-out cancellation, all Access navigation policies and exact keyboard
parity remain open; this increment does not claim to implement them.
Existing main-window/Currency targets pass 2/2 (2.01 s): actual 300-row fetching,
Ctrl+End/Home, rejected keys including both Shift+Tab encodings, valid Tab cache
updates and unchanged SQL until explicit Save. No new targets. Full suite/package
baseline remains local e15ff33 / API c381897, 48/48 (88.41 s); this increment is
focused-verified. Theoretical checklist: 279 checked / 70 open scoped points.


## Native mouse validation and datasheet context menu (2026-11-03)

A normal native editor-close request now validates the field before disposal.
Rejected input stays alive, and mouse selection/focus returns to its original
cell after the transition. Escape's explicit RevertModelCache remains a discard
operation; model destruction/reset can still invalidate an editor. The mouse
loss was reproduced before the fix. Complete inter-window/focus policy parity
remains open.
The table view also has a real native context menu for Copy, checked range Paste,
text Zoom, record/column/all selection. Right-click within an existing selection
preserves it. Paste is enabled only for its supported clean editable target and
text clipboard; keyboard-triggered context menus anchor at the current cell.
Actions call the existing checked native operations; cut/range delete, filter,
all Access context commands and pixel/menu parity are not claimed.
Existing main-window/Currency targets pass 2/2 (2.25 s), including actual rejected
mouse navigation, retained editor/focus/current cell, Escape regressions and
real popup action dispatch into quoted TSV clipboard with dirty Paste disabled.
Last full suite/package remains e15ff33 / API c381897, 48/48 (88.41 s). These
Qt view increments are focused-verified, with the same existing 48 targets.
Theoretical checklist: 281 checked / 70 open scoped points.
