# Bounded active-document commands

After native OpenForm returns, global DoCmd commands address the active document
in the invoking project. Explicit Me/property operations still address the
CodeBehind instance. Before the caller's initialization has completed, commands
use that caller; a document from another connection/project is never selected.
This is a runtime change, not static procedure recognition.

Implemented command paths:

- GoToControl resolves the literal control on the active native document and
  selects editable text. Both DoCmd and Application.DoCmd use that object path.
- RunCommand acCmdSaveRecord saves the active native record. Next, Last and New
  navigate the active document's actual record selector; Maximize changes its
  native window. Blank Requery reloads the focused native combo or list box's
  row source when that control belongs to the active form. Without that focus,
  the command reloads the real form record source.
- The corresponding Application.DoCmd object paths share the native host
  operations, with supported omitted GoToRecord positions and numeric Access
  constants. Unsupported commands/options are rejected, not treated as no-ops.
- Application.DoCmd.RunSQL invokes the same real project SQL executor as the
  existing statement path. CREATE/INSERT/UPDATE/DELETE/DROP are exercised on
  disposable test data without additional LinAccess action-consent dialogs.
- Scalar object defaults are resolved before argument validation. Null/default
  Null, Error Variants, argument counts and 32-bit command overflow are tested.
- Native SendKeys resolves the logical active document's focused widget at each
  actual delivery, including a document switch after Wait=False was queued.
  Independent native windows in the same project are supported; no input goes
  to external applications. Closed invoking forms cancel pending delivery.
  The low-level adapter retains its original scope check without this explicit
  project focus resolver.

## Evidence

`tests/active_document_commands_tests.cpp` uses actual Qt windows, actual VBA and
SQLite records. Before the fix, its independent SQL assertion failed: OpenForm
completed but the intended target row was not saved. It now verifies:

- synchronous input/save after opening another form while a same-named caller
  control remains unchanged;
- Next/Last navigation, a persisted new row and Requery after an external update;
- explicit Me writes still address the caller, and only the active target window
  maximizes;
- queued input follows a subsequent form switch, then persists independently;
- a separately constructed same-project window receives input and saves its own
  selected row;
- a dirty window in another project is not saved by the invoking project's
  command;
- Application.DoCmd SQL actions create/change/delete/drop real test data.

The original-form regression in `report_corpus_tests.cpp` also opens an unchanged
original layout over actual MDB records. It supplies a real original control
name to global GoToControl, checks the focused original control, clears the
filter, navigates to Last and requeries. Bound cells/Nulls are compared with an
independent SQL snapshot. That binding probe deliberately overlays empty
CodeBehind inside its rolled-back test transaction, as before. It is not a
claim that the complete original form application ran. Separate unchanged
wizard, calendar and keyboard-handler regressions remain mandatory.

## Limits

All six AcRecord actions now support active-form navigation, positive relative
and absolute saved-record offsets. Named acDataForm targets resolve live,
case-insensitive same-project forms, including hidden windows, without activation.
Missing/closed/foreign targets raise 2450; ambiguous instances raise 5. Tests
cover both command dispatch paths, new-row boundaries and save cancellation.
Named-control DoCmd.Requery now targets the active form's supported control,
including combo/list row sources, subform sources and calculated controls.
Bound non-list data controls requery the form, preserving save validation.
Me.control.Requery retains the instance context. Table/query/report navigation
and RunSQL's UseTransaction option remain explicitly unsupported. Active subform/record-command edge cases,
end-of-record behavior, unavailable-command error parity, other focused-control
Requery routes and complete OS/window activation scheduling remain open.

The focused combo/list route uses the same native row-source refresh as named
Requery. The main-window integration runs both VBA spellings with live Qt
focus on a second form. It checks the refreshed rows, an unsaved form change,
explicit named routing, and the record-source fallback from a text box. The
scope is directly hosted combos and list boxes in native forms. Windows Access
behavior for an omitted name and focused controls remains unverified.

Access SetWarnings, confirmation options and full action-dialog semantics remain
open; this increment adds no LinAccess-specific SQL approval policy. Original
sale initialization, MDI/window/data modes and the complete sale/report workflow
also remain open. Independent native form/report ownership is covered separately
in DOCUMENT_LIFETIME_RUNTIME.md. No Windows Access
result/event oracle is available. All mutations use isolated copies or in-memory
records; original MDB/Bridge bytes remain unchanged.

Validation: Release build and all 18 CTests passed (71.98 s), GUI-disabled core
1/1 passed (2.85 s), fresh .deb payload/executable Build ID checks passed, and all
four original fixture SHA-256 checks passed. Remote Debian/Ubuntu CI results
remain unverified from this environment.


Active native table SaveRecord is now routed through the same controller as the
UI Save button, including the direct interpreter statement. Other generic table
commands remain scoped/unsupported. See TABLE_DOCUMENT_RUNTIME.md for verified
SQL behavior, readonly and initialization-context boundaries.
