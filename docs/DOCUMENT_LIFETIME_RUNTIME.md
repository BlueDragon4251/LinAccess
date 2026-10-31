# Independent native documents, project close and Quit

OpenForm targets belong to the actual enclosing project window when one exists,
otherwise they are independent top-level Qt windows. They are never owned by the
invoking form. Closing or destroying that caller leaves its target open, with
actual records, CodeBehind state and the real project command bars intact.
Embedded subforms remain owned by their containing document.

Project switch and main-window close first validate all open native Access form
and report documents of that connection and database path. DoCmd.Quit and the
zero-argument Application.DoCmd.Quit path validate all open native Access documents
across connections before requesting Qt application termination. Unsupported Quit
options fail explicitly. This is bounded AccessFormPreview lifecycle support;
other native designers/editors and full Access application/window semantics remain
outside this document manager.

Closing first saves pending records, then checks every bound Unload. Cancellation
or a failing Unload returns false (Quit raises 2501) before any document is
logically closed. Already saved record changes are not rolled back. Accepted
closure dispatches Deactivate and Close, invalidates references, and closes native
windows. QObject deletion still waits for executing native VBA to unwind, including
nested Qt message loops. Closing guards clear on success as well as cancellation;
there is no pending-Quit flag which can survive an early return. Window close emits
finished once. Recursive close/Quit and OpenForm during project shutdown fail
explicitly rather than recursively closing or leaking another document.

## Evidence

- `document_lifetime_tests.cpp`: actual OpenForm then caller Close/destruction,
  surviving target/reference, SQL persistence before cancelled Unload, both Quit
  dispatch paths returning 2501, retained form/report references, project isolation,
  reentrant opening rejection, retry after cancellation, successful application-wide
  close, no execution past statement Quit, repeated close and deferred deletion.
- Existing close, embedded-form, calendar, command and self-close/MsgBox regressions
  remain mandatory. A direct window close additionally checks one finished signal.
- `report_corpus_tests.cpp`: unchanged original opener module SHA-256
  `8cde46e4573af8191abdbf9eb66d3993b640bbb249fee202872492d3457ad77b`
  opens the original target and closes its caller. Original Open/Load/Activate/Current
  results complete; the target survives destruction of the caller. Real project menu
  and toolbar context is retained. The native action uses a savepoint on a disposable
  29-table import; original MDB/Bridge bytes are untouched. Unexpected message boxes
  and timeout fail the regression. No original literals or values are logged.

The original regression exposed error 5 after initially detaching targets without
retaining their project UI owner: ActiveMenuBar was unavailable in the original
Form_Open. Owning targets through the project window repairs that runtime path.
This is distinct from merely instantiating the target without CodeBehind.

No Microsoft Access event/window oracle is available. Full MDI, acDialog and other
window/data modes, focus scheduling, cross-document event ordering, shutdown-time
creation through every API, designer save prompts and complete original sale/report
operation remain open. Native record saves before an Unload cancellation are
intentional; no transaction-wide shutdown rollback is claimed.

Validation (31 October 2026): final Release build and all 18 CTests passed
(48.26 s), GUI-disabled Release core 1/1 passed (0.39 s), current-executable
Build ID/payload/permissions checks passed for the fresh .deb, and all four
original SHA-256 checks passed unchanged. Remote distro CI remains unverified.
