# Native document activation

Qt WindowActivate events on actual top-level AccessFormPreview windows dispatch
the bound Form_Activate (or Report_Activate) handler. Construction-time native
activation is remembered until Load completes. If Activate is bound, the first
Current waits for that activation; record selection before activation updates
the buffer without prematurely firing Current. Returning to an already opened
document fires Activate without inventing another Current.

The active document is a guarded pointer. Activating another native form/report
dispatches the preceding document's Deactivate before the new Activate. Generic
Qt windows and message boxes do not replace that document. Qt WindowDeactivate
alone is insufficient because it also covers leaving the application. Embedded
subforms do not participate as top-level active documents.

Closing runs the existing save/Unload cancellation checks first. Only accepted
closure deactivates the active document, then runs Close. Rejected closure keeps
the active-document reference and does not dispatch Deactivate/Close.

Regression tests use real Qt window activation and actual SQLite event records:
Open/Load/Activate/Current order, a switch between two forms, return activation,
no repeated Current, an intervening non-document window, cancelled closure,
accepted Unload/Deactivate/Close, and an embedded form without Activate.
The unchanged original Kassa form is shown/activated and its bound Open, Load,
Activate and Current procedures must all complete. Its remaining calculated
Null-conversion error is separately reported (CURRENCY_RUNTIME.md).

These are bounded document-event semantics. Resize/focus event families,
Access PopUp/Modal/MDI distinctions, navigator/query/datasheet activation,
all reentrant event combinations and exact nested child/parent close ordering
still require implementation or comparison against Microsoft Access. No
Microsoft Access event-trace oracle is available. Successful execution of this
startup section does not establish complete application compatibility.
