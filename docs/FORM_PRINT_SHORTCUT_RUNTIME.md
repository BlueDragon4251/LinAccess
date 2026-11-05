# Native form Ctrl+P

Ctrl+P in the active top-level native form now opens its real QPrintDialog
through the existing current-layout QPrinter renderer. The toolbar print action
uses the same path. It does not install a QShortcut that could preempt Access
KeyPreview/KeyDown; a VBA KeyCode=0 cancellation prevents the dialog.

Pending control edits are validated/committed to the existing form draft before
the dialog; this does not promise whole-recordset saving or printing. A cancelled
control edit, close or active-window change prevents output. Native execution
and a scoped reentry flag protect the document during modal printing; no closed
form is rendered after dialog acceptance. Inactive windows and other modal
contexts do not print. Embedded-form print contexts remain explicitly bounded.

The output remains the existing single current form layout. Report shortcuts keep their existing commands and the report data/pagination
renderer is unchanged; forms and reports now share the guarded print-dialog
path. Full Access form data/
record/page-range/confirmation/subform/physical-printer/event parity still needs
implementation and a Windows reference. No physical job is sent by verification.

The existing access_2003_layouts target exercises actual Qt key delivery, a real
VBA KeyDown SQL side effect/cancellation, and a genuine print dialog which is
automatically rejected. No new CTest target or fixed catalog ID; F148/F135 remain
open, fixed456/384/72.

The shared dialog path retains report page setup and existing renderer selection,
checks closed/cancelled state after modal acceptance, and refuses inactive/reentrant
requests. Current form control confirmation remains form-specific. The existing
GUI target also opens/cancels the actual report Ctrl+P dialog using a real table
source; no printer job is accepted.

Validation: targeted layout/mainwindow2/2 (13.00 s), final full current
desktop48/48 (86.11 s), current Debian/build-ID/SQLite/four fixture hashes pass.
GUI-free core code is unchanged since navigation1/1 (1.92 s). No physical job
is sent; no new CTest target/ID. F148/F135 remain open, fixed456/384/72.
