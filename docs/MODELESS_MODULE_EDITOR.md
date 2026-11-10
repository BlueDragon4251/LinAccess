# Native modeless OpenModule and module editor lifetime

DoCmd.OpenModule now opens a real modeless current-project standard-module
editor from both form/report and standard-module execution hosts. ModuleName/
ProcedureName are optional individually; at least one is required. Named/
omitted arguments are supported. A procedure-only request searches standard
module metadata; ambiguity requires an explicit module. Missing/unsupported
class/document/library contexts remain guarded3251 rather than claiming full
Access errors or library resolution. Opening performs analysis/navigation, not
procedure execution. Declarations and actual source-line navigation are used.

The existing editor builder is shared with the unchanged modal GUI entry. Its
callbacks own their captured widget/state values after the builder returns.
Modeless windows retain project/connection/generation tags, reuse an existing
window/draft and keep Save open. Imported standard-module source is displayed
read-only; Save/Delete/Run/Debug are disabled there. Original archive bytes are
not rewritten. A native module save updates only its loaded catalog revision;
concurrent changes refuse3420 and retain the draft. New modules use INSERT.

The editor refuses close while busy. Explicit Close(acModule,name,save) supports
Save/Discard/prompt; unrecognized prompt results cancel. Dirty source close
and project close can be cancelled. The existing project close guard closes
modeless module windows and blocks reentrant module opening while closing;
actual form/project closing gates are reused. QObject lifetime, deferred delete
and existing debugger ownership remain native Qt behavior. This is no complete
VBE/MDI/class/event-module/library/modal-scheduling or Windows error equivalence.

App and existing mainwindow target built; mainwindow_documents passed1/1 in
13.50s with the20s watchdog unchanged. Actual form VBA resumes after a named
procedure-only OpenModule; the standard-module host navigates/reuses the same
window and unsaved draft. Native loaded descriptors, source-line position,
Cancel, optimistic revision failure, Close Save/Discard, invalid arguments and
read-only technical imported source are checked. Existing modal editor save/
run/debugger/catalog checks execute in the same bundle. The first run exposed
a prompt result without a clicked button; both the actual Cancel-button fixture
and cancellation fallback were corrected. No new CTest/catalog identity;
456/386/70. Broader F052/F114/F135/F148 remain open.

Latest checked DEB/Original/inventory is source5c82c36546f744621cb1523e2c1d37d3fef15258,
tree20206cfc219378aadf8127937ceb3a9d2962cac8: contents/rights/build IDs/SQLite
passed; existing original/inventory2/2 in165.21s (118.77/46.44), seven protected
hashes unchanged. Inventory100/78/22, categories73/5/3/9/10 and no Access result
oracle. Known Original macro/keyboard/layout-PDF limitations remain. That DEB
predates this modeless editor work.

Public method contract, not an execution oracle:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.openmodule

## Open module selection and repaint

DoCmd.SelectObject(acModule,name,False) now activates an already-open current-
project module through both real hosts. It does not reopen a closed window,
replace its draft or set a false datasheet context. Connection/database/
generation checks, ambiguity and modal/closing guards remain. Activation clears
only minimization, retaining maximized/fullscreen state. Explicit/default repaint
recognizes active modeless module windows, using existing guarded widget paint
paths. Database-window selection and other object families retain their bounds.

App/GUI built; existing mainwindow_documents passed1/1 in13.18s with20s watchdog
unchanged. Actual form VBA and standard-module service verify selection,
minimize/maximize preservation, unchanged drafts, repaint and closed-window
refusal. An initial13.65s pass was followed by a focused correction/check for
window-state preservation found during review. No new CTest/catalog identity;
456/386/70. Latest checked DEB is still5c82c365, before OpenModule and this scope.
