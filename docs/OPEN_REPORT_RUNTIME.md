# Native OpenReport preview

DoCmd.OpenReport and Application.DoCmd.OpenReport now open imported reports in
acViewPreview/acPreview (2), using their original report layouts, Report_ code
modules, real record sources, native read-only controls and the existing Qt
print-preview toolbar. WhereCondition uses the shared filter/query adapter.
Omitted arguments and acViewNormal/acViewPreview/acViewDesign/acWindowNormal
constants are recognized by the VBA expression runtime.

The report uses the same independent project-owned document lifetime as forms.
It remains open when its caller closes. Reopening a live same-name instance
reuses it; a nonempty condition updates the native filter. This is a bounded
behavior, not a claim of identical Access existing-instance/window scheduling.
Open/Load/Activate/Current/NoData failures propagate and temporary failed
windows are disposed; activation happens during showing. Open/NoData
cancellation returns 2501. Missing report/wrong layout kinds return
2103 and missing record sources retain 3078, rather than pretending to be empty
reports. Handler errors retain their actual runtime code.

Native tests check direct and Application calls, omitted arguments, real filtered
values, read-only state, Report_Open/Close/NoData SQL effects, instance reuse,
caller-independent lifetime, stale wrappers, unavailable sources, handler errors
and validation before any opening effects. Existing form-opening regressions
remain mandatory because creation/lifetime code is shared.

The unchanged original module
6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8
executes its original preview handler directly and through its bound button. Its
original target report layout
30f1ebe8559d44c1ca9f965948b55e556460d560286f669cac8f43a298741a6f
opens against an isolated actual MDB import. Displayed first-row fields/Nulls
are compared with a separate prepared SQL reader. Layout/module bytes are hash
checked; the test enables/shows the caller button to exercise its binding. No
authentication/full-startup or Windows report-result equivalence is claimed.
The report corpus setup now includes report artifacts in its project catalog,
matching the real import path instead of relying only on direct constructors.

## Remaining scope

Default/acViewNormal is direct printing in Access. It is rejected with 3251,
never silently changed into preview. Design/other views, named filters, non-normal
window modes and nonempty OpenArgs are also explicitly unsupported. This adapter
retains the existing bounded report renderer; complete sections/grouping/
pagination, print jobs, Access event/error/window ordering and original output
comparison remain open. Four original direct-print calls are not counted solved.
