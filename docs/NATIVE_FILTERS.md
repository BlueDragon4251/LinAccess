# Native form/report filtering

Filter and FilterOn are real document properties, reachable through Me and the
Forms/Reports object wrappers. The native filter editor and toggle use the same
controller. SaveAsText Filter/FilterOnLoad values are retained and applied when
opening a source with a nonempty saved filter. An empty saved expression does
not activate filtering.

Direct table sources filter actual SQLite records while retaining row identity,
typed field conversion and editability. Query sources remain read-only and the
filter is applied to their result. Named ranked TOP views retain their original
TOP result before filtering; a proven stable inline TOP 1 boundary also works.
General inline ranked TOP filtering is explicitly unsupported (3251).

Conditions use the shared Jet compiler: the regressions cover numeric predicates,
Is Null, date literals and ANSI-89 LIKE. Simple calculated record aggregates use
the filtered loaded rows. DSum/DCount retain their separate domain semantics.
Pending edits save through normal validation before a filter change; cancellation
returns 2501 and preserves the edits. Invalid filters restore the previous filter
state, selected record and available data. Disabling a broken saved filter can
restore the source. Syntax/incomplete-input failures report 3075.

Me.Requery and DoCmd.Requery reload actual data and dispatch bound Current after
successful refresh. An unbound form has no record source to requery, so that path
refreshes its layout/calculations without manufacturing another Current. This
distinction prevents the original unbound Kassa Current handler's two Requery
calls from recursively triggering itself. Native explicit SetFocus is also
recognized on Me. Full focus-specific DoCmd target selection remains incomplete.

## Evidence and limits

Native VBA/Qt tests check real record changes after Requery, event records,
filtered sums, UI controls, save cancellation, invalid-filter recovery, readonly
queries, saved/empty/broken filters and TOP result ordering. A hash-checked original
Kassa form additionally filters the actual imported table, compares twelve bound
textbox values with independent source records, then restores its full dataset.
That original-form check measures layout/data binding with CodeBehind absent;
VBA dispatch is tested separately and original startup procedures retain their
own unchanged-module regression. Original MDB/Bridge fixtures are not modified.

This is not full Access filtering/event equivalence. Sorting, advanced query
provenance/qualified criteria, implicit parameters, every criteria function,
all focus/reentrant/nested event sequences, report NoData/requery ordering and
persisting runtime filter changes through the designer remain open. There is
no Microsoft Access UI/event/result oracle. Existing row/budget limits still
apply; this package does not establish complete Kassa workflows.


Native ApplyFilter/ShowAllRecords and actual datasheet filter dialogs now share
the checked document controller. See [DATASHEET_FILTER_RUNTIME.md](DATASHEET_FILTER_RUNTIME.md).
Existing mainwindow_documents8.62s and access_2003_layouts11.77s pass; saved named
filters, complete events/locale/persistence and the Access oracle remain open.
