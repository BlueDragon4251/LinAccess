# Report records and bound opening/closing events

The shared SaveAsText visual decoder now preserves Report versus Form kind,
report RecordSource/Caption and event bindings. Opening an imported report from
the navigator loads its actual Report_ module. Only [Event Procedure] bindings
run; the presence of a named procedure alone does not cause execution.

Implemented sequence in this bounded preview engine:

1. Report_Open(Cancel) runs before reading records. Cancel suppresses display.
2. The actual table, saved query or SELECT source is read through the bounded
   adapter. Report sources are read-only, including direct tables.
3. A successfully executed, empty bound source raises Report_NoData(Cancel),
   which can suppress display. A failed/unavailable query does not masquerade
   as an empty result and does not raise NoData.
4. Closing the actual Qt dialog invokes the bound Report_Close.

Explicit OnLoad/OnCurrent metadata uses the report procedure prefix. Report
format/print/page/group/subreport processing and complete Access event ordering
are not implemented. Close callbacks after opening/NoData cancellation are not
claimed equivalent to Access. The last result per bound event is available for
runtime diagnostics, including procedure-missing/unsupported/error outcomes.

Forms and Reports collections are separated by decoded document kind and scoped
to the project connection. Reports supports name/index lookup, Count and actual
object enumeration. Dirty/NewRecord and adding records are rejected for reports.
Bound read-only or calculated controls reject VBA value assignment before changing
the widget. Unbound values and captions can still be set by VBA. SQL data changes
through CurrentDb remain allowed against the project copy.

## Original artifact evidence

All three previously explicit original Report_Open regressions now also assert
automatic bound OnOpen and OnClose execution. Thus three original Open and three
original Close procedures complete through the native Qt host. The actual toolbar
actions are inspected, and temporary one-page PDFs remain covered separately.

A separate nine-report corpus test imports all 29 original Kassa tables and all
100 query artifacts into a temporary SQLite copy. It loads the original report
layouts without their VBA modules to measure record binding independently of
procedure execution. Current result: six report sources execute and bind 54
non-NULL textbox values. Values and NULL flags are compared directly with a
prepared SQLite reference reader. All reports stay read-only.

Three remaining sources reference controls on other forms. Localized SQL
Formulare!form!control lookup is now implemented, but these forms must be open
with valid inputs; the isolated report inventory does not invent those inputs.
They are explicitly failed
sources, not empty reports. The ignored private-report-inventory.json contains
artifact hashes, status and counts only; vba_executed=false is explicit for this
separate data inventory. Neither query execution nor binding proves agreement
with a Microsoft Access reference result.

Synthetic tests cover stored ranked TOP sources, direct-table read-only mode,
Forms/Reports separation, typed cross-object lookup, Open/NoData cancellation,
missing-source versus empty-source behavior, actual SQL event writes and guarded
bound-control assignment. Full report pagination and formatting remain open.
