# Dynamic native combo properties

Native ComboBox/ListBox adapters expose RowSource, RowSourceType, BoundColumn,
ColumnCount and ColumnWidths through actual VBA getters/setters. SaveAsText
ColumnWidths is retained. Assignment reloads supported Table/Query, Value List
or Field List sources. Invalid metadata or a failed source query preserves the
previous metadata, rows and selected value. Null assignments fail explicitly.

A QStandardItemModel holds actual typed rows and every configured display column;
a QTreeView popup displays columns independently. BoundColumn is one-based, with
zero selecting the zero-based list index. Column remains zero-based and retains
NULL/numeric source values separately from visible text. The first non-hidden
column supplies the displayed combo text; hidden bound IDs remain numeric values.
ColumnCount is bounded to 1–255 and BoundColumn to 0–255, with source-column checks.

ColumnWidths accepts semicolon-separated nonnegative TWIP values or explicit cm,
mm, in and pt units. Comma decimal separators are accepted; empty entries retain
Qt default sizing. Zero hides a popup column. Widths beyond the native display
limit and unsupported syntax are rejected. This does not implement all locale,
layout, rounding or automatic Access sizing rules.

`dynamic_combo_tests.cpp` runs native VBA property assignment, inspects the actual
three-column model and hidden/visible pixel widths, checks NULL/numeric Column
values and separate bound/display text, clicks a real popup row and independently
queries the persisted numeric record value. Failed assignments preserve prior
state. Source replacement, BoundColumn zero and dynamic hidden-column changes
have separate checks. Existing original initialization, Dropdown and typed-column
regressions remain mandatory.

Supported SQL/table/saved-query sources execute through the native query adapter.
User-defined row-source functions, automatic dependency refresh, full combo event
ordering and complete Access multicolumn sizing/rendering parity remain open.
Persistent single-selection ListBox presentation is now in NATIVE_LISTBOX_RUNTIME.md. This adapter is not general Access control compatibility. Tests mutate
only disposable SQLite data; original fixture bytes remain unchanged.

Validation (31 October 2026): final Release build and all 18 CTests passed
(48.26 s), GUI-disabled Release core 1/1 passed (0.39 s), fresh .deb Build ID /
payload / permissions checks passed, and four original fixture hashes passed.
Remote Debian/Ubuntu CI and a Windows Access oracle are not implied.


The first native mouse-arrow / Alt+Down opening now loads a supported row source
without a preceding VBA Requery/Dropdown call. Source failures leave the draft
value intact and can be retried after data-source recovery; they are not presented
as a successfully empty list. Loading choices does not prematurely validate or
commit typed search text against a numeric bound ID. Initial display-label loading
before any interaction and automatic dependent-source refresh remain open.
Native tests include the actual mouse move/click sequence, typed label input,
independent SQL persistence and failure/retry. No synthetic selection signal
substitutes for user input.

## Built-in value and field sources

Value List/Wertliste parses semicolon-separated items into ColumnCount-sized
rows. Quoted semicolons and doubled quotes are decoded; Unicode is preserved.
The current parser distinguishes unquoted blank cells/padded absent cells as
NULL from quoted empty text. This distinction and whitespace/localized-list
edge behavior still require comparison with Access 2003. Cells otherwise remain
text; the native bound field performs its usual numeric/Boolean/date conversion.
Empty sources produce zero rows. Malformed quotes and sources exceeding the
existing 10000-row control limit fail explicitly and preserve the previous model,
configuration and selected value.

Field List/Feldliste describes a table, saved SELECT or SELECT statement and
lists its projected column names in order. It prepares through the native SQL
adapter without fetching user records or evaluating result expressions; declared
parameters need no values for this metadata operation. Internal TOP rank columns
are excluded. Strict preparation rejects unknown quoted identifiers rather than
turning them into string literals, and restores the prior connection quote policy.
The same query-preparation/error path remains used by ordinary result execution.

`row_source_types_tests.cpp` exercises actual value-list popup mouse selection
and independently checks the persisted integer foreign-key value. It also covers
quoted delimiters/quotes/Unicode, empty/padded cells, failure rollback, row limits,
ordered field names, parameterized metadata and hidden-rank suppression. A SELECT
with an eager overflow expression can be described but still fails when executed,
proving the field list does not fetch/evaluate its data. These are native engine
and GUI regressions, not an Access 2003 rendering or edge-semantics oracle.

The original fixtures declare 17 explicit Table/Query source types and no explicit
Value/Field List type. No additional original project success is inferred from
these general built-in source tests. ColumnHeads, value-list AddItem/RemoveItem,
multi-selection/full ListBox behavior and custom row-source callbacks remain open.

## ComboBox ListCount

Read-only `ListCount` returns the actual loaded Qt model's row count. First access
loads the configured real row source without opening a popup or moving focus.
Value List, Field List and Table/Query use the same row-source loader; a loaded
model is reused until explicit Requery or source-property assignment. Empty lists
return zero. Arguments fail with 450; unsupported property assignment remains
438. Persistent single-selection ListBox also exposes ListCount through the shared
model; ColumnHeads and full ListBox behavior remain unsupported.

Native GUI/VBA/SQL regressions check initial lazy loading, preserved selection and
Dirty state, actual SQL INSERT of the VBA result, source changes, explicit Requery,
zero rows and invalid operations. The original source inventory contains two
ListCount tokens, both in comments. No executed original caller is claimed.

Contract reference (current documentation, not Access 2003 runtime verification):
https://learn.microsoft.com/en-us/office/vba/api/access.combobox.rowsourcetype
