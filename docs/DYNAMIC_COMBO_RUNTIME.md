# Dynamic native combo properties

Native ComboBox/ListBox adapters expose RowSource, RowSourceType, BoundColumn,
ColumnCount and ColumnWidths through actual VBA getters/setters. SaveAsText
ColumnWidths is retained. Assignment reloads the supported Table/Query source via
the native query executor. Invalid metadata or a failed source query preserves the
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

Only supported SQL/table/saved-query row sources execute. Value/field lists,
user-defined row-source functions, automatic dependency refresh, full combo event
ordering, Access multicolumn sizing/rendering parity and true ListBox presentation
remain open. This adapter is not general Access control compatibility. Tests mutate
only disposable SQLite data; original fixture bytes remain unchanged.

Validation (31 October 2026): final Release build and all 18 CTests passed
(48.26 s), GUI-disabled Release core 1/1 passed (0.39 s), fresh .deb Build ID /
payload / permissions checks passed, and four original fixture hashes passed.
Remote Debian/Ubuntu CI and a Windows Access oracle are not implied.
