# Native ComboBox ListIndex and ItemData

ComboBox.ListIndex reads the selected zero-based row of the actual native model;
an empty or unmatched selection returns -1. ComboBox.ItemData(index) reads the
specified row's bound-column QVariant, independently of the displayed column or
current selection. BoundColumn=0 returns the row ordinal; numeric/text/NULL and
zero-length text values retain the existing native Variant conversion scope.
The first read lazily loads the real row source. Both properties are read-only
(383 for native assignment); unknown control adapters use 438. ListIndex accepts
no arguments; ItemData requires one index (450 otherwise). The index scalar is
resolved once, rounded with the existing integer conversion, checked as signed
Long (6 overflow), with Null 94 and incompatible text 13. Invalid native row
indices use 5; exact Windows invalid-index error identity is not established.

ListIndex is limited to native Form view; report view rejects with 3251. Its
value during an uncommitted editable text draft also rejects with 3251, because
exact Windows selection timing has not been established. ItemData can still
inspect stored row data during that draft without committing it. The differently documented ListBox property and multi-selection API now have
their own bounded adapter in MULTISELECT_LISTBOX_RUNTIME.md.
Full declared Long/Variant subtype identity remains open.

The regression exposed an older null-selection bug: string matching could pick
a NULL-key list row when assigning Null. Typed load/assignment/mirroring/requery
now pass null state into the widget value setter, clearing combo selection and
text for Null. A non-null empty string can still select its own empty bound row;
NULL rows are not matched merely because their string conversion is empty.
This is a bounded native normalization, not a claim about all Windows duplicate
or NULL-key selection/event timing.

Actual GUI/VBA/SQL tests cover lazy loading, selected/absent indices, numeric,
text and NULL ItemData, BoundColumn changes including zero, scalar evaluation,
rounding/error/read-only paths, report limits, pending drafts, row removal and
requery. Null and empty text yield distinct selected indices despite identical
empty displays. SQL records remain unchanged. Targeted main-window/layout/
original-heavy trio: 3/3 (40.11 s). No new original business execution is claimed.

Public references:
- https://learn.microsoft.com/en-us/office/vba/api/access.combobox.listindex
- https://learn.microsoft.com/en-us/office/vba/api/access.combobox.itemdata

ColumnHeads is now implemented within COLUMN_HEADS_RUNTIME.md. Its documented
ListCount includes a visual header row independently of selectable model data;
exact Windows behavior at that extra index boundary remains unverified.

## Bounded Value List AddItem and RemoveItem

Native ComboBox/ListBox methods now mutate the actual Value List row source and
rebuild the Qt model. AddItem(Item, Index) accepts one semicolon-separated row,
appends by default or inserts at an explicit zero-based index. Existing quoted
semicolons/quotes, Null cells and empty text cells are serialized distinctly.
RemoveItem(Index) supports numeric row indices. Both support native named arguments.
Changing choices preserves the existing bound value and does not mark a record
dirty. Source text is restored if rebuilding fails.

This first bounded scope requires Form view, Value List/Wertliste, no header row,
and a value-bound column. Active combo popups, selected ListBox rows, removal of
the selected combo row and mutation during that control's update are refused3251.
RemoveItem text-key matching, empty AddItem strings, single Null-only one-column
serialization, ordinal-bound/header cases, full event/reentrancy/error precedence
and exact Access2003 behavior remain open. Wrong indices5, Null94, arity450 and
existing row limit3035 use native errors. F162 is not closed.

Existing layout target exercises real combo append/insertion/removal, quoted
multicolumn cells, preserved bound value/clean record, and real ListBox VBA calls
with reversed named arguments. Final target1/1 passes (10.66 s); updated Debian/
build-ID/SQLite3.54.1 and four original hashes pass. Headless runtime unchanged
from FileCopy1/1 (1.69 s), last broader full suite48/48 (87.98 s). No added test
target or catalog ID; fixed456/384/72 unchanged.
