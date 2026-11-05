# Native Access FileDialog pickers

Application.FileDialog supports FilePicker3 and FolderPicker4 with actual modal
QFileDialog windows. Four MsoFileDialogType constants are available as Long;
Open1/SaveAs2 are explicitly unsupported, matching the Access API restriction.
Show returns Long-1 for acceptance and0 for cancellation. It only selects paths;
it does not execute/open a selected document or copy its contents into artifacts.

The bounded adapter supports Title, InitialFileName, ButtonName, FilterIndex and
AllowMultiSelect for file selection. Folder selection is single-directory. Filters
supports Clear/Add (including named Description/Extensions/Position), Count and
one-based Item with read-only Description/Extensions. Native For Each over Filters
returns actual typed filter objects in insertion order, including Add Position;
an empty/cleared collection enumerates no objects. Patterns are limited to
supported wildcard extension lists; unsafe/unhandled delimiters refuse explicitly.
The Windows all-files pattern *.* maps to Qt's * so dotless files remain visible.
Filters metadata feeds the actual Qt name filters; file modes feed actual Qt
existing-file(s)/directory modes. Other Office properties/methods are unsupported.

SelectedItems is a read-only, one-based collection of actual Qt-selected paths,
with Count, Item/default item and native For Each enumeration. The current native
context keeps one picker object, preserving configured properties across calls
and changing its requested picker type. Each Show clears the previous result;
cancellation yields an empty current-result collection. This is a per-document
native context, not a claim of complete Office application-wide singleton,
Windows directory/current-directory side effects or cancel-result equivalence.
All objects guard a weak Qt context; retained collections fail2450 after closure.
No Qt picker window, native host or document pointer is retained in selected paths.

The closed native type registry explicitly includes Office.FileDialog,
Office.FileDialogFilters, Office.FileDialogFilter and Office.FileDialogSelectedItems
(and their short aliases). Typed declarations/TypeOf use actual native identities,
not TypeName guesses. Full Office/COM/reference interfaces remain open.

Existing actual VBA/UI cases select a temporary real text file with a configured
filter, enumerate its selected path, cancel another Show, select an actual temporary
directory and check a retained collection after context closure. No selected-path
result is injected into the adapter. Named arguments, typed Office slots, filter
metadata, one-based items and Show results execute through the normal runtime.
No new CTest target or fixed-catalog item. Windows Access2003 UI/default/error/
filter/case/namespace/lifetime parity has no execution oracle and is not claimed.

References: [Access FileDialog](https://learn.microsoft.com/en-us/office/vba/api/access.application.filedialog),
[Show](https://learn.microsoft.com/en-us/office/vba/api/office.filedialog.show),
[SelectedItems](https://learn.microsoft.com/en-us/office/vba/api/office.filedialog.selecteditems).

## Explicit numeric filter deletion

`Filters.Delete filter:=position` now removes an actual one-based numeric entry
from the shared filter collection. The native picker uses the resulting list on
its next Show. Deletion before the current selection adjusts its ordinal to
retain the selection; deleting the final entry resets the bounded native index1.
This index policy is native behavior, not a Windows Access2003 oracle result.
Delete is refused while the modal picker is active. Invalid positions raise5,
Null94, excess arguments450; omitted/Missing/non-numeric selectors explicitly
raise3251 pending Office default/name/object-selector evidence. No assumed
Clear-all behavior is substituted for an omitted selector.

The existing actual VBA/UI case deletes before/at selection, checks the empty
collection and invalid-position error without mutation, then shows the real Qt
picker after removing a CSV filter. The actual dialog contains only the retained
text filter. Targeted layouts1/1 (10.58 s), updated package/build-ID/SQLite3.54.1
and all four original hashes pass; no new CTest target or catalog item. Last full
baseline48/48 (83.85 s), fresh headless runtime1/1 (1.79 s);456/384/72 unchanged.
Reference: [Office Delete](https://learn.microsoft.com/en-us/office/vba/api/office.filedialogfilters.delete).
