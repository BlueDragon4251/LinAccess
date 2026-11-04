# Native Access FileDialog pickers

Application.FileDialog supports FilePicker3 and FolderPicker4 with actual modal
QFileDialog windows. Four MsoFileDialogType constants are available as Long;
Open1/SaveAs2 are explicitly unsupported, matching the Access API restriction.
Show returns Long-1 for acceptance and0 for cancellation. It only selects paths;
it does not execute/open a selected document or copy its contents into artifacts.

The bounded adapter supports Title, InitialFileName, ButtonName, FilterIndex and
AllowMultiSelect for file selection. Folder selection is single-directory. Filters
supports Clear/Add (including named Description/Extensions/Position), Count and
one-based Item with read-only Description/Extensions. Patterns are limited to
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
