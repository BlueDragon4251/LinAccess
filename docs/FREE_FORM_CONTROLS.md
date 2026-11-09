# Native free form authoring

The existing form designer accepts a blank record source and zero selected bound
fields. Blank forms and bound forms containing only custom controls can be saved
as actual native project objects. A blank source is labelled Ungebunden rather
than an unavailable data source; it does not invent database records or expose a
new-row cursor. Existing bound field authoring remains available.

Steuerelemente opens a real property editor for up to 128 extra TextBox, Label,
CommandButton, CheckBox, ComboBox and ListBox controls. Names, captions,
ControlSource, twip geometry, list source type/source and an optional click binding
are stored in the native form definition and emitted into the existing native
layout/runtime adapter. Text fields and lists use separate Label controls. The
editor validates the complete draft before accepting it, keeps invalid edits in
the dialog for correction, and cancels without replacing the parent draft.

Controls support actual unbound input, calculated expressions, existing source
field bindings, list SQL/Value List rows and the existing native event/VBA path.
An empty click binding automatically uses a matching control-name procedure when
present. Existing supported Change/update/key/mouse procedures are also bound by
name; an explicit click macro/function uses the existing binding validator.
New controls use native identifier names (up to 64 characters). Duplicate names
or event prefixes, generated button names, missing/ambiguous source fields,
unsupported properties, noninteger/out-of-bounds geometry and unsupported source
kinds fail explicitly. Bound fields remain limited to 255. No unsupported control
kind, external callback or arbitrary Access layout behavior is synthesized.

The actual new controls and their properties participate in the form designer's
Undo/Redo and saved-baseline tracking. Authoring saves only the native project
object, preserving imported-original VBA bytes separately. Existing imported
SaveAsText and private original databases are not rewritten by this feature.

This is a native property editor, not complete Access WYSIWYG/toolbox, drag/resize,
container/subform/option-group/custom-control creation, multi-column list editing,
font/theme/property inheritance or Windows event/rendering parity. Read-only query
forms retain the existing bounded write policy; this feature does not make them
editable. Full F080/F081/F083/F148/F155 requirements remain open.

The existing GUI fixture creates a real unbound form through both dialogs,
corrects a rejected duplicate name, uses Undo/Redo, reloads the saved project
object and drives actual text/calculated/label/button/checkbox/combo/list widgets.
Its native click writes a real isolated SQLite record and requeries the list;
independent SQL verifies that effect. A blank zero-control form is also decoded.
Final targeted validation is recorded in handoff/progress, without new CTest IDs.
