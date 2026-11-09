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
container/subform/option-group/custom-control creation, full list-header/callback editing,
font/theme/property inheritance or Windows event/rendering parity. Read-only query
forms retain the existing bounded write policy; this feature does not make them
editable. Full F080/F081/F083/F148/F155 requirements remain open.

The existing GUI fixture creates a real unbound form through both dialogs,
corrects a rejected duplicate name, uses Undo/Redo, reloads the saved project
object and drives actual text/calculated/label/button/checkbox/combo/list widgets.
Its native click writes a real isolated SQLite record and requeries the list;
independent SQL verifies that effect. A blank zero-control form is also decoded.
Final targeted validation is recorded in handoff/progress, without new CTest IDs.

Literal bracket/equals/outer-space field bindings are covered in
LITERAL_CONTROL_SOURCE_BINDINGS.md.

## Authored multi-column lists

The property editor also stores ColumnCount (1–255), BoundColumn (0–ColumnCount)
and ColumnWidths. Zero selects the native row index; other bound columns are
one-based. Widths use semicolon-separated twips or cm/mm/in/pt, including hidden
zero-width columns. Design-time validation and runtime rendering share the same
actual parser; the native authoring path rejects excess width specifications and
non-list controls with nondefault list metadata. Existing definitions without
these properties retain their one-column defaults.

The existing GUI/layout bundle passed 2/2 (13.83 / 12.51 seconds): an authored
combo displays beta while returning its separately hidden value 2; an authored
SQL list hides the first column and selects the actual computed second value 34
after the real button wrote 17. Save/reload, actual item models and mouse selection
are used. No new CTest, full Windows/original-workflow/full-suite or refreshed DEB
proof is claimed. F080/F083/F162 remain broad open requirements.

## Authored defaults, textbox formats and password masks

The control editor stores literal DefaultValue, TextBox Format and InputMask.
Default preflight and actual initialization share the same native literal parser:
finite numbers, quoted strings, Null and English True/False/Yes/No, optionally
prefixed with =. Defaults are capped at 255 UTF-16 units and are refused for
labels/buttons/calculated controls. Unsupported computed defaults are not executed
by the designer. Imported/runtime default behavior remains the existing bounded
contract (CONTROL_DEFAULT_RUNTIME.md).

Format and InputMask apply only to TextBox. Format metadata has the existing
1024-unit limit; full mask semantics are checked when an actual value is formatted,
not by inventing a validation value. InputMask accepts blank or Password; other
patterns remain an explicit boundary. This is display/echo policy, while Value
retains its actual typed content. The existing CONTROL_FORMAT_RUNTIME.md and
PASSWORD_INPUT_MASK.md limits still apply.

The existing GUI/layout bundle passed 2/2 (14.98 / 12.78 seconds), covering actual
creation/save/reload, initial numeric/Boolean defaults, zero SQL insertion from
defaults, computed padded display versus numeric Value, masked display versus
actual technical text and the subsequent real VBA/SQL/list interaction. No new
CTest, core/Windows/original/full-suite/installer execution is claimed. Full
F015/F071/F080/F081/F115 requirements and catalog totals remain unchanged.

## Authored state and tab order

A selected-control panel stores Visible, Enabled, Locked, TabStop and TabIndex.
The actual saved layout applies them through the existing native state/focus
implementation. Locked authoring is restricted to supported TextBox/ListBox;
labels have no tab controls. Boolean and integer ranges are checked before save.
Type changes display only meaningful state settings. Existing controls without
these fields retain the prior defaults. Equal tab indexes keep native stable
layout ordering; full Access automatic renumbering/event equivalence is open.

The existing layout/GUI bundle passed 2/2 (12.28 / 14.29 seconds). The actual
created/reloaded form has a hidden disabled label, locked masked textbox, excluded
tab stops and a real keyboard Tab transition from text input to its configured
combo target. Existing default/format/multicolumn/SQL checks remain in the same
fixture. No new test target, full-suite/original/Windows/installer run is claimed.
F080/F081/F104/F148 and catalog counts remain unchanged.
