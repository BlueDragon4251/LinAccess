# Imported native control fonts

## Implemented scope

SaveAsText FontName, FontSize, FontWeight, FontItalic and FontUnderline are
preserved for controls. Declared per-control-type font templates in the unnamed
defaults block seed subsequent instances; explicit instance properties override
individual fields without changing siblings. Template blocks remain metadata,
not additional visible controls. Unspecified properties keep the native Qt font;
no undocumented Access builtin font is invented.

TextBox, Label, CommandButton, ComboBox/ListBox adapters, CheckBox, OptionButton
and ToggleButton use real QFont attributes when constructed. Names request the
exported family; point sizes and weights are applied directly. FontWeight zero
uses native Normal; sizes outside 1..127, weights outside 0..1000, empty explicit
families and malformed Boolean values reject import with generic diagnostics.
NotDefault uses the existing false-default Boolean decoder for italic/underline.

Editable combo editors require their own font assignment: Qt otherwise retains
its explicit 9-point editor font even when the combo itself has 12 points. Both
the editor and popup use the imported font. A newly created multi-column tree
view receives it again during real RowSource requery. Fonts survive data-field
mirroring, SQL save and record navigation without changing Dirty/data semantics.

## Evidence

The synthetic native Qt/SQLite regression asserts template inheritance, per-field
overrides, sibling isolation, explicit false overrides, normal/semibold/bold,
italic/underline, combo editor and actual two-row requeried popup fonts. Real
font metrics show that the larger label occupies more horizontal space. A real
user edit is persisted to SQL and mirrored while different fonts remain intact;
a subsequent record switch loads the next value without altering typography.
Invalid font declarations fail preflight.

One hash-verified original form and one original report are instantiated with
an empty in-memory database and no original VBA; label font attributes reach
the native widgets. This loads no original table rows/credentials and establishes
no additional original procedure or report-dataset execution. Existing complete
layout/original-heavy targeted tests pass (35.10 s).

## Remaining scope

This is static imported typography, not complete Access visual parity. Missing
families resolve through the host font system; Microsoft fonts are not bundled,
and glyph metrics, DPI, clipping and pagination are not Windows-oracle matched.
VBA FontName/FontSize/etc. adapters, complete builtin/default-property inference,
Form/datasheet/root/section typography, OptionGroup/Tab/ActiveX/subform fonts,
rich text, alignment, colors, borders, themes and automatic sizing remain open.
The existing ListBox is still a bounded combo adapter, not a full Access ListBox.

Validation: host 21/21 (80.67 s), Debian 12 21/21 (68.45 s), Ubuntu 24.04
21/21 (60.40 s), GUI-disabled core 1/1 (0.92 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass.
