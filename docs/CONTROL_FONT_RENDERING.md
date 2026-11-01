# Imported native control fonts

## Implemented scope

SaveAsText FontName, FontSize, FontWeight, FontItalic and FontUnderline are
preserved for controls. Declared per-control-type font templates in the unnamed
defaults block seed subsequent instances; explicit instance properties override
individual fields without changing siblings. Template blocks remain metadata,
not additional visible controls. Unspecified sizes use the documented 8-point default (10 for command buttons);
weight defaults to 400 and italic/underline to false. Unspecified families retain
the native Qt family; no undocumented Windows family is invented.

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

## Native runtime properties

Supported form text-control adapters now read/write FontName, FontSize, FontWeight,
FontBold, FontItalic and FontUnderline through the actual VBA object path.
Getters return configured Qt family, point size, weight and flags. FontBold reads
false below 700 and true from 700, and sets 700/400. Qt itself may classify 600
as bold; its bold() flag is deliberately not used for the VBA getter.
FontWeight zero renders/reads as native Normal. A pixel-sized native font without
a point size reports 3251. This is not a Windows font/default/type oracle.

Assignments resolve scalar defaults once and validate before changing metadata
or widgets. Numeric properties use integral rounding and Integer range checks
(overflow 6), then point-size/weight range checks (5); empty family is 5, Null is
94, invalid scalar types 13 and getter arguments 450. Unsupported control kinds
report 438; closed document references retain 2450. Report font assignments require
OnPrint under the documented contract. That event context is not implemented,
so report setters reject with 3251 before mutation; imported report fonts and
read access remain available. Boolean coercion reuses the
existing runtime rules. Runtime styles remain instance state, not a design writer.

The shared renderer updates combo editors/popups as well as the main control.
Appearance changes do not require writable SQL data and do not commit or discard
pending Text/Value buffers, move focus/selection, mark data dirty or reveal a
Password mask. Requery and record navigation retain the instance's new font.

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

This is bounded native typography, not complete Access visual parity. Missing
families resolve through the host font system; Microsoft fonts are not bundled,
and glyph metrics, DPI, clipping and pagination are not Windows-oracle matched.
Complete builtin/default-property inference,
Form/datasheet/root/section typography, OptionGroup/Tab/ActiveX/subform fonts,
rich text, alignment, colors, borders, themes and automatic sizing remain open.
The existing ListBox is still a bounded combo adapter, not a full Access ListBox.

Static increment validation: host 21/21 (80.67 s), Debian 12 21/21 (68.45 s), Ubuntu 24.04
21/21 (60.40 s), GUI-disabled core 1/1 (0.92 s). Fresh host/distro DEBs
pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass.

Runtime increment evidence: an actual bound native VBA button reads/writes all
six properties; real pending Text/old Value, SQL, selection, focus, asterisk echo,
combo requery and sibling styles are asserted. Errors are handled in VBA and
verified before mutation. A scalar-default probe is read once and integral ties
round as expected. A second instance starts from original styles; its real SQL
view stays read-only while typography changes without dirtying rows. A generic
setter updates a hash-verified original form label; an original report label
rejects assignment with 3251 and retains its font. No original VBA or table rows
are executed/loaded by these probes. Targeted tests pass
(36.43 s); original business-handler execution counts remain unchanged.

Microsoft contract references: [FontBold](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.fontbold),
[FontWeight](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.fontweight),
[FontSize](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.fontsize),
[FontName](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.fontname).
These document the 700 threshold, 8/10-point defaults and report OnPrint setter
restriction. They are not a Windows Access 2003 execution differential. Tests
explicitly cover 600/699/700 and the unexported size defaults. Full report
Print/Format timing and typography mutation remain open.

Runtime increment validation: host 21/21 (70.55 s), Debian 12 21/21 (69.71 s),
Ubuntu 24.04 21/21 (63.59 s), GUI-disabled core 1/1 (0.89 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass. These measurements follow the documented
Bold/default-size/report-Print corrections above.
