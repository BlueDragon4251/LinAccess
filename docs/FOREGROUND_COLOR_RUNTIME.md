# Native foreground colors and RGB

SaveAsText ForeColor is retained as an optional signed Long, with per-control-
type template inheritance and explicit zero overrides. Malformed/overflowing
imported decimal values reject with generic diagnostics. Native text-control
adapters decode ordinary 24-bit OLE colors (red in the low byte) and the two
original system-color intents: window text 0x80000008 and button text 0x80000012.
The shared decoder additionally supports window-interior (5) and button-face
(15) intents introduced with control backgrounds. System colors resolve from the
current application palette, independently of
previous instance overrides. Other OLE/system/palette-index encodings explicitly
report 3251; no Windows palette is invented.

TextBox, Label, CommandButton, ComboBox/ListBox adapters, CheckBox, OptionButton
and ToggleButton receive real active/inactive Qt text palette colors. Combo
editors and native popups receive the same color again after model/view requery.
Disabled and highlighted-text roles are preserved. System-color intents are
resolved at application/update time; live OS/theme parity is not established.

Native form-control ForeColor reads/writes instance metadata and widgets. An
exported system value remains a signed Long in the getter. When no metadata
value was supplied, the getter returns the resolved native RGB color; this does
not claim Windows default OLE-constant identity. Scalar defaults resolve once;
existing integral rounding applies. Null is 94, type mismatch 13, Long overflow
6 and invalid/unsupported OLE encoding 3251. Getter arity is 450, unsupported
adapters 438, closed references 2450. Report runtime setters reject 3251 before
mutation because the needed report context remains unimplemented; static report
rendering and getters work.

The native VBA RGB(red,green,blue) intrinsic returns the corresponding packed
color. Channels use existing scalar/numeric conversion and nearest-even rounding;
values above 255 clamp to 255 under the documented contract. Empty becomes zero,
Null is 94, invalid types 13, negative rounded channels 5, nonfinite values 6 and
wrong argument count 450. The same primitive now also backs the native SQL adapter (SQL_RGB_RUNTIME.md).
QBColor,
complete typed-Integer coercion/error precedence and Windows oracle comparisons
remain open.

Color updates preserve fonts, password echo, pending Text/Value buffers, native
selection/focus, Dirty and SQL permissions. Data does not save merely because
appearance changes. A native canvas palette replaces the former background-only
CSS rule, retaining RGB(240,240,240) while avoiding stylesheet palette resets.
Password styles are not redundantly reapplied, and explicit foreground colors
are restored after control-state updates. Runtime styles are not design writes.
Full disabled/selected-state Windows rendering, other OLE colors, backgrounds,
borders and complete Access visual parity remain open.

## Evidence

Native Qt/SQLite tests exercise template inheritance, explicit black overrides,
BGR channels, custom application system palettes, RGB VBA button assignments
during pending masked edits, actual green label pixels and unchanged gray canvas
pixels, error preflight, scalar conversion, palette-role preservation, actual SQL
save/navigation, popup model requery, fresh instances, read-only data protection
and report preflight. Standalone core RGB tests cover channels, clamping, rounding,
Empty/Null/error/type/arity and three one-time scalar reads.

The initial targeted trio passed 3/3 (41.97 s). An empty-DB/no-CodeBehind probe
hash-verifies 45 original form/report layouts and checks 271 RGB controls and 277
system-color controls through real palettes and getters. The 32 original VBA
modules have zero lexical RGB calls and zero .ForeColor references; these counts
are an inventory, not additional business execution. Final validation is recorded
in PROGRESS.md.

## Public contracts

- https://learn.microsoft.com/en-us/office/vba/api/access.textbox.forecolor
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rgb-function
- https://learn.microsoft.com/en-us/windows/win32/api/olectl/nf-olectl-oletranslatecolor
