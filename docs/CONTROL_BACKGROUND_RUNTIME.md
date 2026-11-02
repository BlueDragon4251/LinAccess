# Native TextBox and Label backgrounds

SaveAsText BackColor and BackStyle now retain instance values and independent
per-type defaults. Explicit zero overrides template color/style. BackStyle is
Normal (1) by documented default, except OptionGroup metadata defaults to zero.
Invalid styles or overflowing/malformed color Longs reject import. Native
background rendering and object properties in this stage cover TextBox and Label
adapters, not option groups, buttons, combos, sections or whole-form backgrounds.

Normal sets the native interior color; Transparent (0) exposes the existing
parent canvas through the control interior, preserving the stored BackColor.
TextBox uses QPalette Base and Label Window plus its actual auto-fill policy.
Native active/inactive palettes preserve foreground, font, highlight/disabled
roles, text, focus, selection and data permissions. Existing borders remain.
Opaque and transparent pixels are checked in a composite native canvas image,
with device-pixel-ratio-aware sampling, rather than inferred from flags alone.

The shared OLE color decoder supports 24-bit BGR encoding plus native system
intents: window text (8), button text (18), window interior (5) and button face
(15), resolved from the application palette. Other OLE/palette encodings reject
3251. Missing BackColor getters return resolved native RGB for Base/Window;
Windows default OLE-constant identity is not claimed. Disabled-state transparency/
colors, live OS theme changes and exact Windows rendering remain open.

Native form-control BackColor/BackStyle setters resolve scalar defaults once
and use existing integral rounding. Null is 94, invalid types 13, Long/Byte
overflow 6, invalid style 5, unsupported color 3251, unexpected getter arguments
450 and unsupported adapters 438. Report runtime setters reject 3251 before
mutation because the relevant report event context is not implemented. Static
report backgrounds and getters work. These explicit runtime boundaries do not
establish exact Access error precedence.

Changing a hidden color while transparent does not make the control opaque;
switching back to Normal uses the retained color. Appearance changes do not
commit pending input or mark data dirty. SQL save, navigation, requery and
password style updates preserve instance settings. A newly opened instance
starts again from imported metadata. Read-only query fields stay read-only
while styling remains available. This is not a design writer.

## Evidence

Native Qt/SQLite regressions cover inherited/overridden colors and styles,
explicit black, actual opaque/transparent composite pixels, RGB VBA button
assignments during pending masked input, preserved selection/font/foreground/
focus and SQL before explicit save, unsupported setters, scalar nearest-even
conversion, system interiors, record navigation/requery, instance reset, query
permissions, report preflight and malformed metadata. Targeted pair: 2/2
(38.30 s).

An empty-DB/no-CodeBehind probe hash-verifies 55 original form/report layouts:
326 Normal and 587 Transparent text adapters, with 414 retained color values.
It checks native palette/auto-fill state and VBA getters without loading original
rows or executing business code. Original 32-module lexical property inventory
contains zero .BackColor and zero .BackStyle references; no additional original
business execution is claimed. Final validation is recorded in PROGRESS.md.

## Public contracts

- https://learn.microsoft.com/en-us/office/vba/api/access.textbox.backcolor
- https://learn.microsoft.com/en-us/office/vba/api/access.label.backstyle

Full Access geometry, disabled/selected rendering, section colors, frame/background
adapters for other controls and Windows visual parity remain open.
