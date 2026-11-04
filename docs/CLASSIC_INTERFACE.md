# Classic native interface

The GUI starts with Qt's portable Windows widget style when available. Menus,
buttons, tables, dialogs and toolbars remain actual Qt widgets with their existing
SQL/VBA behavior. The classic application font is eight points, using installed
Tahoma when available and otherwise the existing family. No Microsoft font is
bundled. This is a native classic baseline, not a pixel-exact Access-2003 shell,
Office command-bar model, MDI reproduction or verified Windows screenshot match.

**Ansicht → Klassische Oberfläche** switches between the classic style and the
original application style/palette/font. The choice persists as ui/classicStyle
in the normal per-user QSettings configuration. An unavailable Qt Windows style
leaves the existing style active and disables the action. No Windows runtime,
Wine, COM or new application dependency is involved.

Unspecified form canvas backgrounds inherit the application Window palette, so
the surface changes with the theme. Explicit section/control colors retain their
source values; report paper remains white. Complete runtime re-evaluation of all
imported system colors and arbitrary report styling invalidation remain open.

Validation uses the existing GUI cases under QT_STYLE_OVERRIDE=Windows, including
actual transparent/opaque composites, navigation, events and report pages. Six
prior assertions used a fixed Fusion gray; their transparent-surface expectations
now follow the selected Window palette, while explicit RGB checks remain fixed.
A real application startup smoke uses isolated XDG configuration/data directories.
No extra CTest target or work-catalog ID was introduced.
