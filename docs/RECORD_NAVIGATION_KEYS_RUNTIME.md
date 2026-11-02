# Native record-navigation keys

Ctrl+Page Down and Ctrl+Page Up now dispatch the existing native next/previous
record commands (65/66). This preserves the current control rather than moving
between fields. Native form/control KeyDown events run first; setting KeyCode=0
cancels navigation. SendKeys uses the same Qt keyboard route.

The existing GoToRecord path commits/validates pending control text and saves a
changed record before navigation. A cancelled field or record BeforeUpdate keeps
the current record/draft. Successful navigation refreshes the actual bound Qt
controls and raises Current. Navigating from the last record to a clean new row
does not insert an empty SQL record; AllowAdditions and record boundaries retain
the existing native command guards. Closing from Current uses the existing
native document-lifetime guard.

The fixture-independent record_navigation_keys regression checks real text and
ComboBox focus, form/control key cancellation, field/record cancellation, SQL
persistence, Current counts, SendKeys in both directions, first/last/new-row
boundaries, disabled additions and self-close from Current. Plain Page Up/Down
and Ctrl+Shift variants do not invoke these record commands.

This validates bounded single-form navigation over a fully loaded direct-table
source. Multi-page form scrolling, continuous/datasheet/subform views, popup
interaction, every source/error/event boundary and exact Windows Access 2003
keyboard identity remain open. Unmodified Page keys require page-layout
semantics and are not claimed as equivalent record-navigation shortcuts. This
adds no complete original business workflow or blanket keyboard compatibility.

Targeted navigation/visibility/original-layout/Kassa-DAO tests pass 4/4 (49.62 s).
GUI suite now has 34 CTests. Full validation is recorded in PROGRESS.md.

Reference:
https://support.microsoft.com/en-us/accessibility/access/keyboard-shortcuts-for-access
(form navigation and forms with more than one page)
