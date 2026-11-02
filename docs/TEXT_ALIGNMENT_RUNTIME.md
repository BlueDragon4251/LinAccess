# Native text alignment

SaveAsText TextAlign is retained per control and inherited independently from
per-type templates. Explicit zero overrides a nonzero template. Malformed or
out-of-range imported values reject with generic diagnostics. Metadata supports
the documented 0..4 settings; native TextBox and Label rendering supports 0..3.

Native horizontal Left, Center and Right use actual QLineEdit/QLabel alignment.
General uses left for text, including numeric-looking strings, and right for
numeric/Currency values and retained numeric/date field types. Calculated values
use their existing cached type. No expression is evaluated just to choose an
alignment. Vertical alignment is preserved. Default initialization, record
navigation, requery and calculated-value refresh update the native alignment.

The form-control VBA TextAlign getter/setter uses the same instance metadata and
widgets. A scalar default resolves once; integral coercion uses existing banker
rounding. Null is 94, a type mismatch 13, Byte overflow 6, an invalid setting 5
and unexpected getter arguments 450. Unsupported control adapters report 438.
Distributed alignment (4) reports 3251 before a runtime assignment changes state;
a supported native control declaring it also rejects construction explicitly.
Runtime Report alignment setters report 3251 because the relevant report event
context has not been implemented. Static report alignment and getters work.

Changing alignment preserves pending Text/Value buffers, font, selection, focus,
data permissions and Dirty state; it does not save or rewrite SQL values. Styling
a control on a read-only query source does not enable data editing. This is
instance state, not a form-design writer. ComboBox/ListBox/button adapters,
distributed/RTL details, complete Date/Currency/Decimal fidelity, exact unbound
first-keystroke typing and pixel-for-pixel Windows rendering remain open.

## Evidence

Synthetic native Qt/SQLite tests cover templates/explicit zero, text versus
number/Currency/date fields, calculated controls, unbound defaults/type changes,
VBA button styling during pending user input, unchanged SQL until explicit save,
selection/focus and validation errors, one-time scalar conversion, navigation/
requery, read-only source permissions, fresh instance reset and report preflight.

An empty-database/no-CodeBehind probe hash-verifies original documents and checks
24 form/report layouts: 208 Labels and 226 TextBoxes with explicit Left/Center/
Right reach the native widgets and VBA getters. It executes no original business
procedure and establishes no original dataset or Windows rendering equivalence.
The initial targeted pair passed 2/2 (39.17 s); final validation is recorded in
PROGRESS.md.

## Public contracts

- https://learn.microsoft.com/en-us/office/vba/api/access.textbox.textalign
- https://learn.microsoft.com/en-us/office/vba/api/access.label.textalign

These references establish property settings and General alignment; they are
not an Access 2003 runtime oracle for exact error precedence or rendering.
