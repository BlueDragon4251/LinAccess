# Native TextBox password input mask

TextBox InputMask metadata is preserved. Empty masks use normal native echo;
Password (case-insensitive) uses QLineEdit Password echo with immediate asterisks.
This changes the displayed text, not the actual Value/Text, scalar default or
SQL parameter. Null displays empty; Unicode values remain intact. Dynamic VBA
InputMask changes update the real editor and survive record navigation without
dirtying data. Null assignment raises 94; unsupported patterns fail with 3251
before changing the existing mask. Non-TextBox InputMask adapters are unsupported.

Other imported TextBox masks are retained, have an unsupported-mask diagnostic
and block user typing; generic Access mask tokens, literal-storage flags, blank
characters and validation are not implemented. Programmatic Value assignment
still follows normal source writability. Unsupported Password-prefixed patterns
also use masked echo while remaining read-only; this does not implement their
extra mask sections. No general InputMask compatibility is claimed.

## Evidence

Native Qt/VBA/SQL tests check actual displayText asterisks versus the original
Value/Text, typed edits and successful SQL persistence, dynamic reveal/remask,
unsupported/Null preflight, Unicode and Null/record-navigation behavior. A real
layout PDF is generated from the synthetic masked form. Its raster was manually
viewed and shows asterisks; the layout printer uses a widget bitmap, so empty
pdftotext output is not used as evidence that secret text is absent. No actual
credential or original record value is used in this rendering probe.

One hash-verified original form contains the Password mask. Its unchanged
layout is instantiated without original VBA in an empty in-memory database;
a synthetic probe verifies native masked rendering. No original credential
data or authentication is read/executed for this check. The separate original
login/column-boundary failure remains open. Other original masks are outside
this implementation; original business-workflow counts are unchanged.

This is display masking, not a change to the stored value. General clipboard,
IME, rich text, Windows rendering, mask validation and full Access input behavior
remain separate scope. Reference:
[TextBox.InputMask](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.inputmask).
