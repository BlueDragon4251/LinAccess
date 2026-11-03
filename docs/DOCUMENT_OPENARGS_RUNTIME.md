# Native document OpenArgs

Native OpenForm accepts text OpenArgs through positional or named arguments.
The value is stored in the independent document before its Open/Load/Current
initialization. It is readable via Me.OpenArgs, the instance's unqualified
OpenArgs and Forms(name).OpenArgs. Native preview reports likewise expose text
OpenArgs during Report_Open and through Reports(name).OpenArgs.

Omitted/Null arguments yield a Null property; explicit empty text remains text.
Each newly opened document owns its own value. Closing and reopening permits a
new value. An already-open document retains its initial value; attempting to
replace it through OpenForm/OpenReport fails 3251 before filter changes.

Only text/Null input is currently supported. Other Variant types, OpenArgs
assignment, exact Empty/Null reopening behavior and Windows property mutability
remain open. The current property facade rejects assignment with 438.
OpenForm accepts acFormPropertySettings with acWindowNormal or acHidden; see
[hidden forms](HIDDEN_FORM_RUNTIME.md). Native icon/dialog modes are also implemented; see DIALOG_WINDOW_RUNTIME.md.
Native acFormReadOnly/acFormEdit policy overrides are now supported for newly
opened forms; see FORM_DATA_MODE_RUNTIME.md. acFormAdd and other view modes
still fail 3251.
This is not general mode parity.

The existing native GUI target opens filtered children, reads values in Open and
Load, closes/reopens with another value, checks a Report_Open value, Null on the
parent and the current assignment boundary. No new CTest target is added.
