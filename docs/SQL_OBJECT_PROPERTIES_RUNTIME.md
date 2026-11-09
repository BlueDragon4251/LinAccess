# Live scalar object properties in Jet SQL

The existing SQL object-value path now also recognizes Forms!Document.Property
and Forms!Document!Control.Property, with corresponding Reports paths, bracketed
names and the existing Formulare/Berichte spellings. Literal dotted control names
remain separate from property names. Existing default field/control values and
explicit .Value continue to use the original three-argument scalar adapter.

A separate connection-local property resolver addresses actual native form,
report, control and form-record-field objects. The SQL compiler passes names as
escaped SQL string literals, and uses SQL NULL to identify the document itself;
it does not concatenate runtime values into SQL text. SQLite functions remain
nondeterministic, so views read current properties when executed.

Only an explicit set of existing scalar getters is admitted. Document getters
include Name, Caption, Visible, Tag, Dirty/NewRecord/CurrentRecord, OpenArgs,
Filter/FilterOn, OrderBy/OrderByOn, RecordSource, supported record-policy flags,
and existing page/window/control-policy properties. Control getters are bounded
to Name, Caption, Tag, Visible, Enabled, Locked, Value, OldValue, ListCount,
ListIndex and ControlType. The actual object's own constraints still apply;
being listed does not give a TextBox a ListCount or a record field a Caption.
Unknown properties and action members such as Undo, Requery, Dropdown and
SetFocus refuse438 before dispatch. Object-returning getters and deeper chains
are not admitted as SQL scalar properties.

The GUI resolver requires its GUI thread, matching connection and project path,
matching form/report kind and an available actual document. Ambiguous matching
names refuse5 for value and property lookup. Missing forms/reports retain2450/
2451; missing controls/fields retain2465. Getter completion rechecks document
liveness. No resolver gives3061. The new property hook preserves the existing
standalone value-resolver API and driver-owned close/reopen behavior.

Scalar conversion reuses the original checked object-value bridge: SQL NULL,
Empty=0, Boolean True=-1, finite numbers, text, canonical Date and bounded legacy
Currency storage. This does not establish exact Date/Currency Variant metadata,
physical Decimal/BLOB support or general arbitrary object access. Parameterized
properties, subform chains, method invocation, implicit undeclared parameters,
full Access error/event/locale semantics and a Windows query-result oracle remain
open. F067/F068/F118/F164/F050/F052 and the fixed456-point total are unchanged.

Existing SQL-object, GUI, layout and original-corpus checks pass4/4 (115.92s),
respectively0.02s/9.01s/11.68s/115.91s. Actual controls and record drafts, readonly
method refusal, live view values, independent connections, ambiguity, missing
forms/reports and the original SQL-reference regression are exercised. No new
CTest target, full-suite rerun or Windows execution oracle is claimed.
