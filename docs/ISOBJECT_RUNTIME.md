# Native VBA IsObject

IsObject now returns a Boolean based on the native VBA value's object-reference
kind. Nothing and invalid/closed stored object references remain objects;
Empty/Null/scalars/Error Variants/byte arrays do not. The intrinsic does not
resolve an object's scalar default or test whether its underlying resource is
still valid.

The argument parser preserves object context specifically for IsObject, including
Me!control, Controls(index), DAO/default-member references and stored object
variables. Explicit .Value/.Number member accesses still evaluate normally;
their errors propagate rather than being caught and converted to False.
Arithmetic still resolves scalar operands normally. Arity errors diagnose 450.

Core tests reproduce the previously unavailable intrinsic and verify Boolean
results, Nothing/Empty/Null/scalars/Error/byte arrays, zero default-getter calls,
normal arithmetic conversion and explicit getter errors. isobject_documents uses
actual Qt/VBA/DAO QueryDef parameters/SQL to check forms, controls, collection
lookup, Err, Recordsets, Fields, closed references, Set Nothing and explicit
closed-field Value error 3420 before SQL side effects. The same checks with a
pending native text draft preserve accepted Value/Text/source SQL and Undo.
Targeted core/native tests pass 2/2 (1.30 s); GUI suite now has 37 CTests. Full
validation is recorded in PROGRESS.md.

Full class modules, COM/ADO/external object instantiation, exact Windows parser/
coercion/error identities and full Forms/DAO object-model compatibility remain
open. This adds no additional complete original business workflow or universal
object validity check.

Reference:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/isobject-function
