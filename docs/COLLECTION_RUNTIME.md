# Native VBA Collection subset

`Set items = New Collection` and `New VBA.Collection` create a native C++ ordered
collection. `Dim items As Collection` and Collection-typed module fields initialize to
Nothing and reject other native classes with error 13. Typed Collection function
results and required parameters are supported. ByVal shares the collection but
keeps reference replacement local; ByRef can replace the caller reference.
Module-held references persist across procedure calls. Existing Object
and Variant references can share the same collection.

Implemented members are Add(Item, Key, Before, After), Item(Index), Remove(Index)
and read-only Count (Long). Positional/named arguments and omitted optional Add
arguments are bound separately from an explicitly supplied Empty. Default Item
calls use one-based numeric indexes or string keys. ASCII keys compare without
case; duplicate keys report 457. Invalid indexes/keys and simultaneous Before/
After report 5. Items preserve Variant subtypes, Null and native object references.
TypeName returns Collection and VarType remains vbObject.

For Each iterates the ordered values through the real runtime. A scoped lease
rejects Add/Remove with 3251 before mutation while an enumeration is active;
normal completion, Exit For and unwinding release the lease. This deliberately
bounds the implementation until Access mutation-during-enumeration behavior is
available. Existing DAO/Forms enumeration behavior is unchanged.

Limits remain explicit: regional/non-ASCII keys, nested Collection items (also
inside arrays), automatic As New, Optional Collection parameters and
complete class/reference semantics are not implemented. Recursive collection
ownership is rejected to prevent shared-reference cycles. Item traversal and
allocation have bounded budgets. Unknown members and write attempts fail; this
is not a full VBA Collection or COM compatibility claim.

Regression checks reuse core and DAO targets: ordered/named/keyed access,
Currency/Date/Null values, aliases and mutation refusal; actual DAO query values
are collected and persisted by For Each, and a stored native DAO.Database
reference remains callable after retrieval.
