# Native Dictionary and indexed object writes (2026-11-04)

CreateObject("Scripting.Dictionary") and New Scripting.Dictionary create actual
independent C++ dictionaries. CreateObject supports the local known class only;
unknown classes/remote activation refuse 3251. This is a native compatibility
object, not general COM/ActiveX or reference-library activation. Use local Object
variables and scalar Private/Public/Global module Object fields. Module fields
start as Nothing, retain actual native object identity across calls within their
host and release references on Set ... = Nothing. Separate hosts own separate
module state. Typed Scripting declarations/reference management remain open.

Add/Exists/Item/Count/Keys/Items/Remove/RemoveAll, distinct-key renaming and
CompareMode operate on real ordered hash storage. Item reads of absent keys
create Empty entries; setters add/update actual values. Duplicate keys refuse
457. String and finite primitive numeric keys are supported; numeric and String
keys remain distinct. Binary String comparison preserves Unicode values; text
comparison is bounded to ASCII. Database/use-option/LCID comparison and other
Variant key families remain 3251. Missing Remove/rename targets and equivalent
normalized-key renaming retain explicit 3251 boundaries. Full Windows error
precedence and Variant-key equivalence are not claimed.

Keys/Items return independent zero-based Variant snapshots. Owned array metadata
permits their direct For Each use without repeating getters; arbitrary other
property/function array sources retain previous boundaries. Direct Dictionary
For Each enumerates keys and guards structural mutation. Loop frames retain the
enumerated object even if its source variable is cleared, releasing it at loop
exit. Existing value replacement remains possible. Native Collection/Dictionary/
array graph checks reject cycles and bound traversal; this is not a general COM
cycle collector. Dictionary allocation is bounded to one million entries.

Positional default indexed writes support d(key)=value and Set d(key)=object.
Arrays retain their existing checked numeric index path. Named indexed setters,
general Set member syntax and full object/default-member coercion remain open.
Stored arrays use native value-copy behavior; stored objects retain identity.

DAO Recordset default writes, Recordset.Fields(index) writes and stored Fields
collection default writes route to the existing actual Field.Value setter.
Edit/Update, writability, type validation, transaction and closed-context guards
therefore apply. A real separate SQL read sees the old value until Update;
after Update the actual keyed SQL row contains the new value. This does not
make guarded joined/unsupported recordsets writable.

Existing core/DAO targets cover real dictionary data, keys/snapshots/comparison,
array isolation, object references, cycle refusal, enumeration mutation and
native SQL execution/GetRows through dictionary-held DAO objects. An actual DAO
clone remains owned during enumeration after clearing the source variable and
is released on Exit For, observed through the existing native clone/transaction
guard. No new CTest target is added; no original data/source is used in these
synthetic cases. Complete Scripting/COM/Access parity remains open.

Current primary references, not an Access 2003 execution oracle:
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/item-property-dictionary-object
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/comparemode-property


Validation (2026-11-04): full desktop 48/48 (77.17 s), freshly rebuilt
GUI-free core 1/1 (1.20 s), targeted core/DAO 2/2 (1.44 s), Debian
payload/permissions/Build-ID/pinned SQLite and all four original fixture
hashes pass. Scoped checklist: 368 checked / 72 open; not total Access parity.


## Persistent scalar module Object fields (2026-11-04)

Private/Public/Global scalar Object fields initialize to Nothing and retain real
native references in per-host module state. Repeated Dictionary calls share
actual storage; separate hosts remain independent; clearing both aliases with
Set ... = Nothing releases bindings. Non-object native binding refuses 13.
Early-bound Scripting/external declarations and automatic New remain open;
bounded native DAO declarations are described in DAO_OBJECT_BINDINGS.md.
Existing core target covers repeated calls, identity, visibility, clear/recreate
and host isolation. Full desktop 48/48 (78.39 s), freshly rebuilt GUI-free 1/1
(1.09 s), targeted core 1/1 (1.05 s), Debian payload/permissions/Build-ID/pinned
SQLite and all four original fixture hashes pass. Scoped checklist 369 checked
/ 72 open; not total Access 2003 parity.
