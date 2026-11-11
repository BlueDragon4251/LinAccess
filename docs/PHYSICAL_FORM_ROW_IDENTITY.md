# Physical form row identity

Direct-table forms choose an available, unshadowed SQLite row identity from
rowid, _rowid_ and oid, and prove the qualified column can be prepared against
the physical main table. A user field named rowid is never used as the hidden
identity. If no hidden identity is available, the source refuses3251.

Reads, bound UPDATE/INSERT RETURNING and record DELETE use the same identity.
Before persistence, the existing physical-source guard also verifies that the
chosen identity has not changed. Requery is required after a schema change that
would reinterpret a loaded row identity. The projection alias also avoids actual
field names, and DataEntry session restrictions address that alias.

This prevents edits/deletions from matching multiple records that happen to have
the same user-defined rowid value, and permits NULL user rowid fields when adding
records. It does not claim arbitrary query-form updates, external concurrent
rekey equivalence or full Access row-lifetime semantics.

Existing GUI scope covers actual edit/Save, confirmed record Delete, DataEntry
insert with rowid/identity-alias collisions and refusal when all hidden names
are shadowed. Existing GUI passed18.36s, DAO/model final2/2 in6.19s (5.74/0.45);
app/helper built, seven protected hashes unchanged. No new ID;
F056/F057 remain open, fixed456/386/70 unchanged.
