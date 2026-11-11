# Optional automatic primary key in New Table

The real New Table designer offers “AutoWert-Primärschlüssel id anlegen”, enabled
initially. When disabled, only the supplied fields are created. An ordinary field
named id is then valid, without an invented AutoNumber generator or primary key.
With the option enabled, id stays reserved for the generated field and conflicting
input remains in the dialog for correction. A table consisting solely of the
selected automatic id is also supported. Without it, at least one actual field
is required. The existing DAO creation/savepoint/index/metadata paths are reused.

Other primary keys and indexes can be defined with the existing index designer.
This does not implement full Access table-design view or original-file writes.
App and existing GUI target built; bundled GUI scope1/1 passed in19.29s.
Actual default-key creation, rejected conflicting id with retained input, manual
id without primary/counter and automatic-id-only insertion are covered.
No new CTest ID/Windows oracle. Fixed456/386/70 unchanged.
