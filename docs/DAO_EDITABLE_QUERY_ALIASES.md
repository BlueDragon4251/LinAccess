# Native DAO editable physical-field aliases

Single-table query dynasets now keep separate ordered physical and visible field
names. Aliases affect Fields/default-member lookup, Name, OrdinalPosition and
visible schema properties. UPDATE/INSERT, generator lookup/reservations and
AutoNumber edit protection always use the physical field. SourceField/SourceTable
retain that physical origin. Clones preserve both mappings. Physical schema types,
selected literal/Date defaults, Size, Properties and binary chunks remain typed.

The common editable-query validator additionally prepares WHERE against the main
physical table with SQLite double-quoted-string fallback temporarily disabled.
No rows or user functions execute in this preparation. A SELECT-output alias
which is not a physical WHERE field cannot become a different writable predicate.
Output aliases ambiguous under the native Qt lookup fold refuse3251. This is a
bounded ambiguity guard, not complete Access/Unicode/locale identifier parity.
Duplicate physical fields, complex/parameter/keyless targets, filtered AddNew
and omitted generator reservations retain the previous boundaries. Aliases do
not grant unproven original Fields.Properties or native Currency CreateField.
The latter remains unsupported; exact Currency alias coverage uses an actual
physical CURRENCY_TICKS table, independently checked as integer storage.

The shared application/DAO/query-model/Currency/layout build passed. One existing
six-check bundle: five passed, DAO fixture failed on an unproven raw-field
Properties assumption; that fixture was corrected to retain3251 and use native
schema for Property evidence. Only DAO was retried; a second fixture boundary
exposed unsupported native Currency CreateField, so its alias assertion moved to
the actual supported physical storage path. Only affected DAO was retried again,
passing1/1 in3.61s. Query model0.20s, Currency0.09/0.06/0.18s and layout13.78s were
not repeated. Actual VBA edits an aliased field by hidden key. Native schema
covers aliased counter allocation/protection, Single/Date defaults, Text size/
Properties and BLOB AppendChunk; SQL proves physical values and exact Currency
ticks. No new CTest/catalog IDs;456/386/70 unchanged; F054/F055/F121/F175 remain
broader open requirements. No Windows result oracle or full query parity claim.
