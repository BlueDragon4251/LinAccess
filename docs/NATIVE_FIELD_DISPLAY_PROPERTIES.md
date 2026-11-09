# Native field captions and descriptions

Stand: 2026-11-09. This is an owned-project extension within F080/F121; it does
not infer original MDB field property blobs or claim the complete DAO catalog.

Own native fields store optional Caption and Description as actual project
metadata. Caption is a single line; both values are bounded to255 valid UTF-16
units without NUL. The field designer edits/clears these values separately from
schema changes, including counter fields. The metadata update is savepoint/CAS
protected and leaves physical names, values, counter state, index order and
relationships untouched. It can participate in the caller transaction.

DAO Database.CreateProperty supports these two names with Type=dbText and a
String Value, with DDL absent/False. Field.Properties supports named Item,
Append of an unattached created Property, Delete and Refresh. It works on detached
fields before append and actual own TableDef fields. Property exposes native
Name, Integer Type10, String/default Value and Inherited=False. Existing Value
is editable; Name/Type are readonly. Physical Recordset field properties can be
read and refuse mutation383. Missing supported properties report3270; unknown
families/property names are explicitly unsupported.

Each stored property has its own persistent identity, retained during edits and
field/table migrations. Held references refuse3420 after delete/recreate rather
than reading a replacement with the same name. Normal field/schema/closed/
rollback guards also apply. The complete built-in property list, Count, ordinal
lookup, enumeration, arbitrary user properties and data types remain unsupported;
no partial list is presented as a complete Microsoft Properties collection.
Original/imported/query property values are not guessed.

Native table datasheets use Caption as the displayed header (empty falls back
to the physical name), retain real SQL field names, and expose descriptions as
header/tool/status tips. Select/reopen refreshes this metadata. Existing form
wizard fields inherit table captions/descriptions while preserving source field
identity; labels are captured in the saved form definition. The report designer
captures captions for a proven simple physical-table source and uses them for
headers. Existing saved overrides remain independent; changing the source
rebuilds inherited report headers. Complex query/alias caption lineage and full
Access designer/event/lifetime equivalence remain open. Unsupported sources
retain their actual result-column labels rather than fabricated captions.

Validation: existing GUI1/1 (5.65s), including real field designer/datasheet/form
and report wizards. Then native type declaration/TypeOf registration was unified;
Core/DAO2/2 passed (4.66s) with actual typed VBA, readonly access, missing-property
errors, identity across delete/recreate, field/table rename and retained values.
App built; no new target or fresh combined3/3/full-suite/Windows claim.
