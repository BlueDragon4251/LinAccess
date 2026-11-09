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

Existing native intrinsic metadata is now also reachable through Property.Value:
Name, Type, Size, Attributes, DataUpdatable, OrdinalPosition, Required,
AllowZeroLength, DefaultValue, ValidationRule and ValidationText, whenever the
underlying field getter is actually proven. Values/types and write guards route
through that same context; counter flags remain17 for TableDef and49 for writable
Recordset. Property setters retain their own shared family after a supported
field change. Built-ins cannot be deleted; full catalog/Count remains unsupported.
DAO1/1 passed (2.46s), app built; no repeated GUI/original import or full-suite claim.


Owned native fields now also retain optional Text Format properties through the
same checked detached/attached DAO Property objects and persistent identities.
CreateProperty/Append/Item/Value/Delete use the existing type, ownership, lifetime,
read-only and generation checks. Rename/rebuild preserve the property. The field
designer edits the mask beside Caption and Description; an empty designer mask
removes this custom property. Legacy imported Format properties without a separate proof marker refuse3251;
they are not inferred absent from caption provenance. New Jet4 import support is
documented in SOURCE_FIELD_DISPLAY_PROPERTIES.md.

Native numeric/date/text table display applies nonempty masks through the shared
native formatter. EditRole and SQL values retain their existing unformatted
representation. Clipboard copy retains the checked raw typed text contract.
Native table print/PDF uses a separate checked formattedText adapter, so it emits
the same formatted values without parsing display strings or printing error text.
A formatter error is shown as #Fehler(code) in the datasheet and aborts output
snapshot preparation. Null remains blank and Boolean checkbox display is retained.
Unsupported masks are preserved as metadata with explicit rendering failure.

This scope covers owned native table fields and editable table-backed models.
Full Jet3/ACE field Format extraction, query/view format lineage, automatic
form/report-control inheritance, DecimalPlaces, complete string/scientific masks
and Access 2003 display/printing parity remain open. No original files are edited.

Existing DAO/window checks2/2 (8.00s); application built. Actual PDF text includes
the stored numeric mask; unsupported masks refuse before snapshot output.
