# Boolean SQL parameters

Declared Access SQL parameter types BIT, BOOLEAN, LOGICAL, LOGICAL1 and YESNO
now bind Boolean values through the native SQL/DAO path. BIT/LOGICAL/LOGICAL1
canonicalize internally to BOOLEAN; original QueryDef SQL remains unchanged.

The alias list follows Microsoft's Access SQL data type reference:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/equivalent-ansi-sql-data-types
This concerns Access BIT, not the different ANSI SQL BIT representation.

Binding preserves SQL NULL. Case-insensitive True/False text, surrounding spaces
and finite numeric input map to native 0/-1. Invalid text, empty text and nonfinite
numbers reject with error 13 before SQL execution. This replaces Qt's permissive
QString-to-bool conversion, which could silently store invalid text as true.

Synthetic DAO tests execute real temporary parameterized UPDATE QueryDefs for
all five aliases, verify stored 0/-1/NULL values and unchanged records after
invalid input. Direct binding tests reject empty text, infinity and NaN.

This does not add Parameter.Type, complete DAO assignment-time conversion,
Variant Empty identity, localized Boolean spellings, Boolean DDL aliases or
complete Access/Jet parameter/error timing compatibility. Full validation is
recorded in PROGRESS.md.
