# Imported AutoNumber metadata and single-primary-key generation

MDBTools' MySQL schema exposes `auto_increment` modifiers that its SQLite schema
does not consistently preserve. AccessCatalog now carries these modifiers through
encoded structural comments; the restricted schema parser preserves a separate
AutoNumber flag. Quoted default text containing the keyword does not mark a
counter, and unknown/noninteger counter targets are rejected.

All imported column metadata in `_la_table_schema` records `auto_number` and
`auto_number_native`. An INTEGER AutoNumber that is the single primary key uses
native SQLite `INTEGER PRIMARY KEY AUTOINCREMENT`, retaining explicit imported
values, advancing after higher explicit inserts, and avoiding reuse of a deleted
highest committed key. Ordinary composite keys and noncounter fields retain their
existing definitions. No source DDL is executed directly.
These rules apply to newly imported tables; previously imported projects are not
silently rebuilt or assigned guessed counter flags/seeds.

The original Kassa schema contains 21 AutoNumber modifiers. All 21 flags are
preserved; 16 are single primary keys and have the native generation above. Five
have no primary key: their flags are retained with `auto_number_native=false`,
but automatic generation is still unsupported. Their imported values/constraints
remain as before. This is not complete support for all 21 source counters.

Targeted tests verify synthetic explicit/generated/no-reuse/higher-key inserts,
metadata flags, malformed targets and keyword quoting. The real 29-table copy
checks all 21 flags/16 native counters and performs one savepoint-isolated
insert/delete/insert probe against an eligible copied original table. Original
rows remain unchanged and continue to be compared cell-by-cell; the source MDB
is read-only and hash-checked. A counter without a primary key explicitly tests
metadata-only status and failure to generate a required missing value.

Boundaries remain: non-primary-key/composite counters, Random/Replication IDs,
the source database's stored next-counter seed when higher than remaining data,
exact Jet signed-Long overflow, gaps after failed/rolled-back allocations,
assignment timing before form save, DAO Attributes and an Access 2003 runtime
oracle. SQLite AUTOINCREMENT behavior is not a blanket Jet AutoNumber claim.
