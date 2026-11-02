# Currency/Single/Double source metadata

MDBTools' MySQL schema backend labels both MDB_MONEY and MDB_FLOAT as FLOAT.
Older versions also label MDB_DOUBLE as FLOAT. The Access backend distinguishes
Currency, Single and Double. Native import now reads
that existing tool's Access schema output in addition to SQLite schema and MySQL
key/index metadata, preserving CURRENCY/SINGLE/DOUBLE in source_type annotations and
_la_table_schema. No new tool/runtime dependency or external source code is added.

Known column names must match; every ambiguous MySQL FLOAT must have an Access
Currency/Single/Double label. Missing/conflicting mappings diagnose an error before a
schema is returned for import, rather than guessing Single. The bounded parser
preserves raw bracket-wrapped MDBTools names, Unicode, punctuation and type case;
it ignores unrelated type labels and rejects contradictory duplicate labels.

The original regression first failed with 0/43 Currency metadata labels. New
imports now identify all 43 actual original Currency columns and 57 Double columns across the 29 Kassa
tables. The independent CSV cell comparison, 127 Long-column/471904 non-NULL typed
value checks, Boolean/date/counter regressions and fixture hash checks remain
active. Synthetic contracts check Currency/Single parsing, misleading type
prefixes, unusual names, conflicts and persisted metadata. Targeted import/original
tests pass 2/2 (7.81 s). GUI suite remains 37 CTests; full validation is recorded
in PROGRESS.md.

This is source-type metadata recovery. SQLite REAL storage, existing floating
CSV conversion and generic numeric field adapters are deliberately not relabeled
as exact Currency/Single execution. Exact decimal import/storage, Currency SQL
operators/aggregation/sorting/coercion, Single round trips, numeric precision/
scale and complete Jet/Windows behavior remain open. Existing imported projects
are not migrated; lost precision cannot be recovered merely by changing a label.
No complete original business workflow or general money compatibility is added.

Reference inspected outside the repository:
https://github.com/mdbtools/mdbtools/blob/v1.0.1/src/libmdb/backend.c
