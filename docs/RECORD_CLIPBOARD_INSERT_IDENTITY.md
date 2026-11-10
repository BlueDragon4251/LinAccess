# Native record clipboard physical insertion identity

Whole-record PasteAppend now checks the existing form main/TEMP source boundary,
reads optional source-counter metadata from main and inserts explicitly into the
main table. INSERT and DEFAULT VALUES consume exactly one non-NULL RETURNING
rowid before counter completion, retaining the valid key0. The existing savepoint,
filter-exclusion rollback, typed MIME, source restrictions and unsupported insert-
event guard remain. No broad external clipboard/DataEntry/event parity is added.

The shared existing access_2003_layouts check passed13.78s. Actual VBA appends a
typed copied record with manual key0, selects that row, checks the real count and
physical SQL identity; main/TEMP shadow refusal leaves both sources unchanged.
Existing managed/default-only counter, duplicate key and excluded-filter rollback
cases also pass. This check was not repeated for independent DAO fixture fixes.
No new CTest/catalog IDs;456/386/70 unchanged; F168/F340 retain broader open scopes.
Latest checked DEB/Original is stillc67cbd9e before later query/form changes.
