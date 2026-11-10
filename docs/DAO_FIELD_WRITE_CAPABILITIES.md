# Native DAO field write capabilities

Recordset Field.DataUpdatable now returns an actual Boolean for every resolved
field, using the native cursor's write capability and actual connection readonly/
query_only state. Ordinary fields no longer fall through the counter-only guard.
Supported physical tables and checked editable query fields returnTrue; readonly
snapshots/queries/frozen cursors returnFalse. This is a capability, not a promise
that every current-row, constraint, pending-mode or source-generation write will
succeed. Existing measured AutoNumber capability remainsTrue on writable cursors;
its separate Edit assignment protection remains3164. TableDef/QueryDef catalog
field behavior and original/unproven Properties guards are not widened.

Recordset Field.Attributes now uses actual validated native schema flags, adding
dbUpdatableField32 only for writable native cursors. Native fixed/variable field
flags1/2 and AutoNumber17 therefore produce33/34/49; readonly physical cursors
retain1/2/17. Only fields carrying validated native_attributes are generalized.
Other contexts keep the pre-existing checked generator path or explicit3251;
there is no inferred attribute catalog for arbitrary original/provider fields.
Both properties retain arity450, read-only383 and closed/stale field guards.
Native Property objects delegate to these actual getters.

App and existing DAO target built. One existing dao_execution bundle passed1/1
in3.69s. Actual table/query-alias/counter/text/Single/BLOB fields, readonly query/
physical snapshots, calculated output, PRAGMA query_only transitions, real Property
values, getter/write/close guards and actual aliased-field VBA access are covered.
No new CTest/catalog IDs;456/386/70 unchanged; F054/F055/F121 stay broader open.
No Windows provider/error-order/complete DAO equivalence claim.

Public contracts, not a Windows execution oracle:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-dataupdatable-property-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/field-attributes-property-dao
