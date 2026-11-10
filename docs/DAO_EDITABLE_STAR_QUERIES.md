# Native DAO editable star-query dynasets

Mode2 SELECT cursors can edit a single keyed physical main table when the existing
editable-query validator proves a full `*` or `t.*` projection, a supported WHERE
predicate and physical ORDER BY fields. No parameter bindings, joins, aggregates,
computed/projection aliases, keyless targets or TEMP source shadows are promoted.
Mode4 remains a read-only snapshot. Physical schema types, origins, clone state,
bookmarks, transactions and checked main-table mutation paths are reused.

Reload applies the validated predicate and initial ordering. UPDATE and DELETE
also include that predicate beside the physical row identity: a cached row which
no longer satisfies it cannot silently be written. Changing a predicate field can
remove the row on refresh; native3021/no-current and unavailable LastModified are
bounded refresh behavior, not a claim of full Windows dynaset identity semantics.
Filtered AddNew is explicitly3251 before buffers/defaults/counter reservations.
Unfiltered star cursors support actual AddNew with native family append order.
Filter/Sort-derived explicit snapshots and actual QueryDef provenance remain
available. Explicit field lists remain read-only in DAO in this scope.

App/DAO/query-model/Currency targets built; one existing five-check bundle passed
5/5 in3.88s (DAO3.34s). Actual SQL and VBA cover filtered writes, shared clones,
row removal from membership, stale-predicate refusal with retained edit buffer,
successful retry, deletion, sorted insertion, temporary QueryDef writes, read-only
projection/snapshot/keyless boundaries and independent physical SQL verification.
No new CTest/catalog IDs; F054/F055 retain their broader open requirements and
456/386/70 is unchanged. No Access2003 SQL/DAO result oracle was supplied.

The preceding published source c67cbd9e has a checked DEB and successful existing
Original DAO event check1/1 in120.41s; seven protected hashes are unchanged.
That package predates this query-writing scope. Latest actual query inventory is
still source4ebc3f6c:100 queries,78 native executions,22 bounded failures, without
an Access result oracle. No fresh GUI/inventory/package proof for this scope.
