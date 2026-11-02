# Bounded Currency across SQL objects and query results

SQL Forms/Formulare/Reports/Berichte reference callbacks now pass native Currency
values through the checked lossless numeric REAL encoder, instead of rejecting
the scalar kind or binding decimal text. The same strict +/-2^39 bound applies.
This preserves numeric comparisons, accepted Value versus pending Text, and
range failure before SQL effects. The existing named-reference syntax and
optional .Value suffix are retained; SQL Screen syntax is not added.

The query compiler now records proven projection types for complete unqualified
CCur calls and direct declared Currency/Money parameters. Parentheses and unary
+/- preserve this bounded encoding. SELECT modifiers retain ordinal alignment;
filters copy the proven source descriptors. This is bounded token inspection,
not a general SQL type inferencer. UNION/INTERSECT/EXCEPT, wildcard expansion and
ambiguous output names do not acquire guessed Currency descriptors. Unknown
expressions keep existing generic SQLite metadata.

Proven output types travel with actual query results into native DAO snapshots
and bound read-only form/report controls. Field.Type returns dbCurrency (5),
including EOF; actual values use fixed Currency units and Null remains Null.
The query table's display role formats proven Currency from integer units;
its edit role retains the underlying numeric encoding. This avoids shortening
large amounts through generic Double-to-string formatting. Prepared output
metadata remains available for empty results without executing expressions twice.

New currency_sql_objects and currency_query_results regressions first reproduced
both missing boundaries (0/2). Native form/report/VBA/DAO/SQL tests verify named
and localized references, .Value, numeric comparison, pending/no-save, large/
unsupported unbound Currency, root conversion/parameter/Null, parentheses/signs,
TOP/DISTINCT, saved queries, EOF, filter/empty metadata, unknown/ambiguous
projection guards, query-table exact display and actual read-only Qt controls.
Converter/parameter/field/new tests pass 5/5 (0.25 s); extended report/saved-query/
TOP/empty tests pass 2/2 (0.12 s). GUI suite has 42 CTests; full validation is
recorded in PROGRESS.md.

General field/view/alias/aggregate/arithmetic type propagation, mixed-column SQL
typing, full-range/decimal storage, complete object-reference grammar, locale and
Windows parity remain open. Existing full Jet Currency SQL arithmetic is not
implemented by these descriptors. No complete original business workflow is added.
