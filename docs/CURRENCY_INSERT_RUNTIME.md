# Currency in simple INSERT VALUES statements

A single-row INSERT INTO a physical table now resolves verified live Currency
metadata and converts each included Currency target before writing. The path
supports explicit column names in any order, or the table's ordinary column
order. Omitted columns retain SQLite/default behavior. Scalar NULL remains NULL;
text conversion uses shared checked CCur rounding and bounded numeric encoding.
Proven Currency parameter/CCur arithmetic in VALUES reuses the exact fixed-point
expression compiler. Declared parameter binding order is preserved.

`currency_sql_inserts` uses actual DAO QueryDef.Execute, imported synthetic
Currency metadata and a native Qt form running VBA. Tests cover four-decimal
text ties, NULL/omitted columns, reordered columns, positional values, Currency
arithmetic, ordered text/Long parameters and the largest supported stored value.
Invalid text/range in the second money column inserts no row. Malformed VALUES
syntax is preserved as an error. Ordinary unmarked REAL remains ordinary REAL.
The form requeries and navigates to the inserted record, checks VarType Currency,
its exact units and visible text. All data is synthetic and in-memory.

INSERT SELECT/append queries, multiple-row VALUES extensions, views/generated
columns, comprehensive defaults/constraints and full Jet/Windows/locale behavior
are not covered by this bounded path. The backing range from
CURRENCY_FIELDS_RUNTIME.md still applies. Full Access parity remains open.
