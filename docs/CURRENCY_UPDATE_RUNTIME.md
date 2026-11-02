# Currency in simple UPDATE statements

Simple physical-table UPDATE statements now inspect verified live Currency
metadata before execution. Currency assignment targets convert their scalar
value to four-decimal fixed-point units, then use the bounded numeric backing
encoding documented in CURRENCY_FIELDS_RUNTIME.md. NULL remains NULL. Text
conversion uses the shared checked CCur conversion, including nearest-even ties;
invalid inputs report 13 and unsupported backing values report 3251.

The token-based path supports optional table aliases, qualified assignment
names, multiple assignments, WHERE and declared parameters. Original parameter
order/bindings are preserved. Proven Currency +, - and * on assignment RHSs use
the existing fixed-point expression compiler. SQLite evaluates simultaneous
assignments against the old row. Conversion errors abort the whole statement,
including earlier rows; no partial Currency update is accepted.

`currency_sql_updates` exercises real DAO QueryDef.Execute, text/Currency
parameters, alias/qualified names, simultaneous assignments, mixed ordinary
columns, NULL, rounding, range errors and rollback. A native Qt form executes
VBA that updates Currency via a text parameter and requeries; its exact value
and visible boundary text are checked. Unmarked ordinary REAL updates retain
their existing numeric behavior. Tests use synthetic in-memory copies.

This is bounded single-table support. Joined actions retain their separate
existing implementation; views, nested SELECTs, INSERT/append/make-table money
coercion, general expression typing and complete Jet/Windows/locale semantics
remain open. No full Access 2003 parity claim follows from these tests.
