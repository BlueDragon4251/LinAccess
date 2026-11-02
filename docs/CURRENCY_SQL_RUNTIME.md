# Bounded native SQL CCur

Unqualified Access-SQL CCur calls now rewrite to a native SQLite callback that
uses the same checked C++ Currency conversion and lossless REAL encoding as
native Currency fields/parameters. Plain decimal text is converted before any
binary64 step, preserving four-decimal ties-to-even rounding. Numeric input uses
the existing native CCur conversion. The supported output interval is strictly
inside +/-2^39; valid full-range Currency outside that backing diagnoses 3251.

The callback diagnoses invalid text/BLOB (13), Null (94), wrong arity (450) and
Currency conversion overflow (6). Integer SQL input beyond exact Double range
is explicitly unsupported (3251). Function errors use the existing typed native
SQL error channel. Exact Windows/Jet Null/error/coercion identities remain open.
Quoted SQL literals are preserved; same-projection aliases use the existing
bounded pure-expression expansion with the new converter registered.

currency_sql_documents checks actual SQL values, signed booleans, positive/
negative decimal ties, numeric expression coercion, large tick recovery,
quoted-literal preservation, aliases and typed errors. A real DAO action query
fails on the second row after a valid first-row conversion; raw SQL checks prove
all original rows/NULL remain unchanged. A native VBA QueryDef UPDATE with Text
input then requeries actual bound Qt controls, including the largest supported
tick, and verifies rejection before unwanted SQL changes. Converter/parameter/
field tests pass 3/3 (0.16 s). GUI suite now has 40 CTests; full validation is
recorded in PROGRESS.md.

This is an explicit bounded SQL conversion. The subsequent [query type adapter](CURRENCY_QUERY_TYPES_RUNTIME.md) preserves
proven root CCur/parameter result types. Other query metadata still uses generic
SQLite types; full Currency result-type propagation,
full-range storage, regional Currency strings/formatting, exact SQL Currency
operators/aggregates and Windows parity remain open. Unwrapped raw arithmetic
is not made Currency-aware. No complete original business workflow is added.

References:
https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/sql-data-types
