# Bounded native report aggregates

The Access-expression parser and native report host now handle Avg/Min/Max in
addition to Sum/Count. They operate on actual complete loaded SQL rows, for a
direct field or the existing row-scoped expression binding. NULL observations
are omitted; no observations produce NULL (Count produces zero). The existing
empty-compound binding guard remains 3251 until its field references can be
validated independently of a row. Missing/truncated sources and nested aggregates
retain their explicit guards. Group scope is not fabricated.

Avg accepts finite native Number values (Integer/Long/Single/Double), including
supported numeric IIf/arithmetic expressions, and uses an online long-double mean
returning Double. Min/Max accept homogeneous Number, Date or Currency observations
and retain the selected native value. Currency compares signed scaled integer
units directly; Date compares finite serial values. NULL observations are skipped.
Mixed kinds, text, Boolean and other types are explicitly refused. Actual SQL-row
CCur/CDate expressions are checked through control Value identities and PDF output.
Full Access result typing, exact Windows floating rounding, locale text ordering,
Currency/Date Avg and a Windows oracle remain open.

A shared scalar-only accumulator preserves existing Sum/Count behavior. Direct
Count does not coerce non-NULL field values merely to count them; compound errors
propagate. Sum keeps its existing numeric conversion/overflow rules. Render-only
512-entry bounded scalar caching and Requery/exit invalidation also apply to the new
functions; no widget, cursor or host is retained.

Existing report cases populate an actual nullable numeric table. Native controls
and PDF Avg/Min/Max plus compound Avg are compared with independent SQL, including
real insertion/Requery and all-NULL results. A text minimum refuses output and
preserves a previously generated numeric PDF. No new CTest target or catalog ID.
Original aggregate evidence is unchanged: three available Sum expressions have
SQL comparisons; the six original compound calls still lack prepared context.
No original Avg/Min/Max output oracle is claimed.

References: [Avg](https://support.microsoft.com/en-us/access/functions/avg-function),
[Min/Max](https://support.microsoft.com/en-us/access/functions/min-max-functions).
Statistical variance/deviation cardinality edge cases remain separate work.

## First/Last over the actual report source

First/Last now return the scalar value of the first/last loaded source record,
including a NULL endpoint. They do not search for the first/last non-NULL value.
Supported row expressions use the existing row scope; objects, arrays and binary
values are refused. Empty direct-field sources yield NULL; the existing empty
compound-binding guard remains. Unsupported groups and complete Jet/Access
ordering/type/oracle equivalence remain open. An explicit ordered record source
is used for deterministic evidence; arbitrary underlying scan order is not
presented as a matching Access insertion order.

Render caching now also permits immutable Empty/Boolean/Date/Currency and text
up to 4096 UTF-8 bytes. The 512-entry cap limits retained text to at most 2 MiB;
larger text is computed without caching. No object, array, DAO or widget reference
is retained. Requery and render exit still invalidate all entries.

Existing actual SQL/Qt/PDF cases compare ascending and descending source
endpoints, compound selectors, changed endpoints after insertion/Requery and
NULL endpoints through the actual control Value API. The existing SQL First/Last
implementation is separate; this extends report-expression execution.
Reference: [Access 2003 First/Last](https://documentation.help/MS-Office-Access-2003/dasqlfirstlast.htm).

## Exact Currency report Sum

Homogeneous Currency observations now accumulate signed scaled integer units and
return Currency. Overflow raises error 6 before addition; mixed Currency/non-
Currency observations remain explicitly unsupported until coercion is validated.
NULL observations are omitted; all-NULL input still yields NULL. Other existing
Sum conversion behavior remains unchanged. This report accumulator does not use
SQLite REAL as an intermediate; SQL Currency storage limits remain separate.

An existing report case reads exact decimal strings from actual SQL rows, uses
Sum(CCur([amount])) and checks the native control Value equals the maximum signed
Currency units. Adding 0.0001 through the database and Requery then raises error 6;
failed PDF output preserves the previous successful PDF. No Windows result oracle
or complete aggregate coercion/group equivalence is claimed.
