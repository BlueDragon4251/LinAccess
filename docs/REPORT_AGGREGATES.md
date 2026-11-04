# Bounded native report aggregates

The Access-expression parser and native report host now handle Avg/Min/Max in
addition to Sum/Count. They operate on actual complete loaded SQL rows, for a
direct field or the existing row-scoped expression binding. NULL observations
are omitted; no observations produce NULL (Count produces zero). The existing
empty-compound binding guard remains 3251 until its field references can be
validated independently of a row. Missing/truncated sources and nested aggregates
retain their explicit guards. Group scope is not fabricated.

The new Avg/Min/Max subset accepts finite native Number values (Integer/Long/
Single/Double), including supported numeric IIf/arithmetic expressions. Text,
Date, Boolean, Currency and other value kinds remain explicitly unsupported for
these new functions rather than guessing comparison/coercion/precision semantics.
Avg uses an online long-double mean and returns Double; Min/Max retain the chosen
native numeric value. Full Access result typing, exact Windows floating rounding,
locale text ordering, Currency/Date aggregates and a Windows oracle remain open.

A shared scalar-only accumulator preserves existing Sum/Count behavior. Direct
Count does not coerce non-NULL field values merely to count them; compound errors
propagate. Sum keeps its existing numeric conversion/overflow rules. Render-only
512-entry scalar caching and Requery/exit invalidation also apply to the new
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
