# Native SQL string adapters

The checked native SQL adapter now executes CStr, Str/Str$, Left/Left$,
Right/Right$, Mid/Mid$ and Replace through the shared VBA evaluator. It does
not substitute SQLite character counts or its three-argument Replace contract.
The same scalar argument decoder preserves full text bytes including NUL and
rejects binary or inexact large integer input. Only fixed internal names and
argument variables enter the evaluator; user values are never expression text.

The token compiler carries proven per-argument Boolean, Byte/Integer/Long,
Single, Double, Date and bounded Currency kinds into a guarded native callback.
It retains optional arguments, UTF-16 positions, Str leading-sign space and
fraction notation, Null propagation versus String-result94, Replace start/count/
comparison modes and native argument/range errors. Existing CStr behavior and
metadata are preserved. Result metadata is native TEXT even for nullable calls.

Proven types now reach WHERE/HAVING expressions and GROUP BY/ORDER BY keys too,
within the existing resolver's context bounds; ASC/DESC and multiple keys retain
their order. numericOnly continues to refuse rewriting unproven Date/Currency
operators. Resolver restrictions on joins, subqueries, source lineage and
HAVING contexts are not removed. This is no general Jet/locale proof. Native
UTF-8 storage still cannot represent isolated UTF-16 surrogates; the existing
checked conversion remains. Non-ASCII comparison requiring a locale provider
refuses3251; no complete text/database collation is claimed.

App/DAO/Currency query targets built. The existing DAO bundle first passed in
2.78s for the shared string family. After extending the same work through scalar
clauses, one existing bundled run passed2/2 in3.02s: currency_query_results0.07s
and dao_execution2.95s. Actual DAO paths cover whole surrogate pairs, two-/three-
argument Mid, typed first and replacement arguments, optional Replace modes,
Null and String-result differences, Date position coercion, original CStr
checks, real Boolean sort/group keys, Date text filtering and four-place
Currency text near the supported strict2^39 bound. No new CTest/catalog IDs;
456/386/70. Latest checked DEB is still application62d00929 plus packaging
correction5ec76beb, before CStr/OverGroup/these adapters. No Windows/full-suite
claim.


The shared scalar binding also covers proven numeric/calendar/conversion/format
arguments as described in DATE_RUNTIME.md, TYPED_SQL_ARITHMETIC.md,
SQL_NUMERIC_CONVERSIONS.md and NATIVE_FORMAT_RUNTIME.md. The later proven-Date
comparison/IS-NULL binding is the explicit exception to the older predicate
Date-operator guard above; unrelated Date/Currency operators retain their bounds.

The character/base/repetition family in CHARACTER_FUNCTIONS_RUNTIME.md extends
this transport to Hex/Oct, Space/String, Asc/AscW/Chr/ChrW and the supported
Hex$/Oct$/Space$/String$/Chr$/ChrW$ SQL aliases. Integer versus Long subtype
matters for base conversion width; Dollar aliases preserve their own Null rules.
Unsupported ANSI codepages, split surrogates, output limits and unproven source
lineage are not bypassed. Generic TEXT results still do not imply a particular
Jet Text/Memo field declaration. There is no new independent Access oracle.
