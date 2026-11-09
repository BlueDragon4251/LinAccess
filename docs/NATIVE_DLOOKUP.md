# Native field DLookup

Global DLookup and Application.DLookup accept positional/named Expr, Domain and
optional Criteria in a live project context. The current subset binds Expr to
one actual existing field (including bracketed names) and Domain to a physical
table or installed non-parameter saved query. Canonical domain/field lookup is
case-aware and rejects ambiguity. Calculated Expr, explicit NULL criteria and
unproven non-NULL result types remain explicit unsupported cases. Missing/Empty
optional criteria acts as omission; ordinary string predicates use the existing
bounded Access SQL compiler. Full expressions, parameter-domain/reference binding,
collations and exact Access error identity remain open.

The implementation prepares actual source schema, then executes a read-only TOP1
field selection with the supplied predicate. It returns NULL for no match or a
NULL matching field; an empty string remains String. It never turns a missing
domain, unsupported field or failed query into a fabricated empty lookup. Proven
result-position types feed the existing native converter, retaining supported
Date/Boolean/Currency rather than guessing from cell text. No arbitrary ordering
guarantee is asserted when criteria do not uniquely select a record.

Each call reads the real database. Unsaved form buffers are excluded; after an
actual record save the new value is visible. No persistent lookup cache or host/
cursor retention is introduced. Only SELECT/non-ranked field projections are
allowed; original MDB files are not edited by lookup.

The existing native form case executes actual VBA and a calculated textbox over
a real nullable table and saved-query alias. It covers named/global/Application
calls, apostrophe criteria, empty/NULL/no-match distinctions, Date/Boolean/Currency
identity/value, and a dirty label before/after actual Save. Missing domains and
computed Expr refuse explicitly. No additional CTest target or catalog ID.

Reference: [Access 2003 DLookup](https://documentation.help/MS-Office-Access-2003/file9o.htm)
describes table/non-parameter query domains, no-match NULL, arbitrary selection
without unique criteria, and exclusion of unsaved records. This is a bounded
field implementation, not complete expression/domain compatibility.

## Native field DMin/DMax/DAvg

Global and Application DMin/DMax/DAvg use the same field/domain/criteria binding
and named-argument rules. Schema preparation can expose proven result-position
types without stepping records or evaluating user expressions. Unsupported or
unproven aggregate types are refused before loading the field values.

DMin/DMax select homogeneous Number, Date or Currency values through the shared
native report accumulator. Currency uses scaled integer comparison; the selected
subtype survives. DAvg accepts Number fields and returns the existing online-mean
Double result. Currency/Date averages, text collation, Boolean coercion, computed
Expr and full Access result/error/precision equivalence remain open.

Actual matching field rows are read through the bounded Access query executor;
NULL values are omitted and no non-NULL observations yield NULL. The existing
100,000-row and execution-budget guards fail explicitly rather than returning a
partial aggregate. There is no persistent result cache. Original MDB files are
not modified. Existing actual VBA cases check positional/named/Application calls,
criteria, no-match/all-NULL results and Date/Currency identities and values.
No additional CTest target or catalog item is introduced.

References: Microsoft Access
[DMin/DMax](https://support.microsoft.com/en-us/access/functions/dmin-dmax-functions)
and [DAvg](https://support.microsoft.com/en-us/access/functions/davg-function).
No Windows/original-output oracle is claimed.

## Application DSum/DCount and SQL result types

Application.DSum/Application.DCount now use the same existing SQL-expression
execution path as global calls, including named Expr/Domain/Criteria arguments.
That shared path now passes proven per-position SQL result types to the scalar
converter. In particular, SQL Currency sums remain Currency rather than losing
the subtype at the host boundary. SQL Currency storage/precision guards remain.
Existing expression support and NULL-criteria behavior of Sum/Count are retained;
this does not extend the stricter field-only Min/Max/Avg rules to calculated Expr.
Full Access coercion/error/reference parity remains open.

Actual VBA checks global/Application Currency identity, a calculated CCur sum over
explicit non-NULL input rows, Count(*) versus Count(nullable field), named calls,
criteria and an all-NULL sum. CCur(NULL) correctly remains runtime error 94.

## Native field DFirst/DLast

Global and Application DFirst/DLast use the field/domain/criteria binder, named
arguments and proven native result converter. They select the first/last row of
the actual executed field query, including a NULL endpoint. DFirst uses bounded
TOP1; DLast reads the complete bounded field result and refuses a row/budget-limit
failure instead of selecting a partial endpoint. No matches yield NULL. Text,
Date, Boolean and Currency keep the same supported converter semantics as DLookup.
Calculated Expr and full Access binding/type/error equivalence remain open.

The existing VBA case uses an explicitly ordered saved query to check endpoints,
NULL last values, filtered endpoints, named/Application calls, empty selections
and Currency identity. No Jet/Access physical scan-order equivalence is claimed.
Without a defined source order, no stable arbitrary selection promise is made.
Reference: Microsoft Access
[DFirst/DLast](https://support.microsoft.com/en-us/access/functions/dfirst-dlast-functions).

## Native sample DVar/DStDev

Global/Application DVar/DStDev accept the field-only numeric domain/criteria
subset with positional/named calls. They use the same schema proof, complete
bounded query execution and native Welford sample accumulator as report Var/StDev.
NULL observations are skipped; fewer than two Number observations return NULL.
Actual VBA checks independent hand-calculated sample values, named Application
calls, one match and no matches. Currency/Date/text/Boolean inputs, calculated
Expr, DVarP/DStDevP and full Access error/rounding/reference parity remain open.

## Bounded calculated domain expressions

DLookup/DFirst/DLast/DMin/DMax/DAvg/DVar/DStDev now accept bounded scalar Expr
text in addition to existing field names. The real SQL compiler executes the
projection over the named domain and checked criteria. Numeric arithmetic,
concatenation, IIf and supported native scalar calls can be combined with fields.
Quoted strings containing words such as FROM remain literals. Escaped bracketed
identifiers still address fields; [qty]+[id] is parsed as an expression rather
than misread as one bracketed name.

A dedicated token guard rejects added clauses, subqueries, top-level comma
projections, unbalanced parentheses, aggregate/window calls and functions outside
the bounded native scalar allowlist. The domain remains an object name, never SQL.
Prepare inspects output shape without stepping value expressions. Selectors use
actual returned scalar Qt types when a calculated projection has no stronger
proven type; proven Currency and Date metadata are retained. Missing rows return
Null. First/last retain the domain query's existing order semantics; no arbitrary
order is invented. Numeric calculated extrema/mean/sample statistics validate the
actual scalar values, and nonnumeric values retain explicit unsupported errors.
Physical field paths retain their existing type checks.

This adds no arbitrary project/library function binder, nested query grammar,
text/locale extrema, or new mean/variance Currency/Decimal conversion. Existing
field-only unsupported cases and reference boundaries are not declared complete.
The existing native-domain fixture checks actual global/Application host execution
alongside independent SQL state; no new CTest or business record fixture is added.
F050/F068/F118/F164 remain open and the fixed catalog stays456/386/70.
