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
