# Bounded native Currency query parameters

PARAMETERS accepts Currency and its Money synonym. DAO Parameter.Type returns
Integer dbCurrency (5); assigned Value/default access returns fixed Currency
(vbCurrency=6). Unassigned/assigned Null retains existing Null semantics, with
missing execution bindings diagnosed as 3061. Type remains read-only (383), and
SQL replacement/Close retains generation/lifetime error 3420.

Parameter values now retain canonical exact Currency text independently of the
SQL backing, including both signed64 extremes. Scalar object input is evaluated
once, then the parameter generation/lifetime is checked again before mutation.
A scalar callback that replaces SQL or closes the query cannot write through an
obsolete parameter. Invalid conversion retains the previous value. The SQL binder
still uses the existing lossless Currency REAL backing strictly inside +/-2^39:
unsupported execution refuses without discarding the exact parameter value.
The query binder sends an actual
numeric Double encoding, rather than decimal text: a Currency parameter 0.5
compares numerically below 1. Direct numeric/text bindings use checked Currency
conversion and the same range guard. Money/Currency use identical descriptors.

currency_query_parameters checks actual native DAO/SQLite binding, type/value/
default getters, missing/Null/scalar-once/error retention, readonly/lifetime,
four-decimal rounding and large tick recovery after INSERT into a verified
Currency field. Numeric SQL predicates independently check QVariant binding kind
and results. The native currency_field_documents VBA handler also executes a
Currency-parameter UPDATE and requeries real bound/calculated Qt controls.
Targeted parameter/field/DAO tests pass 3/3 (0.33 s). GUI suite now has 39 CTests;
full validation is recorded in PROGRESS.md.

This adds bounded typed parameter values and numeric binding. It does not add
full-range storage, exact Currency SQL operators/aggregates/formatting, complete
Jet coercion/type propagation or Windows parity. [Proven root query projections](CURRENCY_QUERY_TYPES_RUNTIME.md) subsequently
preserve Currency types; general query typing still uses existing generic SQLite
metadata. For exact formatted TEXT use explicit native
VBA formatting into Text parameters; automatic SQLite numeric-to-TEXT formatting
is not claimed lossless for large amounts. Non-NULL Currency LIKE conversion
explicitly diagnoses 3251. Currency width syntax is rejected by the bounded
compiler (3085); exact Windows syntax/error identity remains open. No complete
original business workflow is added.

References:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/sql-data-types
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/sql-reserved-words
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/datatypeenum-enumeration-dao

Exact-value follow-up: both signed64 Currency extrema can be assigned/read even
when the current SQL REAL binder cannot encode them. Unsupported Execute retains
the exact value and writes no rows. Scalar default callbacks are checked against
the original parameter generation after evaluation. Existing parameter/DAO2/2
passed (2.40s), app built. This does not establish full-range Currency SQL/storage.


Full-range projection follow-up: a table-free SELECT consisting entirely of bare
Currency/Money parameters now binds signed64 ticks rather than binary64 amounts.
Each output receives a checked, unique alias and the private CURRENCY_TICKS
result descriptor. DAO retains Currency Value, Type=5 and Size=8; query display
and clipboard emit exact logical amounts, including both Currency extrema.
NULL stays NULL. An alias equal to a parameter name remains an SQL identifier.
Modified plans and textual filter/DAO Find wrappers refuse with 3251 rather than
comparing raw ticks with logical currency amounts. Arithmetic, FROM/WHERE,
ambiguous duplicate labels and all other SQL shapes retain the existing REAL
conversion boundary. The private result descriptor does not enable creation of
Currency database fields, arbitrary SQL money operators or MDB write support.

Existing parameter/DAO checks passed together 2/2 (2.69s); application built.


## Exact parameter-only arithmetic

The table-free signed64-tick projection now also accepts +, -, *, unary signs and
parentheses over declared Currency/Money parameters. Every result must be a
proven Currency expression; computed expressions require explicit unique aliases.
Bare parameters retain their previous implicit names. The proof parser is bounded
to64 nested levels and4096 nodes. It never assigns this representation to raw
numeric literals, mixed types, source fields, predicates, subqueries or functions.
Those shapes retain their existing adapters and explicit range/unsupported guards.

Private integer-only SQL callbacks reuse native currencyBinary with checked wide
intermediates and ties-to-even multiplication. They return exact signed64 ticks;
there is no intermediate binary64 conversion. Null stays Null; Currency overflow
raises6. Native DAO Type5/Size8 and Currency Values survive snapshots and EOF;
query display and clipboard format logical amounts from the result descriptor.

Binding now regenerates the proven exact SQL and checks it against the immutable
compiled statement as well as descriptors. A different source expression with the
same result count cannot reuse another expression's proof. Textual filtering and
DAO Find on these private tick projections remain refused; raw ticks never become
logical amount predicates. Physical Currency storage remains the existing verified
REAL-backed interval, and action/table/MDB storage semantics are not widened here.

Validation: existing parameter/arithmetic/results/DAO bundle4/4 passed (2.84s),
application built. Checks include maximum addition/display/clipboard, both extrema
as inputs, exact multiplication, precedence/grouping/signs, positive/negative
half-tick rounding, Null descriptors, overflow, DAO/EOF type/size and source-proof
tampering. Existing physical REAL fields, legacy arithmetic and rejected mixed/
source/action/filter shapes remain in the same targets. No new CTest ID or new
original/Windows/full-suite/package proof is attributed to this arithmetic stage.
