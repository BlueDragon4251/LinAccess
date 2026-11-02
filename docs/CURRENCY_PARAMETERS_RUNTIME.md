# Bounded native Currency query parameters

PARAMETERS accepts Currency and its Money synonym. DAO Parameter.Type returns
Integer dbCurrency (5); assigned Value/default access returns fixed Currency
(vbCurrency=6). Unassigned/assigned Null retains existing Null semantics, with
missing execution bindings diagnosed as 3061. Type remains read-only (383), and
SQL replacement/Close retains generation/lifetime error 3420.

Assignments use the existing lossless Currency REAL backing strictly inside
+/-2^39. Scalar object input is evaluated once. Conversion/range failure occurs
before replacing the old parameter value. The query binder sends an actual
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
