# Exact arithmetic for proven Currency query expressions

Bounded SELECT projections now parse proven Currency operands into native fixed-
point addition, subtraction and multiplication calls. CCur outputs, declared
Currency/Money parameters and verified single-table Currency columns can supply
operands. The compiler preserves arithmetic precedence, left associativity,
parentheses and unary signs. Result descriptors retain Currency through native
DAO, query-table and read-only Qt controls.

The callbacks decode checked canonical Currency values and use the existing C++
int64/checked-wide-intermediate arithmetic. Multiplication rounds four decimals
ties-to-even. Null produces Null. Full Currency arithmetic overflow diagnoses 6;
each SQL node must also fit the supported lossless REAL interval and diagnoses
3251 outside it, including unsupported larger intermediate results. Proven sums
can consume the rewritten Currency expressions using the exact accumulator.

currency_sql_arithmetic first reproduced generic Double evaluation (0/1).
Arithmetic/sum/query/schema/object tests pass 5/5 (0.28 s). Independent expected
units cover 0.1+0.2, negative differences, product rounding, associativity,
precedence/parentheses/unary signs, nullable fields/parameters, Sum(a+b), overflow/
backing guards and an actual read-only Qt report displaying 0.3. Source rows stay
unchanged. GUI suite now has 45 CTests; full validation is recorded in PROGRESS.md.

This is a bounded projection expression path, not a complete typed SQL AST or
full Jet arithmetic. Mixed Currency/other-number rules, division, complex
function/condition/alias/view/join/set-operation typing, action target coercion,
full-range SQL output, locale and Windows parity remain open. Unproven expressions
keep existing generic behavior; SQL Currency arithmetic is not claimed generally
compatible. No complete original business workflow is added.

Table-free expressions entirely over Currency/Money parameters now have a
separate full signed64 tick path for +, -, *, signs and parentheses; see
CURRENCY_PARAMETERS_RUNTIME.md. The physical-field/CCur/REAL boundaries above
remain in force, including the exact backing interval and unsupported mixed types.
