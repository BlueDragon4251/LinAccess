# Native SQL exponentiation

Access SQL ^ now emits a real native scalar call that uses the existing VBA
power evaluator. Atomic operands are bound structurally through names, qualified
fields, parameters, literals, complete parentheses and function calls. Rewriting
is left to right, preserves unary minus outside the base, consumes signed
exponents and handles nested explicit grouping. String contents and quoted
identifiers containing ^ remain intact. Arithmetic and WHERE predicates use
actual SQL values, not precomputed fixture results.

MS-VBAL5.6.9.1 defines exponentiation above unary negation and all binary
operators as left associative. Thus -2^2 is -4, (-2)^2 is4, 2^3^2 is64 and
2^(3^2) is512. Bound arguments are evaluated once by SQL. Root power expressions
have proven Double metadata through DAO; Null results remain Null. Generated
helper names do not replace the original displayed expression heading.

Existing native evaluator boundaries remain: zero with a negative exponent and
underflow explicitly reject3251; invalid negative-base fractional exponents
reject5, overflow/nonfinite numeric input6 and unsupported scalar input13.
Missing/unsupported operand syntax rejects3075; at most128 power rewrites are
accepted. Full Date/object/Decimal transport, SQL grammar, locale/Windows
floating-point equivalence and complete typed expressions remain open.

The existing access_query_model and dao_execution cases jointly check chains,
signs, grouping, compound precedence, nested functions, field/parameter/WHERE
execution, quoted/literal carets, unchanged headings, actual DAO Double/Null
values and precise failures. No new CTest or catalog IDs. F068/F071 remain open.
