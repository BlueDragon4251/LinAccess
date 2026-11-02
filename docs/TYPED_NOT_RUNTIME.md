# Typed native VBA Not

Reproduced two independent failures before the fix: the core type assertion and
native bound/calculated Byte value assertion both fail against the previous
Double-only Not implementation. Native Not now evaluates scalar default values
once, preserves Boolean negation/Null and inverts typed Byte/Integer/Long at
8/16/32 bits. Byte 4 therefore becomes Byte 251 rather than Double -5. Empty
becomes Integer -1; supported floating/Currency numeric inputs use checked rounded
Long conversion and produce Long. Existing comparison/power precedence is
preserved. Numeric range failures raise 6; Error/type mismatch 13 and Nothing 91.

Core regression verifies boundary bit patterns, subtype/name, double negation,
scalar-once, Boolean/Null/Empty, checked coercion and precedence. The actual
byte_value_documents Qt/VBA/DAO/SQL test now calculates the inverted bound value,
returns it through a declared Byte variable and writes integer 251 to a constrained
SQL table. Failed overflow does not execute subsequent SQL; the source record
still saves only when explicitly requested. Both targeted tests pass after the
fix: 2/2 (1.08 s), compared with 0/2 before it.

This is not complete numeric/string/locale/operator or Jet SQL Not parity.
Boolean-looking textual operands, all mixed promotion rules and a Windows Access
oracle remain open. SQL rewriting is unchanged. No extra original business
workflow is claimed. Full validation is recorded in PROGRESS.md.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/not-operator
- https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-varnot

The Wine automation implementation was inspected outside this repository to
cross-check typed inversion behavior. No Wine source/dependency is included.
