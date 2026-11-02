# Native VBA Like operator

VBA expressions and If/loop conditions support Like at comparison precedence.
The native matcher handles *, ?, # (ASCII digits), character lists, negated
lists and ranges, literal brackets and empty [] groups. Matching consumes UTF-16
code units. Binary matching is ordinal; text comparison uses the active host
and the current module's Option Compare. Null operands return Null. Invalid or
descending ranges fail with VBA error 93 before matching.

The implementation uses dynamic programming with two rows rather than recursive
wildcard backtracking. Existing core and DAO targets cover actual expressions,
Option Compare Text, precedence and data-changing VBA conditions.

This is separate from the existing Jet SQL LIKE translation. Full Windows
locale/collation, ligature/expansion behavior and BSTR identity remain open.
Non-ASCII text-collation ranges and locale-aware comparisons of isolated UTF-16
surrogate units fail explicitly with 3251. No Windows Access oracle was available.
