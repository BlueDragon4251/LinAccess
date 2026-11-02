# Native classic QBColor palette

QBColor(index) now runs as a native VBA intrinsic and Access SQL/DAO adapter.
Both use the same fixed sixteen-entry BGR palette, including dark blue/green,
classic white (192/192/192), gray and bright colors. These are fixed colors, not
OS system-palette references. Microsoft documents the 0..15 names; the public
Microsoft Visual Basic reference implementation establishes the numeric color
table. Reading that source introduces no .NET application dependency and is not
an Access 2003 runtime oracle for coercion or error order.

The native primitive resolves one scalar default, uses the existing integral
rounding/conversion and requires index 0..15. Empty becomes zero; Null is 94,
invalid scalar types 13, out-of-range indices 5, conversion overflow 6 and wrong
arity 450. Complete typed-Integer/locale/error precedence remains open. Results
fit exactly in the existing numeric Variant representation; full Variant subtype
identity is not claimed.

The SQL compiler preserves quoted QBColor fields, supports mixed-case calls,
parameters, alias reuse and compiled pure source views. SQL NULL is Null rather
than Empty, binary arguments reject 13, and numeric errors use the existing
SQL error protocol. Palette expressions can drive actual DAO joined updates;
an invalid source index rolls back the statement. Existing joined-target and
ambiguity restrictions remain. This does not add every VBA intrinsic to SQL.

## Evidence

Core regressions cover all sixteen documented table entries, nearest-even index
rounding, Empty, numeric strings, Null/type/range/arity and one-time scalar access.
Actual Qt/VBA ForeColor assignments and Application.Eval change a real masked
editor to classic bright red/blue while preserving pending input, font, focus/
selection, SQL values and data permissions through the existing color regression.
SQL tests read all sixteen indices, preserve a same-named quoted field, reuse
aliases, bind parameters and reset failed results. Native DAO joined updates
through a compiled palette view persist colors and roll back invalid indices.
The targeted core/DAO/query-model/join/GUI set passes 5/5 (9.75 s).

Original lexical inventories contain zero QBColor calls in 100 query artifacts
and 32 VBA modules. No additional original query/business execution follows.
Final full/distribution/package validation is recorded in PROGRESS.md.

## Public references

- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/qbcolor-function
- https://github.com/dotnet/runtime/blob/main/src/libraries/Microsoft.VisualBasic.Core/src/Microsoft/VisualBasic/Information.vb

Full Access color themes, every OLE system color, disabled/selected rendering
and Windows visual equivalence remain open.
