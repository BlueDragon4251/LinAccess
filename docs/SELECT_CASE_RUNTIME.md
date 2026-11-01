# Native Select Case runtime

## Implemented scope

The native C++ interpreter executes multiline Select Case / Case / Case Else /
End Select blocks. The selector and its scalar object default are evaluated once
per entry. Case lists stop at their first match; later branches and inactive
outer branches do not evaluate their expressions. Inclusive To ranges evaluate
the upper bound only after the lower comparison succeeds. Relational clauses
accept =, <>, <, <=, > and >= with optional Is. Case Else must be last.

Nested selections share the conditional-frame stack with If blocks. Existing
For/For Each Exit For restores the enclosing stack depths; outward GoTo and
error-handler jumps discard exited selection frames. Resume restores a failed
selector/case frame and retries it. Under On Error Resume Next, a failed selector
suppresses that selection's branches; a failed Case skips that branch and allows
subsequent cases. These error-boundary behaviors are regression contracts, not
claims of an independently observed Access error oracle.

Case comparisons reuse expression-operator semantics, including Variant Null
(no match), Empty, numeric/text Variant ordering, scalar defaults and the module's
Option Compare. Text/Database comparison retains the existing host limitations:
Qt comparison in the GUI and bounded ASCII folding without Qt. This is not full
Jet collation, locale-sensitive coercion or typed Date/Decimal parity.

Whole-procedure validation rejects missing/misnested terminators, Case outside
Select, statements before the first Case, clauses after Case Else, invalid
relations and incomplete expressions before any procedure effects. Commas in
quoted strings, calls and bracketed identifiers remain inside their expressions.
Whitespace and existing continuation/comment rules are retained.

## Evidence

Core tests exercise numeric/list/range/string/Null/Empty cases, first-match
priority, all six relations with/without Is, Option Compare Text/Binary, quoted
commas/To, bracketed commas, one-time object default resolution, nested If/Select/
For and Exit For, skipped side effects, failed-selector/case Resume and Resume
Next, and invalid-block preflight. The Qt transaction-button regression now uses
real Select Case branches, including case-insensitive tag matching: actual SQL
commit, rollback, handled duplicate-key rollback and displayed balances are
asserted. The three targeted core/layout/original-heavy tests pass (37.97 s).

All 32 original VBA files contain zero Select Case starts. No additional original
handler execution is claimed. The language reference is
[Select Case](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/select-case-statement);
a Windows Access 2003 runtime differential remains unavailable.

## Remaining scope

This extends the existing statement interpreter; it does not supply a complete
VBA lexer, AST, type system or grammar. General colon-separated statements (for
example Case 1: action), arbitrary block-internal labels, full error/focus
semantics, remaining Do/Exit variants, arrays, classes and optional parameters
remain open. Existing unsupported typed module fields still fail explicitly;
Select Case does not silently make them executable.

Validation: host 21/21 (86.07 s), Debian 12 21/21 (65.05 s), Ubuntu 24.04
21/21 (59.69 s), freshly rebuilt GUI-disabled core 1/1 (0.90 s). Fresh host/
distro DEBs pass exact payload/backend/permissions/fixture-exclusion checks;
isolated installed Debian/Ubuntu packages pass backend/counter execution and
GUI startup. All four original fixture hashes pass.
