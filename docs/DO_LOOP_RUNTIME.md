# Native Do/While loop runtime

## Implemented scope

The native C++ statement interpreter executes Do While / Do Until ... Loop,
Do ... Loop While / Loop Until, unconditional Do ... Loop, and While ... Wend.
Entry conditions can skip the body; exit conditions execute after the first body.
A loop cannot have both entry and exit conditions. Until reverses the truth test;
a Null condition behaves as false, so While Null stops and Until Null continues.
Conditions use the existing expression engine and are evaluated at their actual
entry/exit point, not cached or evaluated in an inactive parent branch.

Exit Do binds to the nearest enclosing Do, including an inline If branch. A
While/Wend does not become an Exit Do target. The compiler records the exited
conditional, With and For/For Each depths; execution removes precisely those
frames before continuing after the target Loop. This preserves outer With/For
contexts and exits inner Select/If/With/For Each contexts. Existing Exit For can
leave a For nested within a Do without leaving the Do itself. Outward labels,
Resume and error handlers retain their existing scope and restrictions.

Whole-procedure preflight rejects Exit Do outside Do, mismatched Loop/Wend,
misnested blocks, missing ends, malformed conditions and simultaneous entry/exit
conditions before procedure effects. The existing 10,000-step execution budget
also stops unconditional loops; this is bounded execution, not an unbounded VBA
runtime. A small Select Case tokenizer correction accepts Is adjacent to its
operator (Is>=2), using keyword boundaries so identifiers remain intact.

## Evidence

Core tests exercise all entry/exit variants, zero-versus-one iteration, While/Wend,
Null conditions, inline Exit Do, innermost binding, Exit For within Do and Exit Do
through While/Wend. A nested For/With/Do/For Each/With/Select/If test verifies both
the outer object context and the outer loop survive an exit. Body Resume Next
and exit-condition Resume retry are exercised; the latter must not execute the
body again. Invalid syntax is rejected before a preceding field assignment and
a pure unconditional loop reaches the step budget.

The native Qt transaction-button test now executes both actual account SQL
updates through Do Until and Select Case. Commit, rollback and a handled
post-loop duplicate-key failure retain independently queried balances and real
requeried controls. Core/layout/original-heavy targeted tests pass (35.98 s).
The 32 original modules contain seven existing Do While starts and no Do Until,
bare Do, postcondition Loop, While/Wend or line-start Exit Do. No new original
business-handler execution is attributed to the added variants.

Language references:
[Do...Loop](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doloop-statement),
[While...Wend](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/whilewend-statement).
No Windows Access 2003 runtime differential was available. Arbitrary condition-error
recovery and the complete Access compiler/error ordering are not established.
General colon syntax, block-internal labels, full grammar/type fidelity and
unimplemented language features remain open; this does not replace the existing
statement interpreter with a complete lexer/AST.

Validation: host 21/21 (63.64 s), Debian 12 21/21 (66.30 s), Ubuntu 24.04
21/21 (60.17 s), rebuilt GUI-disabled core 1/1 (0.98 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass.
