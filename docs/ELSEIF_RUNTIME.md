# Native multiline ElseIf chains

The native C++ statement interpreter executes multiline ElseIf clauses in block
If statements. Conditions are evaluated in source order until the first true
condition. Later conditions/bodies do not execute. If all are false, the optional
Else executes; Null conditions count as false. An inactive outer branch does
not evaluate inner If/ElseIf expressions. Each chain retains one conditional
frame, so nested Select/Do/For exits keep the correct enclosing stack depths.

Whole-procedure preflight rejects ElseIf without an enclosing If, after Else,
inside an unrelated block or with malformed expressions before any procedure
field/SQL effects. Error-handler Resume can retry a failed ElseIf condition;
Resume Next after an error within a selected body keeps that selected branch.

Core tests count condition-function side effects, check first/second/third/Else
selection, Null, skipped division errors, nested If/Select/Do/For with Exit Do and
Exit For, body/condition error recovery, no-match continuation and invalid-block
preflight. The actual Qt transaction-button regression now displays a separate
branch value for normal, rollback and duplicate-key modes while executing real
Do Until/Select Case SQL commit/rollback. Both branch selection and independent
account/control balances are asserted, including Option Compare Text.
Core/layout/original-heavy targeted tests pass (36.71 s).

The 32 original modules contain zero line-start ElseIf clauses. No additional
original business-handler execution is claimed. The contract reference is
[If...Then...Else](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ifthenelse-statement);
no Windows Access 2003 runtime differential was available.

This is the multiline subset of the existing statement interpreter. TypeOf,
full declaration/type binding, full inline/colon/label grammar and arbitrary
condition-error ordering remain open; this does not supply a complete lexer/AST
or establish full VBA conditional compatibility.

Validation: host 21/21 (65.43 s), Debian 12 21/21 (67.65 s), Ubuntu 24.04
21/21 (60.80 s), rebuilt GUI-disabled core 1/1 (1.29 s). Fresh host/distro
DEBs pass exact payload/backend/permissions/fixture-exclusion checks; isolated
installed Debian/Ubuntu packages pass backend/counter execution and GUI startup.
All four original fixture hashes pass.
