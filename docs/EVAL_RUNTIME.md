# Bounded native Access Eval

`Eval(expression)` and `Application.Eval(expression)` evaluate text through the
native expression runtime. Statement forms `Eval expression`, `Eval(expression)`
and `Call Eval(expression)` discard the result. Evaluation uses database comparison
mode, supports single- and double-quoted strings, escaped quotes, arithmetic,
supported built-ins, Forms references and Access `Is Null` / `Is Not Null` tests.
It does not inherit the calling procedure's local variables.

Public Functions are resolved through the host's module lookup and execute actual
VBA/DAO effects. Private Functions, Subs and missing procedures fail with 2425
before their bodies run. Required arguments report 450; Null input reports 94;
recursive Eval is bounded with error 28 and resets after failure. Unsupported
array results fail with 13. Native form shutdown defers physical deletion until
the complete evaluation returns, including a subsequent modal InputBox.

Regression coverage includes actual native Forms/Application objects, saved SQL
results, all three statement forms, restricted function access, recursion recovery
and deferred deletion while a modal dialog processes events. An unchanged original
module identified by SHA-256
`aaea2a6871d7e2249e0e213b6acd74ac5a22ec4e12608be82a053368d16355b1`
executes through native Application.Eval on prepared synthetic controls in 20
states: non-Null numeric values 0, 1 and -1, all Null, each of the 15 controls
individually Null and an alternating mixed state. A separate counted run checks
15 expression references plus one extra original Forms lookup per Null branch.
Core regressions and actual native controls independently check true/false Null
results. The complete operational context and an independent Access result oracle
are not established by these tests. No original rows, authentication
values or VBA source are reproduced here.

This is not full Access 2003 Eval compatibility. Complete expression grammar,
collation/locales, exact function visibility/name resolution, declared result
coercion, arrays, named arguments and a Windows Access reference comparison remain
open. The original project's complete sale workflow remains unverified.
