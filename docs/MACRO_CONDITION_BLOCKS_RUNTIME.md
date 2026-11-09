# Native contiguous macro condition blocks

A native action condition of `...` continues the immediately preceding conditional
block. The first expression is evaluated once per block and repetition. Its
scalar Boolean result controls every consecutive continuation action. A row
without a condition ends that block. A new expression starts a new block.
Orphan continuations, including continuations after an unconditional row or at
the start of another group, refuse3251 during complete definition preflight,
before any prefix action executes.

Each invocation and repetition has its own block result. A child macro's
conditions cannot replace its parent's result. False continuations are skipped
without expression reevaluation, while still counting against the existing
visited-action budget. StopMacro, StopAllMacros, cancellation, scoped errors,
frame leases, repeat conditions and depth/action budgets retain their existing
behavior. All groups continue to be checked before selecting one for execution.

The actual native macro editor stores and reloads the continuation text through
its existing Condition column. The existing GUI check passes (8.89s), including
real editor save, SQL-audited once-only public function evaluation in both hosts,
true/false repetitions, nested frame isolation, and rejected raw orphan/group
programs before their SQL prefix. Existing layout checks pass (11.36s). No new
CTest target or Windows Access execution oracle is added. Noncontiguous implicit
inheritance, full imported macro grammar and complete Access macro parity remain
open under F088.

Macro execution/frame helper definitions now compile once in the existing data
static library alongside the parser. UI dispatch remains a callback supplied by
the real GUI host; the data layer acquires no Qt Widgets dependency. This keeps
subsequent macro control-flow edits out of the large window translation units.
