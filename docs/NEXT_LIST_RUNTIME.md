# Native multi-counter Next statements

Next inner, outer validates control names in lexical inner-to-outer order and
lowers them to the existing native Next steps with one original source line and
a shared source-statement boundary. Each loop retains its own matching Next
partner. Numeric and For Each counters therefore use their actual existing
advance/termination paths; zero-iteration inner loops and Exit For reach the
outer continuation without treating every counter as the innermost loop.

On Error Resume Next and an explicit handler's Resume Next skip the remaining
steps of the same source Next list after failure. Group loop state is released
when abandoning those terminators. Error snapshots are released after every explicit Resume, Resume Next or
Resume label. Restored live scopes retain their actual ownership, while stale
snapshots no longer retain collection enumeration locks or DAO clones after
the live scope ends. Single Next failures also release abandoned iteration state. Resume targets retain the existing bounded runtime
rules; full Access error/retry/alias precedence is not claimed.

Inline If wrapper merging relocates both loop partners and shared statement
boundaries. Native procedure-plan cache entries contain these source-only integer
metadata fields; storage accounting already includes Statement size. Single Next instructions use the same source-boundary cleanup metadata. GoTo now supports labels in the same active block or an ancestor, with checked
scope unwinding. Jumps into inactive/sibling blocks remain refused before body
effects. GoSub/error/Resume targets must still be top-level. See DAO_RUNTIME.md.

Existing core execution covers ordered counters/final values, empty inner/outer
loops, Exit For, empty/nonempty Collection/array loops and released enumeration
leases, Do scope, inline If and overflow handling, implicit Resume Next, explicit
Resume Next/Resume label with real Collection mutation afterward, wrong control
order refusal before body effects and retained blocked-label diagnostics. No new
CTest target or original data/source is used. General compiler/classes/control
flow and an Access 2003 execution oracle remain open.


## Native multi-counter Next source groups (2026-11-04)

Next lists lower to matched native loop steps and one source-line/error group.
Zero/Exit For, mixed For Each, Do and inline merging retain actual scope flow;
Resume Next skips remaining terminators and abandoned iteration state is
released. NEXT_LIST_RUNTIME.md records compiler/jump/error/oracle boundaries.
Existing core cases cover values, counters, empty loops, mutation after error
resumption and retained invalid-order/blocked-label diagnostics. Full desktop
48/48 (78.61 s), freshly rebuilt GUI-free 1/1 (1.39 s), targeted core 1/1
(1.31 s), Debian payload/permissions/Build-ID/pinned SQLite and all four original
fixture hashes pass. Scoped checklist 384 checked / 72 open; no extra CTest
target or complete Access 2003 parity claim.

## Error-resume resource ownership

Existing core cases exercise a retried loop-body conversion, mutation refusal
while enumeration is live, mutation after termination, Resume label cleanup and
a failed single For Each terminator. Existing DAO execution uses a real SQL
recordset clone in With: after retrying a field read, transactions remain refused
inside With and succeed after End With. Clone position is explicitly set with
MoveFirst; no clone/transaction safety guard is weakened. No new CTest target
is introduced. General Access error-handler equivalence remains open.
