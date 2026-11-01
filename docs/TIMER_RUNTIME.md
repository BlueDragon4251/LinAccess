# Native form timers

Form TimerInterval is imported and exposed as a writable Long interval in
milliseconds (0 disables, maximum 2147483647). A form-owned precise QTimer
dispatches the bound Form_Timer procedure, or a supported function-expression
binding, through the existing native VBA runtime. Form_Load can configure the
interval; firing starts after opening. Changing the interval restarts the timer.
Hidden forms continue receiving events. Invalid imported values are rejected;
runtime Null/negative/overflow writes raise 94/5/6 without changing the setting.

Successful document/project close stops the timer. Cancelled Unload leaves it
running. Dispatch skips loading, saving, navigation, record commands, teardown
and an already running handler; missed ticks are not queued for replay. Native
execution lifetime guards defer destruction until a timer handler that closes its
own form returns. Handler failures remain visible in OnTimer results/status.
Reports do not expose this form adapter (438).

Native asynchronous regressions assert actual SQL changes, a three-tick self-stop,
dynamic restart, hidden execution, cancelled Unload, self-close without later
ticks, handler errors and a function-expression binding. Deadline waits return
bools checked by assertions in the standalone test executable.

An unchanged original Kassa clock module (SHA-256
265db230b1d253b070d2a58047e7db80ccc1d4b9f4e7ff38739e195371bc7deb)
also runs through its original OnTimer binding, form caption and native field. The original
layout/module hashes are verified before preparing the test: unrelated lifecycle
bindings are removed in memory and TimerInterval is set to 5 ms for the probe.
Both targets are first set to an old timestamp; the original Now assignments
then replace them and are checked against independent local timestamp bounds.
The probe uses an isolated imported project and rolls back its savepoint. It is
not a claim that the original startup/full application ran.

Exact Access modal/reentrancy/message scheduling, hard real-time guarantees,
original clock locale/display formatting, Windows event traces, action macros
and report timers remain outside this adapter. Original timer execution is one
handler: the second textual occurrence is the duplicate SaveAsText export, not
an additional executed procedure.
