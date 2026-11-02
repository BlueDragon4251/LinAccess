# Native DoCmd.RunSQL transaction option

RunSQL now accepts its UseTransaction Boolean, positionally or by name. The
omitted/Empty default is True. True wraps native DAO action execution in a unique
SQLite savepoint; success releases it and failure rolls back/releases it before
propagating the original VBA error. False executes without an extra action
savepoint. Neither path commits or bypasses a surrounding DAO workspace
transaction. Native transaction/cleanup failures are reported explicitly.

The existing GUI execution target performs actual named/positional calls,
rejects Null SQL with 94, checks late unique-constraint failure and successful
subsequent execution, and rolls back a surrounding DAO transaction containing
both successful and failed RunSQL calls. No new CTest target is added.

This is a native SQLite-backed action boundary, not proof of all Jet transaction
or partial-failure semantics. SQLite statement atomicity remains in force even
with False. Cross-connection/distributed transactions, side-effecting query
callbacks and full Jet action/undo/SCODE parity remain open.
