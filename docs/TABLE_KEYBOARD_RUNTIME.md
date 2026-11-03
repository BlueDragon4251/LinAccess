# Scoped native SendKeys

Global SendKeys and DoCmd.SendKeys(Keys, Wait) share the real Qt dispatcher.
Named and omitted method arguments are supported; omitted Wait is False.
Malformed sequences are rejected before emitting keys; Null raises 94.

Each key resolves actual focus again. Active table cells, their modeless Find
and modal Zoom dialogs participate, using explicit widget parent traversal
across Qt window boundaries. The originating database connection and database
identity stay fixed. Switching to another project cannot redirect old keys.
Database explorer SendKeys remains unsupported (3251); a missing eligible
receiver fails with 2110. Asynchronous failures appear in the requesting
form status. Closing the requester or stopping/Quit cancels remaining keys.

Batches retain FIFO order. Continuations are scheduled before native delivery,
allowing subsequent keys to run inside a key handler's modal event loop.
Wait=True waits for handler completion; Wait=False queues delivery. DoEvents
drains queued keys through the interactive Qt event loop. Transient activation
may settle for up to 500 ms. Limits: 4,096 keys per sequence, 1,000 repeats,
16 parser nesting levels, 65,536 queued keys and 32 nested native key handlers.
These are explicit resource boundaries, not measured Access limits.

Real form TextBox typing changes Text first. Value remains the original bound
value until validation/commit; Save then reaches actual SQL. Table typing uses
its checked draft cache and normal Save pipeline. No OS-wide key injection,
external application targeting, parentless table context menus or root file
picker routing is claimed. Full Access keyboard grammar/timing equivalence
remains open.

Existing main-window, layout and original Kassa event targets exercise ordered
queued batches, DoEvents, focus changes, dialog continuation, named method calls,
checked typed edits, requester isolation and SQL after Save. The original
Wait=False probe yields until the native queue is drained rather than assuming
one Qt processEvents call delivers a complete sequence; its content, receiver
and event assertions remain intact. No additional CTest targets were added.
