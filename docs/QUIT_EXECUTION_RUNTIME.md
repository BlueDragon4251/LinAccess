# Successful Quit stops active native VBA execution

After native application close preflight and document Unload accept Quit, the
GUI-thread execution epoch advances. Existing procedure/expression frames stop
without continuing assignments, function operands or suspended dialog callers.
On Error Resume Next cannot suppress that stop. Cancelled Quit does not advance
the epoch: native error 2501 remains handleable and execution can continue.

Native construction scopes keep the stop alive through Open/Load initialization,
including UI-initiated opening without an outer VBA procedure. A startup Quit
marks that document cancelled before later events/showing; public opening returns
no new window. This is not a persistent poison flag for later independent core
engine executions. Cross-thread execution is outside the current scope.

Existing core and native GUI targets check nested expression operands, fresh
independent execution, cancelled Quit, a successful/cancelled Quit from acDialog,
no subsequent child/caller SQL effects, and Quit during Open/Load without a new
shown window. A cancelled/closed preview cannot export a layout PDF.

Variant/Byte array function results now retain typed return values without
forcing legacy text materialization. Existing core cases exercise both.

Full shutdown through arbitrary external C++ callbacks, every Windows error/UI
path and cross-thread/distributed behavior remain open. No new CTest target.
