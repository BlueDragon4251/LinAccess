# Native DoEvents execution

The imported native form/report host dispatches pending Qt events with
`QCoreApplication::processEvents(QEventLoop::AllEvents)`. VBA DoEvents returns
Integer zero, including expression calls, bare statements and VBA.DoEvents().
This is the Access-host result, not a count of standalone Visual Basic forms.

Native execution retains closed document objects while nested event frames are
active. An accepted DoCmd.Quit requested by a dispatched event stops the suspended
VBA frame as well; On Error Resume Next cannot turn application termination into
continuation. Ordinary document closing retains the existing lifetime rules.

Hosts without a native event loop reject the operation with 3251. Calculated
Access expressions also reject it explicitly. Arguments produce 450. Complete
Windows message ordering, SendKeys queue equivalence, reentrant event semantics
and converted legacy form-host support remain outside this increment.

Verification uses existing core/layout targets. The document-lifetime regression
queues an actual Qt callback that inserts SQL data, then executes a nested Quit
from a subsequent callback while the outer procedure is suspended in DoEvents.
