# Native Close and combo Dropdown commands

The late-bound DoCmd object is available through `DoCmd` and `Application.DoCmd`
in a native document context. Its implemented Close member accepts zero through
three supplied positional arguments: object type, name and save option. acDefault,
acForm, acReport and the three acSave constants are recognized. Names match without
case sensitivity within the same project connection. With no name the active
matching document is used, falling back to the calling document when necessary.
Unsupported object types, missing/ambiguous targets, bad arguments and cancelled
closes fail explicitly. Numeric enum arguments receive checked Long conversion.

Closing uses actual Qt close events and the existing record validation, Unload,
Deactivate and Close lifecycle. Unload cancellation returns VBA error 2501 and
keeps references alive. Recursive close during an already-running close cannot
re-enter the lifecycle indefinitely. Successful closure invalidates native control
and form references and removes the document from Forms/Reports collections.

Save options concern object design, not record transactions: acSaveNo still saves
validated record edits before Unload. This preview has no pending design-edit
session. The command does not implement saving or discarding designer changes.
Omitted positional slots, named arguments, full application-owned DoCmd lifetime,
closing non-form/report objects, all focus/activation corner cases, and full
OpenForm/OpenReport/window semantics remain open. This implementation is not a
complete DoCmd adapter; SetWarnings/action-query warning semantics remain open.

ComboBox.Dropdown uses the real native QComboBox popup. It loads a supported
Table/Query RowSource on first use, preserves typed bound values, and refuses
hidden, disabled or non-combo controls. A native mouse selection is tested against
actual SQL rows. Value-list/field-list sources, complete column presentation,
automatic row-source refresh after every dependent-control change and the complete
Access combo event sequence remain open.

Regression evidence:

- `tests/close_command_tests.cpp`: native form/report selection, cancellation,
  recursive close, reference invalidation, record saving despite acSaveNo,
  numeric/argument failures, active-document close, SQL-backed Dropdown and a real
  mouse selection with a numeric result.
- `tests/original_control_tests.cpp`: the hash-verified original form and original
  module `dbd3623b9404e490a814b70de6318d2c0877969b7feef081696d3e235ae3b427`
  initialize their six previously tested properties; their actual original Close
  button executes unchanged and closes the native form without an error dialog.
  This isolated control/lifecycle test does not provide its original record source.
- `tests/message_box_tests.cpp`: the unchanged original module
  `fa72dad21f4875398da7de601750d963c1d5d1292a36f3db6fbc240ed66959ba`
  executes both missing-input paths. The first opens the actual SQL-backed combo
  after its original dialog; the second focuses the other field. Inputs and rows
  are synthetic non-credentials. This does not establish login/authentication or
  an ActiveX/COM adapter, and no original message/input text is logged.

Stage validation: all 18 CTests passed (61.98 seconds), the GUI-disabled core
build/test passed, a freshly built Debian package matched the current executable's
Build ID and passed content/permission checks, and all four original fixture
SHA-256 checks remained unchanged. Remote distro CI and Windows reference behavior
are not established by these local checks.
