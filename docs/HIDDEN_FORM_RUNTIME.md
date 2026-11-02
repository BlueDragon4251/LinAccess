# Native acHidden form opening

OpenForm accepts WindowMode:=acHidden (1), positionally or by name. New native
forms initialize without show/activation, remain in Forms and keep their own
OpenArgs and data context. Open/Load/initial Current run; an initial Activate is
not synthesized. In particular, an OnActivate handler no longer defers Current
forever for a hidden opening. Visible reads False and named Close still works.

Existing native GUI execution audits Open/Load/Current order through actual SQL,
checks that Activate did not run, reads the hidden instance and closes it. No
new CTest target is added.

Normal opening keeps existing activation behavior. Existing visible forms can now be hidden through the native visibility guards,
and first-time reveal is supported. Guarded lifecycle/pending-edit cases, hidden
reports, full window transitions and Windows event equivalence remain open.
Icon/dialog behavior is documented separately in DIALOG_WINDOW_RUNTIME.md. This is not complete window parity.
