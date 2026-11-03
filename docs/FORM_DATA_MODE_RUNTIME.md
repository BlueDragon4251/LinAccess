# Native form data modes

DoCmd.OpenForm accepts acFormPropertySettings (-1, default), acFormEdit (1), and
acFormReadOnly (2) for new native form instances. The latter modes override
AllowEdits/AllowAdditions/AllowDeletions before Open/Load and widget construction:
all enabled for Edit, all disabled for ReadOnly. Existing control/read-only and
record command adapters enforce those form policies. Query sources retain their
existing native read-only restrictions; Edit does not make joins editable.

Existing GUI regression checks policy values during Form_Open, the actual
read-only editor, Edit overriding imported disabled policies, and a real keyboard
edit persisted through native Dirty=False to SQLite. Core and GUI targets pass
2/2 (1.80 s); desktop build passes. No new target.

This is a form policy, not database authorization. VBA/DAO behavior remains the
existing native object-model behavior. acFormAdd/DataEntry session filtering,
changing DataMode on an already-open document, all view/locking combinations and
Windows event/reference equivalence remain open. Unsupported modes fail before
new window construction; generic hosts without data-mode support still reject
these modes explicitly.
