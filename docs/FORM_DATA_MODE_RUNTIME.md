# Native form data modes

DoCmd.OpenForm accepts acFormPropertySettings (-1, default), acFormAdd (0),
acFormEdit (1), and acFormReadOnly (2) for new native form instances. Modes reach
construction before Open/Load and widget initialization.

Edit enables AllowEdits/AllowAdditions/AllowDeletions and disables DataEntry.
ReadOnly disables the three policies. Add enables AllowAdditions and starts the
real direct-table DataEntry session, excluding pre-existing rows and persisting
new entries. Add currently retains imported edit/delete policies; the complete
Windows override matrix for those two properties is not established. Query
sources retain existing native read-only limits; Add/DataEntry rejects active
non-table sources explicitly rather than pretending to write them.

Existing native GUI regression checks Open-event policies, actual read-only
editors, Edit overriding disabled imported policies, a real saved keyboard edit,
and named Add opening over pre-existing rows with a real inserted row visible
only in its new entry session. Core/GUI targets pass 2/2 (1.94 s); desktop build
passes. Original-fixture follow-up is recorded separately in PROGRESS.md.

These are form policies, not database authorization. VBA/DAO behavior remains
the existing native object model. Changing DataMode on an already-open document,
all view/locking combinations, writable query entry sources and full Windows
property/event/reference equivalence remain open. Generic hosts without native
data-mode support still reject these modes explicitly. See DATA_ENTRY_RUNTIME.md
for session behavior and limits.
