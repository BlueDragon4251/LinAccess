# Native DAO deleted cursor positions

A successful physical Delete now leaves affected cursors at a deleted logical
position instead of resetting the deleting cursor to the first row or treating
another row at the same materialized offset as current. Surviving neighboring
physical identities anchor that position across native family reloads. Shared
clones which pointed at the removed row are marked too; deleted bookmark tokens
remain invalid even if a physical id is later reused.

MoveNext and positive Move from the deleted position select the next surviving
row without skipping it. MovePrevious selects the preceding row. Bookmark,
absolute movement, FindNext and other existing navigation clear the deleted
marker only when selecting a position. Explicit invalidating Requery resets it.
BOF/EOF remainFalse at an unrepositioned deleted current record, including after
removing the last remaining row. Forward/backward empty-boundary state is kept
separately from ordinary initial-empty state. An actual Delete/MoveNext VBA loop
can therefore delete the complete source and reach EOF normally.

Current-row reads/bookmarks remain guarded by native3021 at a known deleted
position; exact Windows3167/3021 distinctions and every property/error-order path
are not claimed. A pending Edit whose row was removed by a clone now checks current
validity before indexing ids or issuing SQL. It cannot write the next row or use
an invalid negative index. The pending buffer remains available for cancel or
navigation. AddNew has no valid old current row at a deleted marker and keeps the
previous bounded inserted-row fallback. Snapshot/frozen refresh, all external
concurrency/reordered neighbor semantics, providers/locking and complete deletion
identity remain open. Reload error restoration includes the new cursor state.

App/DAO built; one existing dao_execution bundle1/1 passed3.40s. Actual table and
dynaset cases cover middle/last/only deletion, clone current/bookmark invalidation,
forward/backward/relative movement, explicit empty Requery, and a clone deletion
while another cursor has a dirty Edit buffer. Independent SQL proves no neighbor
was overwritten. An actual typed VBA loop deletes all three source rows, verified
by SQL COUNT0. No new CTest/catalog IDs;456/386/70 unchanged; F054/F055/F158 full
semantics remain bounded. No fresh GUI/Original/DEB evidence for this scope.

Specific public property/movement contracts (not an Access2003 execution oracle):
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-bof-property-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-eof-property-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-movenext-method-dao
The modern generated Delete-method pages say unsupported despite the actual DAO
method/interface and these property contracts; they were not used to remove DAO
Delete or claim universal engine parity.
