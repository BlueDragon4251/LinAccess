# DAO navigation and pending Edit/AddNew buffers

Successful movement to a different cursor position now discards the pending
Edit/AddNew buffer without Update, matching the documented DAO navigation
contract. The common movement point clears pending values, changed fields and
binary chunk state, then enters Browse. Source SQL records are never written by
this discard. Other clone cursors and their independent pending buffers remain
separate. Existing generated-number reservation behavior is unchanged.

MoveFirst/Last/Next/Previous, Move, AbsolutePosition, Bookmark, Find and Seek use
that movement point after target/argument/search validation. Seek preserves its
entry-mode/revision reentrancy guard; Move/AbsolutePosition recheck mode/revision
and connection validity after scalar argument evaluation. Failed searches that
move to BOF/EOF set their existing NoMatch behavior after navigation. Invalid
positions/bookmarks/criteria preserve the pending buffer. Full Windows error
precedence and every reentrant/provider path are not claimed.

Pending no-op movement still refuses3219, preserving the prior bounded guard:
behavior when the target equals the current position has not been measured.
Pending Clone, Index changes, Requery, GetRows and unsupported cursor families
retain their separate restrictions. Native form UI BeforeUpdate/Cancel/save
semantics are unchanged; this is the DAO copy buffer, not an implicit form save.

App/DAO built. One existing dao_execution bundle passed3.67s. Real table/dynaset
matrices cover Edit and AddNew discard across first/last/next/previous/relative,
AbsolutePosition, Bookmark and Find movement. Invalid positions/criteria/foreign
bookmarks retain a dirty buffer. Independent SQL checks prove constant row count
and no discarded777/999 values. Existing indexed Seek checks remain passing;
no separate Windows pending-Seek oracle is claimed. No new CTest/catalog IDs;
456/386/70 unchanged. F054/F055 remain open. No GUI/Original/DEB refresh here.

Public DAO contracts (not an Access2003 execution oracle):
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-edit-method-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-addnew-method-dao
