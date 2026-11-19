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

## Access 2003 same-row navigation

Hidden Access 11.0 and DAO 3.6 automation created a fresh Jet 4.0 database
with two rows, `(1, 10)` and `(2, 20)`. Each operation started with a new
recordset and either an edited value of 999 or a pending insert `(99, 999)`.
The table-type cursor used the primary index. The dynaset used the local
table as its source. No Update ran. A separate read after the matrix measured
two stored rows and a value sum of 30.

| Operation targeting the current stored row | Table Edit | Table AddNew | Dynaset Edit | Dynaset AddNew |
|---|---|---|---|---|
| MoveFirst on the first row, MoveLast on the last row | Discard | Discard | Discard | Discard |
| Assign current Bookmark, Move 0 with current Bookmark | Discard | Discard | Discard | Discard |
| Move 1 from the preceding row's Bookmark | Discard | Discard | Discard | Discard |
| Assign current AbsolutePosition, FindFirst/FindLast matching the current row | 3251 | 3251 | Discard | Discard |
| Move 0 with StartBookmark omitted | Discard | Discard | Discard | Clear buffer, with Null fields |

Discard returned EditMode 0 with the stored row's values and BOF, EOF and
NoMatch all false. Error 3251 preserved EditMode and the pending value.
For dynaset AddNew followed by bare Move 0, EditMode became 0 and both field
values were Null. A second probe on an ordered dynaset confirmed a populated
AbsolutePosition and a readable byte-array Bookmark in this state, with BOF
and EOF false. Assigning the saved pre-AddNew Bookmark restored a stored row.

LinAccess discards the buffer for the measured same-row operations with a
stored row immediately available. The dynaset AddNew case with omitted
StartBookmark retains the existing 3219 guard until its cursor state can be
represented. Same-row Seek and LinAccess's table-type Find and AbsolutePosition
extensions retain their existing guards. A separate dynaset FindFirst for the
uncommitted value reported NoMatch in Access while the preceding stored row
remained readable. Failed-find positioning is outside this change.

The existing DAO cursor regression failed with error 3219 against the previous
DAO implementation and passed after relinking the patched DAO. It checks the
current bookmark, committed field values including Null, and persistence.
Pending Clone, Index changes, Requery, GetRows and unsupported cursor families
retain their separate restrictions. Native form UI BeforeUpdate/Cancel/save
semantics are unchanged; this is the DAO copy buffer, not an implicit form save.

Earlier different-row coverage passed the dao_execution bundle in 3.67 seconds.
Table and dynaset matrices cover Edit and AddNew discard across first/last/next/previous/relative,
AbsolutePosition, Bookmark and Find movement. Invalid positions/criteria/foreign
bookmarks retain a dirty buffer. Independent SQL checks prove constant row count
and no discarded777/999 values. Existing indexed Seek checks remain passing;
no separate Windows pending-Seek oracle is claimed. No new CTest/catalog IDs;
456/386/70 unchanged. F054/F055 remain open. No GUI/Original/DEB refresh here.

Public DAO contracts (not an Access2003 execution oracle):
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-move-method-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-edit-method-dao
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-addnew-method-dao
