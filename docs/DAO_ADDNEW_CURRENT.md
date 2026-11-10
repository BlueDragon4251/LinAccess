# DAO AddNew/Update current record and LastModified

After a successful AddNew/Update, an existing valid current row remains current
when its identity still exists. The inserted row receives this cursor's
LastModified bookmark independently of current position. Assigning Bookmark to
LastModified navigates explicitly to the inserted row. Edit/Update retains its
existing current-row behavior. Clone positions remain independent; retained
bookmarks survive insertion before the current row in physical ordering.

No-prior-current (empty/BOF/EOF) or disappeared prior identities retain the
previous bounded fallback to the inserted row; their exact Access behavior is not
claimed. Pending-navigation/discard behavior remains separate open work. Full
external concurrency, triggers, dynaset append ordering, allocation/error timing,
unique-index policy and Access2003 cursor oracle remain open.

A real zero physical rowid exposed an existing QSQLITE boundary: Qt6.8's
lastInsertId returns an invalid QVariant when sqlite3_last_insert_rowid is0.
The native DAO insertion now reads the actual connection's signed64-bit insertion
identity directly, preserving zero. No extra SQL or numeric Double conversion.
This does not change generator allocation/reservation or original MDB writes.

App/DAO built. First existing DAO check failed on the zero-id LastModified case;
only that actual identity issue was fixed and the affected DAO bundle rerun:
1/1 passed3.27s. Existing synthetic checks that need the newly inserted row now
navigate explicitly by LastModified. Actual table/dynaset cursor/clone/bookmark
checks cover current retention, insertion before current, zero identity, Edit,
CancelUpdate, unique violation followed by retry and independent SQL row counts.
An actual typed VBA procedure retains row2, then selects inserted row3 through
LastModified. No new CTest/catalog IDs;456/386/70 unchanged, F054/F055 remain open.
No new GUI/Original/package proof for this scope.

Microsoft contract states that the record current before AddNew remains current:
https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/recordset-addnew-method-dao
Qt6.8 zero-id implementation:
https://github.com/qt/qtbase/blob/v6.8.0/src/plugins/sqldrivers/sqlite/qsql_sqlite.cpp
These are contract/implementation references, not a Windows Access runtime oracle.
