# Native table text Find

The actual native Qt datasheet supports Ctrl+F, Shift+F4, the native Edit menu,
and DoCmd/Application RunCommand acCmdFind (30) and acCmdFindNext (341). The modeless dialog belongs to its table;
closing that table destroys the dialog. Opening from a native cell editor checks
and retains the field in the manual-submit cache without immediately writing SQL.
A dialog's search reactivates its owning page when necessary, using the existing
source validation/modal guards. Me stays bound to the original caller form.

The dialog searches visible declared TEXT/CHAR/CLOB columns, either one field or
all available text fields. Matching is literal, with part/whole/start modes and
Qt case-sensitive or case-insensitive comparison. First search includes the
current row; repeated searches advance and wrap through the actual view order.
Matches select and scroll to the actual field. No match is normal completion;
FindNext without a prior nonempty search is 2046. Empty tables can open Find.

The real model fetches rows past Qt's initial SQL batch. Search is bounded to one
million text cells, with no partial-result success when the budget is exceeded.
NULL/BLOB values and hidden rows/columns are skipped. Numeric/date/display-format,
Jet wildcard/Windows collation, reverse direction and Replace are not
implemented by this text-search increment. Bounded programmatic DoCmd.FindRecord
is specified separately in [FIND_RECORD_RUNTIME.md](FIND_RECORD_RUNTIME.md).

Before scanning, the production controller saves a sole changed existing current
record through the checked transactional SaveRecord path and restores its primary
key. Other pending row drafts, insert/delete caches or unsafe record identity
return 3251 without committing them. SQL failure retains the draft. A standalone
view without the production record-navigation guard refuses dirty-cache search.
Opening and closing Find alone does not submit the pending record.

View-level command refusals now set a current LinAccess model error as well as
emitting the UI message. VBA receives the actual checked error instead of a stale
SQL/model error; binary Copy returns 3251 and preserves the prior clipboard.

Existing main-window/Currency targets pass 2/2 (2.49 s): actual modeless dialog,
300 SQL rows, owner-page reactivation, match modes/case, native/keyboard/menu
FindNext, readonly search, checked sole-record SQL persistence, retained multirow
drafts, Ctrl+F from the cell editor, and document/dialog disposal. Binary payloads
remain unchanged across record saves and blocked clipboard Copy. No new targets.

Full rebuilt desktop suite passes 48/48 (92.69 s), GUI-free core 1/1 (2.45 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. No additional test targets. This also covers the preceding
native table sort commands and the final weak-owner reentrancy guard.
