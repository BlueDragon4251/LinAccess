# Bounded native DoCmd.FindRecord

DoCmd.FindRecord dispatches to the actual active native table or native form,
while Me remains bound to the calling form. It uses real SQL-backed record
ordering, field validation, record saving and selection. No hidden surrogate
form, fake search result or original MDB mutation is involved.

The implemented subset searches the current visible declared text field using
acSearchAll/acUp/acDown, literal nonempty text, part/whole/start matching,
MatchCase and FindFirst. FindFirst=True starts at the first record; False starts
after the current record in the requested direction. SearchAll wraps; Up/Down
stop at the record boundary. No match completes normally, with no Boolean return.
Tables use the bounded SQL-fetch search described in TABLE_FIND_RUNTIME.md.
Forms use the existing native record/validation path, including readonly fields.
A canceled form field/record validation or event abort returns 2501 to VBA before
any following statement can execute.

Named/positional omitted arguments are Missing, distinct from explicit Empty.
Defaults are Match=acEntire(1), MatchCase=False, Search=acSearchAll(2),
SearchAsFormatted=False, OnlyCurrentField=acCurrent(-1), FindFirst=True.
Explicit Empty coerces normally: for example Match becomes acAnywhere(0), and
FindFirst becomes False. It is not silently replaced by a default. Null arguments
retain checked error 94; invalid option values are error 5.

Numeric/date FindWhat, formatted search, all-field search (acAll=0),
wildcard/empty search and general Jet/Windows comparison equivalence are
explicit unsupported boundaries (3251). Current nontext fields and unsupported
record sources are also bounded; full Access error precedence and data-selection
appearance remain unverified. The existing native UI search stays explicitly
literal and does not claim these broader APIs.

An explicit table context prevents later document commands in the database
explorer from falling back to the old caller, including a never-activated caller.
A form's ordinary initialization context is preserved until it explicitly
switches documents. Source form/table validation and modal guards still apply.

References:
- [DoCmd.FindRecord](https://learn.microsoft.com/en-us/office/vba/api/access.docmd.findrecord)
- [AcFindMatch](https://learn.microsoft.com/en-us/office/vba/api/access.acfindmatch)
- [AcFindField](https://learn.microsoft.com/en-us/office/vba/api/access.acfindfield)
- [AcSearchDirection](https://learn.microsoft.com/en-us/office/vba/api/access.acsearchdirection)

Existing main-window/form-Find/layout targets pass 3/3 (15.83 s), with actual
interpreted named/omitted calls, explicit Empty defaults, current-field searches,
first/next origins, normal no-match completion, unsupported/Null option handling,
readonly control search, retained BeforeUpdate drafts and caller Me separation.
GUI-free rebuild/core passes 1/1 (1.99 s) after the expression-binding/constant
changes. No new test targets.

Full rebuilt desktop suite passes 48/48 (88.05 s). Debian payload/permissions/
Build-ID/pinned SQLite checks and all four original fixture hashes pass. The
GUI-free core baseline after binding/constant changes is 1/1 (1.99 s); subsequent
form extensions are Qt-only. No additional test targets.

## Native search direction extension (2026-11-03)

Table/form dialogs and DoCmd.FindRecord now support Up/Down without wrap and
SearchAll with wrap. FindNext retains the chosen direction. FindFirst starts at
the first record; otherwise the API starts after the current record in that
direction. Existing main-window/form-Find/layout trio passes 3/3 (16.70 s),
covering real opposite-boundary refusal, reversed matching, UI and interpreted
named direction arguments. Last complete suite/package baseline is a71acc9 /
private API f9367b0, 48/48 (88.05 s), GUI-free core 1/1 (1.99 s). This Qt-only
extension is focused-verified; no extra targets.
