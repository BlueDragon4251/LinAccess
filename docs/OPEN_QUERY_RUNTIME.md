# Native query documents

DoCmd.OpenQuery now opens saved SELECT queries as actual native document tabs
when View is acViewNormal and DataMode explicitly acReadOnly. A bounded
editable single-table path also supports default/acEdit, described below. Native JSON query
definitions and imported query artifacts use the real Access query compiler and
AccessQueryModel executor, including checked declared parameters prompted through
QInputDialog. Displayed rows are real materialized SQL results, with editing
flags/triggers disabled. An already open query activates its existing document
without rerunning it. Native/imported name ambiguity refuses 5.

Named/positional/omitted arguments bind with Missing distinct from explicit
Empty. Access defaults are retained: omitted View is normal, omitted DataMode is
acEdit; the bounded editable subset below is supported, and other editable
queries refuse 3251. Null arguments are 94, missing required arguments
450, absent saved names 7874. Other views, action queries and make-table queries
refuse 3251 before execution. Parameter cancellation returns 2501 and creates no
document; actual Currency parameter conversion is covered through SQL results.

The tab belongs to MainWindow, independently of the calling form. Caller Me
remains unchanged. Existing field/document validation runs before switching;
rejected validation, modal windows and project changes retain checked errors.
The unregistered result page/model is owned locally during query execution so
native callbacks cannot destroy it through the project window mid-execution.

Named/default DoCmd.Close for query documents and Ctrl+W/Ctrl+F4 close the actual
tab; save options validate but have no pending design session to write. Readonly
results contain no record drafts. CurrentData.AllQueries.IsLoaded tracks actual
tab membership, including background tabs and immediate logical close before
deferred destruction. Window-menu/document cycling uses the real table view.

Existing main-window integration executes native/imported SELECT definitions,
reads actual result values/readonly flags, checks reuse, default/named close,
parameter cancel/checked Currency, rejects action queries without data changes,
and verifies a query remains after its VBA opener closes itself. No new targets.

This readonly path is complemented by the bounded editable path below. General
updateability, other views, action-query UI/warnings, query design Save,
linked/provider behavior and complete Access
window/event/error-priority parity remain open. SQL compatibility remains the
executor's measured subset; query presence is not execution compatibility.

Microsoft reference:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.openquery

## Requery and Shift+F9

Requery on the active readonly query document and its native Shift+F9 shortcut
reload the current saved definition and prompt declared parameters again. A new
real SQL result is materialized separately; only successful execution replaces
the existing model. Errors and cancelled parameters preserve the old rows/model.
Successful replacement selects the first row and retains the bounded column.

The old document may close during execution without deleting the executing
replacement model or resurrecting the closed tab. Same-query recursive execution
refuses 3251. Project close/replacement is blocked while the SQL executor runs;
project identity and document lifetime are checked before publishing results.
The existing main-window target covers fresh SQL values/definitions, actual
Shift+F9, parameter cancellation/conversion and an actual SQLite scalar callback
that requests reentry/project close and destroys the original page mid-query.
These are native lifetime guarantees, not an Access-wide reentrancy claim.

## Navigation, activation and clipboard

Default-target GoToRecord and the first/previous/next/last RunCommand commands
navigate the actual materialized query result. Explicit acDataQuery with named
or active query targets supports the same actions and acGoTo offsets. Field
position is retained; unavailable records return 2105 without moving, acNewRec
returns 2448 on readonly results. Explicit acDataTable never targets a query.
DoCmd.SelectObject acQuery activates an existing named tab case-insensitively,
validating the departing document and preserving caller Me. Missing open query
selection currently refuses 3251; exact Access closed-query errors remain open.

RunCommand SelectRecord/SelectAllRecords select actual rows/cells. Copy, Ctrl+C
and Ctrl+Insert export the selected contiguous rectangle as quoted TSV through
the same bounded formatter used by native tables. Query cell conversion uses
actual SQL values and proven type metadata; Null is empty, doubles use roundtrip
precision, binary/unrepresentable values fail before changing the clipboard.
Hidden/disjoint selections, NUL, over one million cells and over four MiB UTF-16
text refuse explicitly. This is text interchange, not Access record MIME parity.
Native SendKeys uses actual focused query controls through the existing project
scope/FIFO dispatcher; no global OS injection or inactive-caller fallback.

The existing main-window target checks VBA navigation, boundaries/readonly
rejection, named target activation across two query tabs, actual Ctrl+C/Ctrl+Insert,
hidden-cell refusal, select-record/all and real SendKeys copy. No new test targets.
Full desktop suite 48/48 (91.62 s), GUI-free core 1/1 (1.54 s), Debian package
payload/permissions/Build-ID/pinned SQLite and all four original fixture SHA-256
checks pass. General query editing, readonly query sort/filter and full Access
window/event/keyboard semantics remain open.

## Native result search

Ctrl+F/RunCommand Find open a real modeless Qt search dialog on the query view;
Shift+F4/FindNext repeat the actual previous search. Current-field DoCmd.FindRecord
uses the same materialized result search. Literal text supports partial/whole/
start and Qt case handling, Up/Down do not wrap and All wraps. Number, Currency,
Date and Boolean use checked typed whole-field comparison; Null is skipped.
Only proven types are searched: compiled projection metadata and direct physical
column lineage supply per-position UI types. Unknown expressions/compound output
and unsupported/binary columns are not guessed from displayed cell strings.
The existing DAO result-type map retains its prior contract. Invalid search
input leaves selection/history intact; one million searched cells is the bound.
All-field/formatted/wildcard FindRecord, locale and full Access search parity
remain open, as for the existing table path.

Successful Requery closes the previous search dialog and resets search state;
failed Requery retains the original result. New query tabs select their first
real result cell synchronously, so immediate VBA FindRecord works without an
extra Qt event-loop turn. Search-field Copy/Paste/Cut/Undo/Redo use actual focused
QLineEdit text through the shared data-document command helper. Readonly result
cells remain immutable. Model replacement checks document/view/model lifetime
after Qt selection callbacks before publishing success.

Existing main-window integration covers typed number/Currency and literal text
search, Ctrl+F/Shift+F4, invalid dialog input, focused Copy/Paste/Undo/Redo and
immediate post-OpenQuery Currency search. Currency SQL objects/schema/results and
query model targets pass 5/5 (4.46 s) with the main-window target. Full suite
48/48 (90.46 s), GUI-free core 1/1 (1.50 s), Debian package and all four original
fixture hashes pass. No new test targets.

## Edit menu and readonly zoom

The native Edit menu recognizes query views: Copy and Select All forward real
key events to the view, while grid Cut/Paste stay disabled for readonly results.
The query view handles forwarded Copy directly as well as Ctrl+C/Ctrl+Insert;
all paths use the same checked formatter. Selection ranges are checked before
allocating their individual indexes; query Select All also enforces the existing
one-million-cell bound. No hidden/binary/size guard is bypassed by menu routing.

RunCommand acCmdZoomBox and Shift+F2 open the actual existing native Qt zoom
dialog with a readonly QPlainTextEdit holding the checked current cell text.
Close leaves SQL and model data unchanged; destroyed query/model targets return
2501. A focused modeless search dialog is not treated as the underlying grid's
active Zoom field. Original binary/OLE rendering and full Access Zoom/edit-menu
semantics remain open.

Targeted existing main-window/layout pair passes 2/2 (17.70 s), including actual
menu Copy/Select All, disabled grid mutation actions, VBA Zoom and Shift+F2 with
real readonly dialog content. Debian package payload/Build-ID/SQLite checks and
four original SHA-256 checks pass. Previous full 48/48 baseline is the query-search
checkpoint (90.46 s); this bounded GUI-only increment used the targeted pair.


## Editable single-table query documents (2026-11-03)

Default/acEdit OpenQuery supports an unparameterized `SELECT * FROM table` over
one physical main-schema user table with a real primary key. The native
AccessTableModel/View executes actual SQL and uses the existing typed manual
submission, editor validation, record navigation/search, clipboard and record
commands. The document remains an actual query for named Close, SelectObject,
CurrentData.AllQueries.IsLoaded and explicit acDataQuery navigation; it does not
mark the underlying table as an open table document. Its lifetime is independent
of the calling form. First-cell selection is synchronous for immediate VBA.

Save writes actual cached record changes. Close validates/submits records even
with acSaveNo (that option concerns design), retaining the open dirty document
on validation/SQL failure. Requery supports an unchanged captured definition,
saving the current sole draft through the existing table path. A changed saved
definition refuses 3251 without rebinding drafts. Opening the same query in a
different data mode refuses 3251. A second editable native datasheet for the same
physical table is refused before switching documents; readonly views may coexist.
This is a bounded safeguard, not Jet record locking or concurrency parity.

Projection/joined/parameterized queries, views, missing primary keys,
temporary shadows and hidden/generated columns are outside this editable subset.
Save/Requery verify the captured query still resolves to the same eligible main
table and that its CREATE TABLE SQL matches the opening snapshot. Changed source
structure refuses 3251 and retains drafts; unrelated schema changes do not
invalidate the snapshot. This does not detect every drop/recreate or external
row update, and does not promise general concurrent-update safety.

Existing main-window/query-model targets cover actual default VBA OpenQuery,
SQL save/requery, changed definition refusal, CHECK failure/retained dirty Close,
record commit with acSaveNo, caller self-close, source metadata/temporary shadows,
and refusal of a second writable source document without changing active focus.
No new CTest targets were added.

Full desktop 48/48 (90.33 s), fresh GUI-free core 1/1 (1.48 s), Debian package
payload/permissions/Build-ID/pinned SQLite checks and four original fixture
SHA-256 checks pass for this editable-query implementation.


## Filtered single-table editing and key proof (2026-11-03)

The editable subset now accepts a bounded WHERE predicate after its single main
table, retaining the compiler-translated condition in the real SQL table model.
Comparisons, boolean combinations, parentheses, literal IN lists and NULL checks
use that compiled SQL predicate. Subqueries, extra clauses, predicate function
calls, parameters, projections and joins remain excluded. This is not arbitrary
Jet predicate/type/collation parity. Save/requery continue using the captured
filter; a saved value that no longer satisfies it disappears on the model's
post-commit SQL refresh. The underlying keyed row is still actually updated.
The existing GUI regression verifies this against real SQL, with an excluded
control row and no new test target.

Editable source keys must be structurally non-null: declared NOT NULL primary
keys or a proven single INTEGER rowid alias. Nullable TEXT/composite keys and
INTEGER PRIMARY KEY DESC aliases with a separate primary-key index refuse 3251.
SQLite's allowance for duplicate NULL primary-key values is not treated as a
safe record identifier. This remains a native datasheet safeguard, not full
Jet locking, concurrent update or Access refresh behavior parity.

Full desktop 48/48 (89.83 s), fresh GUI-free core 1/1 (1.49 s), Debian package
payload/permissions/Build-ID/pinned SQLite and all four original fixture SHA-256
checks pass; no new CTest targets.
