# Native readonly query documents

DoCmd.OpenQuery now opens saved SELECT queries as actual native document tabs
when View is acViewNormal and DataMode explicitly acReadOnly. Native JSON query
definitions and imported query artifacts use the real Access query compiler and
AccessQueryModel executor, including checked declared parameters prompted through
QInputDialog. Displayed rows are real materialized SQL results, with editing
flags/triggers disabled. An already open query activates its existing document
without rerunning it. Native/imported name ambiguity refuses 5.

Named/positional/omitted arguments bind with Missing distinct from explicit
Empty. Access defaults are retained: omitted View is normal, omitted DataMode is
acEdit; unsupported editable/default modes refuse 3251 rather than silently
pretending to be editable. Null arguments are 94, missing required arguments
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

This is a bounded readonly datasheet path. Default/acEdit updateability, other
views, action-query UI/warnings, query record navigation/Save/Copy/Find/Requery,
SelectObject query activation, linked/provider behavior and complete Access
window/event/error-priority parity remain open. SQL compatibility remains the
executor's measured subset; query presence is not execution compatibility.

Microsoft reference:
https://learn.microsoft.com/en-us/office/vba/api/access.docmd.openquery
