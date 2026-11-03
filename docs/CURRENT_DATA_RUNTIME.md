# Native CurrentData catalogs

CurrentData and Application.CurrentData expose AllTables and AllQueries using
real native SQLite storage. AllTables reads physical user tables from the main
schema, excluding _la_ and sqlite_ internal names consistently with the browser.
AllQueries reads saved native query definitions and imported query artifacts;
generated SQLite views are not counted as separate saved queries. Temporary and
attached namespaces do not replace the saved main catalog.

Both catalogs share checked Long Count, indexed/name/named Item, For Each and
readonly AccessObject Name/IsLoaded with the project catalogs. CurrentData,
catalogs and held item descriptors preserve native identity and remain valid
after requester close; accepted project replacement invalidates the complete
reference graph with 3420. Local variables/user functions named CurrentData keep
the existing namespace lookup behavior.

Table IsLoaded uses actual tab membership, including inactive background pages;
a removed tab is closed even before deferred widget destruction. Query IsLoaded
tracks actual named native SQL editor/source/result dialogs. This reports UI
state, not whether a DAO query or a query used by a form has an active cursor.
A stored query's presence is not a claim that its entire Jet SQL executes.

Existing integration checks actual SQL table/query names, internal-table
exclusion, VBA enumeration and named lookup, a real query artifact dialog,
readonly/background/closed table pages, requester close, same-file project
replacement and genuine empty newly initialized catalogs. Core tests preserve
a user function named CurrentData. No new CTest targets.

Limits and catalog snapshot/version/ambiguity rules from PROJECT_CATALOG_RUNTIME.md
remain. Full Access/COM data catalogs, linked-table metadata, raw-view/unsaved
query equivalence, exact locale/index/error ordering, class typing and complete
object state/properties remain open. CurrentData is readonly metadata and does
not introduce a fake Connection or database mutation API.

Microsoft references:
- https://learn.microsoft.com/en-us/office/vba/api/access.application.currentdata
- https://learn.microsoft.com/en-us/office/vba/api/access.currentdata.alltables
- https://learn.microsoft.com/en-us/office/vba/api/access.currentdata.allqueries
