# Currency in initial native table creation

The initial new-table dialog now offers CURRENCY and maps that selection to actual
DAO dbCurrency5 instead of the binary-field fallback. It shows the supported
four-place amount range. Numeric default prechecks apply; final exact precision/
range checking stays in the already implemented DAO creation path. TableDefs
append occurs only after every detached field/default is validated, so a rejected
Currency default leaves no partial table and retains the user's dialog input.

App and existing GUI target built. An initial GUI run hit the existing20s watchdog, as did the first primitive-lookup
optimization retry. Phase timing showed the wizard itself took26ms; repeated
construction of immutable VBA regex patterns was then removed. The existing
core/GUI/DAO bundle passed3/3 in15.18s; mainwindow_documents passed in12.29s with the unchanged20s watchdog. The actual table dialog selects
Currency, refuses a five-place default without creating its table, retains the
field/default text, accepts the corrected four-place value and creates actual
Type5/Size8 metadata. Independent SQL and typed Currency reading verify its real
new-row default. The technical extra field is removed before continuing existing
mixed-type table/field workflows. No new CTest/catalog IDs;456/386/70 unchanged.
Native Currency range/Decimal/full Jet schema and Windows oracle remain open.

The preceding published sourceda72e727 has a fresh checked DEB, existing Original
DAO event check1/1 in120.33s and query inventory1/1 in46.11s. All seven protected
hashes are unchanged. Its100 queries remain78 native executions/22 bounded errors:
73 ordinary and5 test-parameter executions without an Access result oracle;
3 unresolved original columns,9 missing dependencies,10 unsupported Jet/binding.
The existing missing-macro boundary is retained, not a positive initialized sale/
printing proof. This package precedes the initial-table Currency selector.
