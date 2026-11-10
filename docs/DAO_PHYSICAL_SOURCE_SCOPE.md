# Physical DAO table source scope

Physical native Recordset table reads and INSERT/UPDATE/DELETE now qualify their
actual target in main. Before loading/reloading or entering Edit/AddNew/Delete/
Update, the cursor verifies the physical main table and refuses an equal-name
TEMP table/view with3251. TEMP-only physical tables are unsupported. A vanished
main source retains3420. This prevents SQLite's normal TEMP-first resolution from
mixing cached main field metadata with another table's rows or mutation target.

A held cursor's already materialized main values/origins remain truthful while a
TEMP shadow exists. Update refuses before SQL and retains its pending buffer;
after removal of the shadow, the caller can perform the intended main update.
Close remains available. Data/schema/counter helper names are not reconstructed
from temp data. Scope is physical Recordsets, not a claim of complete shadow/
external-schema policy for every raw SQL/provider/metadata API. Full external
same-name recreation/generation and all concurrency/reentrancy remain open.

App/DAO built; one existing dao_execution bundle1/1 passed3.42s (ctest total3.43s).
Actual main/TEMP tables have different values. A pending Update refuses while
both remain unchanged, then succeeds on main after removing TEMP. Edit/AddNew/
Delete/Requery and new physical opens refuse shadowed names; TEMP-only opening
refuses. Independent SQL and cached origin/value checks prove target isolation.
No new CTest/catalog IDs;456/386/70 unchanged. F054/F055/F121 remain open.

Preceding published sourcea4bbf4a42d23d8ce5eb98c0c3cc5d63c694281fd has a fresh
checked DEB at /workspace/work/dao-dynaset-order/package/linaccess-0.7.0-Linux.deb.
Identity/content/dependencies passed. Existing Original check1/1 passed118.03s;
all seven protected hashes unchanged, missing-macro boundary retained. Already
passed GUI/inventory were not repeated. Last actual inventory is4ebc3f6c:
100 classified/78 executed/22 boundaries, without Windows result oracle. This
DEB predates the physical-source-scope change documented here.
