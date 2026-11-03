# Native project form/report catalogs

CurrentProject.AllForms and AllReports expose actual saved native definitions
from _la_objects and imported form/report artifacts from _la_access_artifacts.
Count is Long; Item accepts a zero-based numeric index or case-insensitive name,
including explicit/named Item(Index:=...) and indexed property shorthand.
For Each returns native AccessObject descriptors with readonly Name and IsLoaded.
Collections and item references preserve identity while held, independently of
the requesting form, and inherit checked project-generation invalidation (3420).

Catalog reads query real SQL storage. Imported versions of the same logical
case-folded name use the latest stored artifact, matching the native opener's
version preference. Native name collisions and native/imported collisions raise
5 rather than choosing a source silently. Ordering is native Qt case-folded name
order; full Jet locale/index-order equivalence is not claimed. Catalog work is
bounded to 1,000,000 stored source rows. Missing Item is 3265, Null index 94,
wrong argument count 450 and unsupported/readonly members 438.

Enumeration and descriptor names are snapshots of saved names; new catalog
queries reread storage. IsLoaded examines live native state: completed native
preview initialization, including hidden, modal and embedded documents, excluding
closed/cancelled documents and children whose owning preview has closed. The
native form runner and stored form/report designers carry project/connection/
generation tags and count while their actual dialogs are visible. Closing a
requesting form does not destroy independent metadata references. Reopening the
project, even the same file, invalidates old catalog/item references.

The existing main-window integration exercises real SQL catalog definitions,
VBA For Each and named lookup, missing/Null/ambiguous names, actual legacy form
runner dialogs, hidden reports, modal forms, embedded form parent-close and
project-generation changes. No extra CTest targets.

This is not a full Access/COM catalog: unsaved designer objects, other catalog
collections, mutation, AccessObject state/property sets, typed Access class
coverage, exact Open/Load/Unload timing and complete locale/error priority remain
open. Native descriptor snapshot/ordering behavior is explicit.

Microsoft references:
- https://learn.microsoft.com/en-us/office/vba/api/access.currentproject.allforms
- https://learn.microsoft.com/en-us/office/vba/api/access.currentproject.allreports
- https://learn.microsoft.com/en-us/office/vba/api/access.accessobject.isloaded

Catalog schema and row fetch failures propagate 3146 rather than returning a
partial count. Stored definitions are explicitly read from main.sqlite_schema
and main._la_objects/main._la_access_artifacts; temporary names do not shadow
the saved project catalog. A database without either native catalog table
refuses 3251 instead of inventing an empty catalog. A newly initialized native
project with the actual empty catalog tables correctly returns zero counts.
The native integration target checks a real temporary shadow table, an unrelated
SQLite file without catalog tables and a newly created empty project.
