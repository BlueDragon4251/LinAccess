# Native CurrentProject file metadata

CurrentProject and Application.CurrentProject expose a native object with
read-only String Name, Path and FullName from the actually opened SQLite
project file. Name includes its native filename extension; Path is the absolute
containing directory and FullName the absolute opened path, without resolving
symlinks to a different canonical filename. Original MDB source locations are
not substituted for the running native project's location.

A native MainWindow caches weak object references by connection/database
identity. Repeated global/qualified access shares object identity while a
reference is held, including access from different requesting forms. Objects
hold no requesting form reference: closing that form does not close the project.
A closed connection, destroyed owner or changed project generation raises 3420.
Accepted project replacement advances the generation and clears the cache,
including reopening the very same filename. Cancelled file dialogs or cancelled
project Unload preserve the current generation and references. Foreign preview
connections under the same window have separate cache entries and do not retain
closed database handles through the cache.

The main window resolves its project filename to an absolute path before
closing/replacing a project, keeping metadata stable across working-directory
changes. Standalone relative, memory/URI/non-native database names do not become
invented filesystem paths;
file metadata access there refuses 3251. Assignment and unsupported members
raise 438; property arguments raise 450. TypeName is CurrentProject. Locally
bound variables and user functions named CurrentProject retain normal lookup;
Application.CurrentProject explicitly selects the application property.

Native MainWindow lifetime/identity is covered with actual file-backed projects,
a second preview connection, caller close, cancelled project replacement, same-
file reopen and different-file replacement. Pure interpreter regression also
checks a user function named CurrentProject. No new CTest targets were added.

This supplies file metadata, not a complete Access CurrentProject/COM object.
Connection, Properties, AllForms/AllReports/other catalog collections, IsTrusted,
ProjectType, arbitrary project creation/closure and exact Access error priority
remain unsupported. Standalone previews without a native MainWindow do not
claim the window's cached identity or same-file generation tracking.

Microsoft API references:
- https://learn.microsoft.com/en-us/office/vba/api/access.application.currentproject
- https://learn.microsoft.com/en-us/office/vba/api/access.currentproject.name
- https://learn.microsoft.com/en-us/office/vba/api/access.currentproject.path
- https://learn.microsoft.com/en-us/office/vba/api/access.currentproject.fullname
