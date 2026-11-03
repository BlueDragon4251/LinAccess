# Native form identity after primary-key updates

Saving an existing native direct-table form now consumes UPDATE RETURNING rowid
from the pinned SQLite backend. The form follows the committed identity instead
of reselecting the old rowid after an Integer primary-key change. Existing insert
handling remains on lastInsertId.

The captured old values remain available to AfterUpdate. Existing form recordset
snapshots refresh with a checked identity mapping so already-open clones and
previous byte-array bookmarks follow the same updated row, including a changed
sort position. Mappings reject unknown sources, duplicate targets and ambiguous
identity collisions. Records that leave the visible recordset do not retain a
usable bookmark in that recordset.

Existing GUI regression edits the actual bound key with real keyboard events,
saves it, checks CurrentRecord against sorted rows, follows a pre-existing clone
and bookmark, and verifies the committed SQL value. Desktop build and existing
GUI/DAO targets pass 2/2 (0.90 s). No new target.

This covers committed updates made by the native adapters. Arbitrary external
SQL/AFTER-trigger identity replacements, all filter/deletion/requery event paths,
concurrency and full Windows identity/event equivalence remain open.
