# Native literal text replacement

Native table Replace supports whole-field, substring and prefix modes on a
selected raw text column. Current Replace changes the first substring occurrence
(or the single prefix/whole value) in the pending record; normal Save still
controls its SQL write. Replace All replaces every nonoverlapping substring
occurrence in each matched field, or one prefix/whole value, and commits the
entire matched batch through the existing transaction path. Case matching uses
the existing native UI comparison. Count feedback reports matched fields.

A null replacement QString means empty text, not SQL NULL; NULL source fields
remain skipped. Unicode/NUL/size validation is shared by the current and batch
paths. No wildcard or formatted-value reverse replacement is enabled. Batch
output is fully prepared before cache mutation, bounded to16Mi UTF16 units and
one million rows, with a five-second scan bound. Primary-key/readonly/dirty-cache
batch restrictions, active-record identity checks and native SQL constraints
remain in force. A failed SQL batch rolls back SQL and retires all of its drafts.

This extends TABLE_REPLACE_ALL_RUNTIME.md's older whole-field-only boundary.
General within-field FindNext cursor/highlight ordering, wildcard replacement,
full locale/Windows collation and Access event/undo equivalence remain open under
F373/F340/F168. No new Windows Access observation is claimed.
