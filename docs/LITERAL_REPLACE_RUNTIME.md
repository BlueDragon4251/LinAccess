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

## Multiple visible text fields in one replacement transaction

The real table replacement dialog now offers all visible text fields. Replace All
collects the visible non-key, nongenerated text columns and prepares every matched
cell before staging any change. The model's existing single-column API delegates
to the same multiple-column implementation. SQL updates share one transaction;
a later field/record constraint failure rolls back earlier writes and retires
only this operation's drafts. Actual matched field values, not records, are counted.

Hidden, primary-key, generated, numeric and binary fields are excluded from the
all-field batch. Generated expressions continue to derive normally from changed
inputs. SQL Null remains untouched. Whole-field/substring/prefix modes, Unicode
case comparison and nonoverlapping literal replacements retain existing rules.
Current Replace in all-field scope operates on the actual current text cell as a
pending draft, and retains the all-field search scope for subsequent Find actions.
The dialog explains the batch's key/generated-field exclusions.

The view snapshots table, selected field identities, visibility and the current
record key around record validation. Changed scope/source/current identity refuses
2501 before the batch. Success restores the original key and selected column even
when sorting moves the row. Existing read-only/dirty-cache/transaction guards and
16Mi output/five-second scan/one-million-row bounds remain; matched cells also
have an explicit one-million limit. Duplicate input columns refuse before writes.

Application built. First existing mainwindow/model bundle2/2 passed16.60s. The
actual dialog replaces four values across two columns, preserves TEXT primary
keys/hidden fields/numeric/BLOB/Null values and generated output, stages current
replacement without saving, rolls back a later second-field CHECK error, preserves
an unrelated pending draft and rejects validation-time visibility changes. No new
CTest IDs, Original/Windows/full-suite/package proof. This extends F373/F168, not
full Access locale/wildcard/undo/concurrency equivalence; catalog456/386/70 stays.
