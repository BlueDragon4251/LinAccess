# Native forms over saved queries and SELECT sources

The owned-form editor now accepts a physical table, a saved query name or a
single-line non-parameter SELECT in its editable data-source control. The existing
definition table key retains the source text/name; field order/captions/geometry,
VBA and original archive bytes retain their existing storage contract.

Physical-table field behavior remains unchanged. Query field names come from
actual schema-only preparation without stepping rows or evaluating SQL expressions.
The generator validates up to 255 selected, existing, unambiguous fields. Native
blank/unbound forms and extra controls are now also supported (FREE_FORM_CONTROLS.md). Action and
parameter sources are refused. SELECT/view/query forms emit AllowAdditions/Edits/
Deletions=False and omit automatic New/Save command buttons; they execute actual
query records and Requery through the native form host. General query updateability,
parameter binding, original layout round trips and full Access parity remain open.

The common accessQuerySource reader now also resolves definitions from
main._la_objects, preserving a saved query name as RecordSource. Current SQL is
read on Requery; no artificial rows, persistent cursor or source snapshot cache
is introduced. Invalid/ambiguous owned definitions and owned/imported name conflicts
fail explicitly. Imported source lookup keeps its latest-artifact behavior and
uses the main catalog. The existing temporary-view installer now also reads owned
query definitions and revisions. Supported non-parameter SELECT dependencies use
the same ordinary/ranked view resolution as imported queries. Ambiguous owned/
imported names receive no executable view. Unsupported SQL/dependencies remain
explicit failures; this is not complete Jet dependency or parameter support.
Non-table form/report loads and Requery refresh views first; a failed refresh is
an explicit source error rather than stale output. Designer metadata likewise
refuses a busy view refresh.

An existing main-window flow prepares SELECT 42 field metadata, saves a form over
an actual owned COUNT query, opens it through the explorer, checks read-only values
and absent edit buttons, changes its real source table, calls Requery and observes
the updated count, then reopens the designer and checks the saved source/fields.
Existing physical-table/native subform/VBA/save cases still run. No new CTest target
or fixed-catalog item.

The existing UI case now uses a two-query chain whose dependency sorts after its
consumer, reads its real row via Application.Eval/DLookup, changes the base-query
SQL while the form is open and verifies Requery updates the dependent count to
zero. No persistent source snapshots or invented fallback rows are used.

The same native visual query designer now authors form SELECT sources, including
its bounded calculated fields, INNER/LEFT joins and grouping. See
VISUAL_QUERY_DESIGNER.md for exact limits. Matching design metadata persists with
the form and participates in form-level undo/redo. It does not add writable
query forms or parameter binding.
