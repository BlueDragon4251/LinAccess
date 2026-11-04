# Native forms over saved queries and SELECT sources

The owned-form editor now accepts a physical table, a saved query name or a
single-line non-parameter SELECT in its editable data-source control. The existing
definition table key retains the source text/name; field order/captions/geometry,
VBA and original archive bytes retain their existing storage contract.

Physical-table field behavior remains unchanged. Query field names come from
actual schema-only preparation without stepping rows or evaluating SQL expressions.
The generator validates 1-255 selected, existing, unambiguous fields. Action and
parameter sources are refused. SELECT/view/query forms emit AllowAdditions/Edits/
Deletions=False and omit automatic New/Save command buttons; they execute actual
query records and Requery through the native form host. General query updateability,
parameter binding, original layout round trips and full Access parity remain open.

The common accessQuerySource reader now also resolves definitions from
main._la_objects, preserving a saved query name as RecordSource. Current SQL is
read on Requery; no artificial rows, persistent cursor or source snapshot cache
is introduced. Invalid/ambiguous owned definitions and owned/imported name conflicts
fail explicitly. Imported source lookup keeps its latest-artifact behavior and
uses the main catalog. This does not install arbitrary owned-query dependency
chains as temporary views; unsupported dependencies remain explicit failures.

An existing main-window flow prepares SELECT 42 field metadata, saves a form over
an actual owned COUNT query, opens it through the explorer, checks read-only values
and absent edit buttons, changes its real source table, calls Requery and observes
the updated count, then reopens the designer and checks the saved source/fields.
Existing physical-table/native subform/VBA/save cases still run. No new CTest target
or fixed-catalog item.
