# Native form OrderBy and OrderByOn

Bound native forms now expose OrderBy (String) and OrderByOn (Boolean).
OrderBy stores a comma-separated list of output fields with optional ASC/DESC;
changing an inactive definition does not reorder or save a record. OrderByOn
activates the definition. Direct table forms also accept verified table-qualified
field names. Imported OrderBy and OrderByOn/OrderByOnLoad properties initialize
the stored definition and requested initial ordering.

Active sorting executes through SQLite against the real source. Writable table
forms retain row identities/bookmarks and editability; the current physical row
is preserved when its ordinal changes. Query forms apply an outer ordering to
the existing compiled/filtered source, preserving the existing result subset and
proven type metadata. Query forms without physical row identities retain their
current ordinal rather than claiming bookmark-preserving reselection.

Changing active order saves through existing validation, then reloads and
updates CurrentRecord/calculated controls/Current events. Invalid or ambiguous
sort fields are rejected before saving. Failed reload restores the previous
sort state and position. Reentrant changes during record/filter/sort operations
are explicitly unsupported. Scalar property inputs are evaluated once, Null
reports 94. Sorting itself does not modify source values.

The existing form_current_record regression covers stored/inactive definitions,
actual VBA activation, physical identity across reordering, position text,
invalid-field state retention, disabling order, filtering, new rows, imported
initial order and readonly query sorting. No new test target was created.

This bounded grammar excludes sort expressions, numeric ordinals, qualified
query-output fields, ambiguous duplicate names and exact Jet collation/locale
identity. Some complex TOP/query wrappers retain existing explicit unsupported
boundaries. Report sorting API, design persistence/inheritance and complete
Windows event/error identity remain open. Public contracts:

- https://learn.microsoft.com/en-us/office/vba/api/access.form.orderby
- https://learn.microsoft.com/en-us/office/vba/api/access.form.orderbyon

Unqualified native form VBA resolves OrderBy/OrderByOn with normal field/control
precedence. DoCmd.RunCommand acCmdSortAscending (163) / acCmdSortDescending (164)
select the bound field under focus in the active document and apply its order.
Unbound/calculated/binary or missing focused fields reject the command (2046).
The same existing regression executes both commands and checks the real ordered
position and stored definition. Values were checked against the public AcCommand
enumeration: https://learn.microsoft.com/en-us/office/vba/api/access.accommand .
