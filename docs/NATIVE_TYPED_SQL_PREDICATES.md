# Proven native SQL membership and range predicates

Native IN/NOT IN and BETWEEN/NOT BETWEEN decode proven Date, Boolean, Single,
Currency and Decimal operands before comparing them. This keeps Boolean -1/0,
native Single rounding, checked Currency values and Date serial ordering rather
than comparing their raw SQLite transport. The existing Decimal predicate path
and exact key codec remain in use. Each SQL operand is evaluated once.

Results retain proven BOOLEAN metadata and native DAO Boolean values. NULL
membership candidates preserve unknown results unless another candidate matches;
a NULL tested value stays unknown. BETWEEN includes both ends and returns NULL
when any operand is NULL, including cases where one non-NULL inequality would
already be false. NOT preserves NULL. Date operands require Date or NULL; numeric
predicates require supported proven numeric kinds or NULL inside the new adapter.
Date/numeric mixtures with proven descriptors refuse3251. Initial compilation
defers binding until schema resolution; other insufficiently proven non-Decimal
and plain text/untyped SQLite predicates retain their existing path. Decimal
continues to refuse unproven contexts rather than exposing its binary codec.

Operand limits and reversed-bound refusal remain the existing Decimal scope.
This does not establish the Access behavior of reversed ranges, arbitrary Date/
text/numeric coercion, complete collation, saved-query alias lineage or all Jet
expression evaluation rules. General non-Decimal projection aliases are not newly
bound by this change. No Windows execution oracle or full Access parity is claimed.

Microsoft documents inclusive BETWEEN bounds and NULL propagation:
https://learn.microsoft.com/en-us/office/vba/access/concepts/miscellaneous/between-and-operator
https://learn.microsoft.com/en-us/office/vba/access/concepts/miscellaneous/in-operator-microsoft-access-sql

F014/F041/F067/F068/F071 remain open; fixed456/386/70 unchanged.

App built; existing query model0.41s and final affected DAO1/1 in5.32s passed.
Initial source-less compilation now defers binding until schema resolution.
No new CTest target/catalog item.
