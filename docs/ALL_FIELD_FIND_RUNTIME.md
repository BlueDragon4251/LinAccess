# Native all-field FindRecord

OnlyCurrentField:=acAll (0) supports native raw-value FindRecord across the
matching supported field kinds. String needles search text fields; native numeric,
Currency, Date and Boolean needles use the existing checked typed equality on
matching field kinds. Typed matching still requires whole-field mode. This does
not infer cross-type coercions from formatted strings or unknown result columns.

Native tables and readonly query results use their existing all-column scanner
and select the matching cell. Direct-table forms use their bounded record data
and matching source field types; the focused control need not have the sought
value's type. Form FindNext retains the all-field scope. Formatted all-field
String searches remain a separate display route and keep their existing bound
Textbox/control-mask rules. Wildcard String needles use the same text/display
scope and matcher as current-field searches.

Existing activation, modal, validation/save, null/binary exclusion, provenance,
row/cell bounds and native per-record repeat ordering remain. Full Access
all-field coercion/column traversal/focus/event behavior, locale equivalence and
unknown query result types remain open under F373/F071/F130/F050. No new Windows
Access observation or broad parity claim is introduced.
