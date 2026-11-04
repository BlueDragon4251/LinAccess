# Native owned report footer calculation

The owned-report editor now saves an optional calculated report footer and its
ForceNewPage choice in the existing main._la_objects definition:
footer_expression (a single-line expression beginning with =) and
footer_force_new_page (integer 0..3). Invalid property types/ranges, embedded
control characters and a nonzero break without a footer expression are refused.
The existing report definition, VBA, SQL and separately saved page geometry are
preserved by the editor's current-definition merge.

The native layout generator emits an actual FormFooter with a calculated textbox
using a collision-safe name based on native_report_summary. Loaded SQL records,
report aggregates, runtime validation and section pagination execute through the
normal native report host. Footer output is fixed height; this is not a general
report designer, grouping/CanShrink/CanGrow footer implementation or original MDB
round trip. Full Access expression/error/page-break oracle equivalence remains
open. An empty expression with break zero preserves the previous no-footer layout.

An existing real main-window flow saves a report using Sum([id]) over three SQL
records, opens it through the object explorer, observes native Value=6 and two
actual preview pages with a break before the footer, exports PDF, then reopens the
editor and checks that both properties persisted. Previous report SQL/VBA/page
settings and event-lifetime cases still run. No new CTest target or catalog item.
