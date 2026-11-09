# Native calculated/report TextBox Format

SaveAsText TextBox Format strings are retained separately from input masks.
Calculated form textboxes and report textboxes render supported masks through the
shared native formatter. A separate owned value cache retains the actual typed
Variant; a formatted string is never parsed back into Currency/Decimal/Date or
used as a SQL input. Calculated caches and domain/aggregate consumers retain their
original values. Bound report row changes replace their display leases; unbound
report assignments preserve the assigned type and use a separate display pass.

The native TextBox Format getter exposes the actual stored mask. Checked runtime
Format assignment is supported for calculated form controls and report controls,
including clear-to-raw display. Null94, non-String13, invalid masks5, length/Unicode,
closed owners and mutation during report rendering are checked. Failed validation
retains the previous mask. Formatter errors affect presentation separately from
calculation errors, so Value remains the actual typed value; PDF/print refuse
format-error controls instead of printing the visible error placeholder.

Native report creation captures proven direct-source field Format metadata beside
captions, retaining existing saved overrides and recapturing on source changes.
The saved report emits real Format properties for its bound controls. Native table
formats can therefore reach actual report previews and paginated PDF output.
Complex query/view/alias lineage and complete Access rendering/events are open.

Writable form TextBox focus/edit-mask behavior is still unsupported here. Such
source masks are retained and have an explicit3251 tooltip; runtime mutation
refuses3251 rather than reinterpreting formatted strings as edited field data.
Combo/list/other control formatting, DecimalPlaces, full mask grammar and Windows
Access regional/pixel/printing equivalence remain open. This adds no new catalog
ID and does not certify any of the72 broad full requirements.


Validation: existing window target passed5.70s, including real report-designer
inheritance, Value types, actual report PDF extraction, failed-output preservation,
calculated masks, checked runtime mutation and unbound report Long/Currency masks.
The existing layout target then passed separately11.56s after replacing its
obsolete Currency-loop unsupported expectation with actual SQL iteration effects.
App built. These are separate passing invocations, not a fresh combined/full-suite
or Windows oracle claim. Four hundred fifty-six catalog IDs remain unchanged.
