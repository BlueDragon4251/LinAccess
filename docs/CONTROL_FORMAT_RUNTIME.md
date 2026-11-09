# Native calculated/report TextBox Format

SaveAsText TextBox Format strings are retained separately from input masks.
Form textboxes and report textboxes render supported masks through the
shared native formatter. A separate owned value cache retains the actual typed
Variant; a formatted string is never parsed back into Currency/Decimal/Date or
used as a SQL input. Calculated caches and domain/aggregate consumers retain their
original values. Bound report row changes replace their display leases; unbound
report assignments preserve the assigned type and use a separate display pass.

The native TextBox Format getter exposes the actual stored mask. Checked runtime
Format assignment is supported for form/report textboxes,
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

Writable form textboxes now show canonical raw editing text on focus and their
supported mask outside focus. Formatting uses blocked native text signals and
does not trigger Change or write SQL. The original typed value stays owned until
an actual edit invalidates it. Numeric editing uses enough invariant digits for
binary64 round trips; date/time editing retains the complete canonical timestamp.
Masks are never inverted to parse user input. Existing control validation, accepted
Value versus pending Text, Undo and SQL save rules remain authoritative.

Duplicate bindings copy actual unformatted values and apply their own masks.
Different values hidden by the same rounded display still dirty/save the record.
Native form creation captures proven direct field formats, preserving saved
independent overrides. Malformed format catalogs and metadata-read failures fail
explicitly. Changing Format during a pending edit/BeforeUpdate refuses3251; normal
accepted AfterUpdate formatting is supported. Scalar default values are resolved
once, with owner/writability rechecks before widget mutation; a callback-closing
form refuses2450. Calendar assignment uses the same resolved scalar.

Regional edit-mask parsing, arbitrary reentrant focus/IME behavior, combo/list/
other control formats, DecimalPlaces and full Windows compatibility remain open.
Form layout output still renders the actual current widgets, which can include a
focused editor's raw text; it is not an Access form-record printing oracle.
Unsupported formats refuse output. No new catalog ID or full requirement closure.


Validation: existing window target passed5.70s, including real report-designer
inheritance, Value types, actual report PDF extraction, failed-output preservation,
calculated masks, checked runtime mutation and unbound report Long/Currency masks.
The existing layout target then passed separately11.56s after replacing its
obsolete Currency-loop unsupported expectation with actual SQL iteration effects.
App built. These are separate passing invocations, not a fresh combined/full-suite
or Windows oracle claim. Four hundred fifty-six catalog IDs remain unchanged.


Writable-format validation: existing window/layout targets passed together2/2
(17.22s), application built. Real native DAO table/form/SQL and Qt input verify
focus changes without Change/Dirty/SQL effects, full-precision Double editing,
cancelled negative input, retained draft/accepted Value, pending-mask refusal,
Undo, accepted AfterUpdate masks, differently formatted duplicate bindings,
masked-equal changes persisted exactly, uppercase display with lowercase stored
text, and a hidden time component retained through saves. A scalar default is
read exactly once and a closing callback cannot write a stale form. Native form
wizard format inheritance and existing report/layout regressions are included.
