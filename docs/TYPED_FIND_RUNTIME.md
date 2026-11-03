# Typed native Find

The native table/form Find dialogs and DoCmd.FindRecord search current supported
fields with actual checked value subtypes. Text retains literal part/whole/start
and case modes. Numeric, Currency, Date and Boolean searches use whole-field
matching, real SQL-backed rows, existing direction/origin semantics and checked
record navigation. No hidden replacement form or stringified numeric comparison
is used.

## Values and boundaries

- Number: native integer/Double fields compare actual numeric values exactly.
  Adjacent binary64 values stay distinct. Unsafe integer storage and malformed
  numeric fields retain the existing checked field-conversion boundaries.
- Currency: native CURRENCY_REAL values compare exact signed four-decimal units.
  The VBA call uses a Currency FindWhat (for example CCur); a numeric Variant is
  not silently rounded/coerced into that field subtype. Existing lossless REAL
  storage limits remain unchanged.
- Date: declared DATE/DATETIME fields compare checked canonical native Date
  values at the existing second-resolution storage precision. Negative OLE Date
  aliases normalize through the real Date storage pipeline. Subsecond search
  outside that storage contract is explicitly refused (3251).
- Boolean: a Boolean FindWhat compares actual Boolean values. SQL NULL is skipped,
  never treated as False/zero. FindNext retains this typed request.

VBA value subtypes cross the real Qt method bridge in checked QVariant values;
Currency uses a signed 64-bit unit value. No JSON/floating-point Currency
intermediate or current-control display rounding is involved in comparison.
Unsupported cross-subtype coercions, nontext partial matching, all-field API
search, formatted/locale/wildcard/empty search and complete Access error/selection
appearance equivalence remain open. A typed search is a bounded native contract,
not a general Jet/ACE search-parity assertion.

## Native dialogs

The field dropdown includes supported visible text/numeric/Currency/date/Boolean
fields. An individual typed field selects Whole Field, disables case and partial
options and shows its input format. Integers use whole numbers, other numbers
and Currency use a decimal point, dates/times use the native ISO input forms,
and Boolean accepts True/False or numeric values through existing Boolean
coercion. Invalid input reports a checked error without navigation or SQL writes.
The combined scope remains All Text Fields; binary/unsupported fields are not
presented as searchable. Numeric-only tables can open Find and repeat a prior
typed search through the Edit menu.

The actual manual-submit/validation policy remains in force: opening Find alone
does not save; a sole current existing record may be saved before table search;
unrelated table drafts stay explicitly protected. Form BeforeUpdate cancellation
still aborts programmatic FindRecord with 2501. Readonly table/control search is
available without data mutation. Modal, page ownership and dialog lifetime guards
are unchanged.

## Verification

Existing main-window/Currency/form-Find targets pass 3/3 (2.79 s). Interpreted calls
and actual Qt dialogs distinguish neighboring Double values, exact Currency
units, a leap-day timestamp and a negative OLE Date alias. Boolean search/FindNext
skips SQL NULL, readonly searches work, incompatible type/mode/subsecond requests
are refused, invalid date input preserves position, and actual SQL values/nulls/
binary payloads stay unchanged. No new test targets.

Full rebuilt desktop suite passes 48/48 (97.33 s), GUI-free core 1/1 (1.86 s),
Debian payload/permissions/Build-ID/pinned SQLite checks and all four original
fixture hashes pass. No extra test targets. This also covers the preceding
search-direction extension and the typed native invocation/state changes.
