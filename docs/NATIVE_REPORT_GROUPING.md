# Native report groups, sorting and scoped totals

The actual report designer offers a grouping/sorting dialog for up to ten unique
source fields. Users can add/remove/reorder levels, select ascending/descending,
choose a group header and enter an optional group-footer calculation. Fields
come from the prepared real SELECT source; no rows are sampled to invent types.
The persisted groups array generates explicit BreakLevel metadata and separate
BreakHeader/BreakFooter sections. Reopening edits the real saved definition.

The layout decoder retains distinct group sections and controls with their actual
level association. Printing requires explicit/proven level links and all seven
source group properties. Supported GroupOn is EachValue0, GroupInterval1 and
KeepTogether0. Unknown/missing/advanced source properties remain guarded. Native
authoring emits that complete supported contract; absent imported properties are
not replaced with guessed Access defaults. Existing grouped-source limitations
are waived only after that complete plan is validated; unrelated limits remain.

Printing stably sorts actual source rows by the configured key hierarchy, keeps
original source order within equal groups and emits outer-to-inner headers /
inner-to-outer footers. NULL is a real group key. Keys use native scalar semantics
(including exact Currency units and chronological native Date interpretation);
text follows the existing native case-insensitive comparison, not a claimed full
Jet/locale collation. Binary/mixed/unsupported keys and bounded memory excesses
refuse output. Group sections remain intact; advanced grouping intervals,
repeated headers, whole-group KeepTogether and growing group bands are outside
this scope. Existing growing/split detail and report page-band/output settings
continue to apply.

Each group band carries its actual half-open row scope. Count(*), field and
compound aggregates use that scope with distinct cache keys; report bands use
the entire report. Actual Currency sums retain scaled units. Sorting is temporary:
source rows/identities, recordset reference, aggregate scope, current record and
control layout are restored after every render/failure. Direct report navigation,
runtime GroupLevel mutation, every Access event and a Windows/original grouped
print comparison remain broader open requirements.

The existing GUI/report bundle checks real grouped authoring/reopen/PDF, nested
mixed-direction keys, NULL, stable row markers, exact group/report Currency sums,
Count scopes, First/Last print order and restoration of the original source order,
a chronological Date boundary, advanced/missing metadata refusal and retained
prior PDF. App/GUI/report targets built. The existing GUI check passed in12.68s for real
group creation/reopen. The targeted report/layout check passed1/1 in9.99s after
its fixture used supported SQL concatenation and supplied the actual packet
record source, just as the application does. Ungrouped aggregate restore caching
retains its previous behavior. Existing original toolbar-handler cases now read
the previously omitted group controls: without a database, all three correctly
refuse layout PDF because calculated Text values carry numeric format13. Their
source-unavailable context, exact known mask hash, type/error and prior-file
retention are explicitly checked. This replaces a formerly incomplete positive
layout observation; it is not evidence of working original report output. No new
private VBA invocation or business fixture replacement was added. No new CTest/catalog identity; F050/F080/F085/F087/F105/F130
remain broad requirements. Latest checked DEB/Original/inventory is code299b946f,
before StrConv/RNG/page-band/group changes. No new Windows parity assertion.

Bounded native repeated group headers and live Section access are now documented in [NATIVE_REPORT_REPEAT_HEADER.md](NATIVE_REPORT_REPEAT_HEADER.md); broader grouping/Windows bounds remain.
