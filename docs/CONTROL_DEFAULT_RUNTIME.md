# Native form-control literal defaults

## Implemented scope

SaveAsText DefaultValue is retained as an expression string, including declared
control-type templates and explicit instance overrides/clearing. Native form
TextBox, ComboBox/ListBox adapters, standalone CheckBox/OptionButton and OptionGroup
objects expose DefaultValue. Grouped individual choices, other control kinds and
report controls reject the property with 438. Getters reject arguments with 450;
Null assignment raises 94, invalid scalar types 13 and expressions over 255 UTF-16
code units 5 before mutation. Scalar defaults resolve once. Metadata assignment
changes future initialization, not the current Value, SQL row or Dirty state.

This stage evaluates only finite numeric, quoted string (VBA doubled quotes),
Null and Boolean literals, with optional leading equals. Unbound controls initialize
once per instance. Bound controls initialize on an eligible new row of a writable
direct table with additions enabled. Existing rows retain database values. Shared
bound fields receive one value across their mirrors; differing default strings
for one field reject with 3251 rather than guessing a priority. New-row load/Undo
reinitializes from current instance metadata. Defaults create no Change/insert/
Dirty events and no stored record until actual data changes. Locked text fields
can receive defaults while remaining locked for user input.

AutoNumber columns stay with the existing generation adapter and are not default-
evaluated; opening/closing an untouched defaulted row does not consume IDs or insert
an empty row. An explicit control literal overrides a known SQLite table default
when that control supplies the inserted field. General table-default fallback for
blank control settings and import of Jet defaults remain unsupported.

Computed expressions are deliberately not evaluated by this literal stage. A
failed/overflowing/default conflict records a numeric error (3251/6 or conversion
error), shows a field placeholder/tooltip and raises on Value access. It cannot
silently bind NULL during an attempted record save. User input or an explicit
valid Value assignment can override the failed default for that row. Updating
DefaultValue alone remains a metadata change. Original expression text is never
included in these diagnostics.

## Evidence

Actual Qt/VBA/SQLite tests verify template inheritance/clearing, metadata versus
existing Value, a native bound VBA button changing future defaults, numeric/text/
Unicode/empty-string/Null/Boolean values, mirror consistency, unbound persistence
across navigation, locked fields, exact SQL empty-versus-NULL storage and an
explicit control override of a SQLite column default. SQL traces prove no insert/
Dirty/Change events during initialization and normal events during later user
input. Defaulted untouched rows create no records or sequence advances. Actual
save generates the next ID and persists the entered/defaulted fields.

Unsupported computed values prevent a partial/defaulted record insert before SQL;
manual user input fixes the field and permits saving. Numeric overflow, Error
Variant override, UTF-16 length, scalar-once conversion, unsupported controls,
reports and conflicting mirror defaults are tested. Targeted layout/original-heavy
regressions pass (39.04 s).

Across hash-verified original forms, all 34 nonempty default metadata entries are
retained by native getters; 15 unbound literal initializations match a separate
native Access-expression-parser evaluation. No unsupported adapters were counted
for these fixture entries. These probes use an empty database and no CodeBehind
handlers or original table rows. Bound original new-record execution, original
field typing and business-handler coverage are not inferred from these counts.

Contract reference: [TextBox.DefaultValue](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.defaultvalue)
documents new-record initialization, the 255-character limit, control precedence
and earlier-tab-order references. A Windows Access 2003 runtime oracle remains
unavailable; the native parser cross-check is not that oracle.

## Remaining scope

Computed/date/field/Forms/UDF defaults, full tab-order dependency semantics,
conflicting-default priority, Jet table defaults, bound blank-setting fallback,
AutoNumber allocation/display timing, OLE/complex fields, exact Currency/Date/
Decimal and validation/locale fidelity, complete new-row OldValue/error ordering,
OnOpen/Load timing and durable design-property writing remain open. This literal
stage does not establish complete Access default or record-lifecycle parity.

Validation: host 21/21 (87.53 s), Debian 12 21/21 (67.46 s), Ubuntu 24.04
21/21 (62.39 s), GUI-disabled core 1/1 (1.05 s). Fresh host/distro DEBs pass
exact payload/backend/permissions/fixture-exclusion checks; isolated installed
Debian/Ubuntu packages pass backend/counter execution and GUI startup. All four
original fixture hashes pass. Original counts remain 34 metadata / 15 unbound
literal initializations / zero unsupported adapters for these entries.
