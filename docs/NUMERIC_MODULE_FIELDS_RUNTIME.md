# Integer and Long module fields

Module-level Dim/Private/Public/Global declarations now retain As Integer and
As Long fields with numeric zero defaults. Untyped declarations remain Variant/
Empty; visibility and per-form/project module-state ownership retain their
existing scopes. Only each explicitly typed declarator receives that type.
Arrays and other unsupported field types continue to reject explicitly.

Assignments use the existing native module-alias synchronization after each
instruction and explicit form-member assignment. Numeric conversion rounds to
nearest with ties to even, then checks signed 16-bit Integer or 32-bit Long
bounds. Overflow uses 6, Null 94 and incompatible values 13. Failed conversion
preserves the previous stored value; On Error Resume Next sees the error and
execution can continue. Existing Boolean/String/Variant behavior is preserved.
Long values such as 40000 are not truncated to Integer range.

Core and actual GUI/VBA/SQL tests cover defaults, visibility, declarator/shadow
rules, rounding and both bounds, overflow recovery and unchanged values, Empty/
Boolean/Currency numeric inputs and unsupported arrays. Native forms verify
independent private state, public field access/coercion, shared standard-module
Long state across forms, database isolation, close/reopen/project reset and real
SQL insertion of the resulting module value. Targeted core/main-window/layout/
original-heavy set: 4/4 (42.34 s). Full validation is in PROGRESS.md.

This adds bounded field storage and assignment conversion, not full VBA numeric
type identity. Values still use the existing Number representation; declared
VarType/TypeName/Len identity, typed intermediate arithmetic/promotion/overflow,
all ByRef type contracts, arrays, Single/Double/Byte/Date module fields, locale
coercion and classes remain open. No additional original business execution
coverage is claimed.

Procedure-local/parameter/result checking now also has a bounded implementation
in NUMERIC_PROCEDURE_VALUES_RUNTIME.md, with its explicit ByRef/loop/type limits.

Subsequent Integer/Long storage identity, CInt and tagged CLng are documented in
INTEGER_VALUE_IDENTITY.md. Full literal/arithmetic/Variant-storage/Len semantics
remain open; the earlier milestone's generic numeric representation is historical.
