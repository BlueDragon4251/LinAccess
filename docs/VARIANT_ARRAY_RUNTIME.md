# Native Variant arrays and multidimensional values

`Array(...)` now creates a native immutable-backed Variant array. `Option Base`
0/1 determines its lower bound in the calling module. `LBound` and `UBound`
return Long values with a checked dimension argument. Empty arrays have upper bound
one less than their lower bound. Indexed reads and variable element assignments
check bounds (error 9), and Null indices fail with error 94.

`IsArray`, `VarType` (8204), `TypeName` (`Variant()`) and `IsObject` distinguish
arrays from objects. Arrays cannot implicitly become text or numbers (error 13).
Nested array values can be read through successive index operations.

Whole-array assignment and ByVal arguments retain independent value semantics.
Whole-array ByRef arguments propagate element replacements through existing VBA
alias machinery. Native Byte arrays also support checked indexing and replacement.
Existing core and DAO tests cover execution, copies, bounds, calls and actual SQL
persistence from a VBA loop on a synthetic in-memory database.

## Resizing Variant arrays

`ReDim a(upper)` and `ReDim a(lower To upper, ...)` allocate Empty
Variant elements with up to 60 dimensions. Omitted lower bounds follow Option Base; explicit lower bounds
may be negative. `ReDim Preserve` retains the overlapping prefix, initializes
new elements to Empty. Only the last upper bound can change; rank, lower bounds
and all earlier dimensions must remain unchanged (error 9 on rejected Preserve). Other value
copies retain their previous elements and bounds. `Erase a` releases this Variant
array and restores Empty; a subsequent ReDim Preserve can allocate it again.

This path targets variables declared As Variant (or implicit Variant variables).
Typed array declarations, non-Preserve String/Byte resizing, Byte-array Erase
and multiple declarations per statement remain unsupported. An explicit 1,000,000-element allocation limit
fails with 3251; this limit is a current implementation boundary, not Access parity.

## Remaining boundaries

This is not complete VBA array parity. Typed/fixed/dynamic array declarations,
full typed ReDim/Preserve/Erase semantics, full ParamArray semantics, full array For Each semantics,
full library/reference binding and native COM arrays remain open. Individual array elements
passed to a ByRef parameter fail explicitly with 3251 rather than silently losing
writes. Nested indexed assignment is not implemented. No Windows Access result
oracle was available for this segment.

## String lists and array enumeration

`Split` creates a zero-based String array (VarType 8200, TypeName String()),
independent of Option Base. It supports default/empty delimiters, limits, trailing
empty elements, explicit comparison modes and UTF-16 matching. `Join` consumes
one-dimensional String/Variant arrays and validates UTF-8; Null elements fail 94.
`Filter` selects/excludes text elements and returns a zero-based String array.
Non-text Filter elements currently fail explicitly with 3251. String element
replacement applies checked string conversion and retains array type/value-copy
semantics. ReDim Preserve now resizes existing String/Byte arrays held by Variant
variables, retaining type and prefix, defaulting new String cells to empty text
and Byte cells to zero. Lower bounds cannot change (Byte remains zero-based).
Value copies/ByVal inputs remain independent, while supported ByRef Variant
writeback receives the resized value. Non-Preserve typed resizing remains 3251.

Array For Each supports named local/module variables and direct Array/Split/Filter
results, with a Variant control variable. Named sources expose later element
changes during iteration, including element updates through a ByRef helper,
while independent copies remain independent. Empty
arrays skip the body. Direct source replacement, ReDim and Erase are rejected
while traversing the source. Allocation replacement detected at Next also fails
3251. Arbitrary function/property array sources, Byte array enumeration, full
cross-module alias/rebinding and mutation semantics remain open. This does not
claim Windows Access enumeration parity for those boundaries.

Variant and Byte arrays can also be returned by native VBA functions without
legacy scalar-text materialization; typed results retain their array values.
Arbitrary function results in For Each still have the documented source bounds.

## Qualified intrinsic array construction

VBA.Array now invokes the native intrinsic with lower bound zero independent of
Option Base; unqualified Array retains the module base. Direct For Each accepts
VBA.Array results. Existing local/host roots named VBA retain normal member lookup
priority. Named qualified Array arguments and complete VBA namespace/reference
binding remain unsupported. Existing core checks compare module base 1 versus
qualified base 0, empty bounds, direct enumeration and local-root lookup.

## Multidimensional Variant storage and DAO batches

Variant arrays now retain independent bounds for each dimension. Indexed reads
and assignments require the actual rank and check every subscript. Immutable
storage preserves whole-array copies and supported Variant ByVal/ByRef behavior.
For Each traverses every cell with the first dimension varying fastest; existing
replacement/resize guards apply. Join/Filter explicitly require rank one.
Dimension bounds, negative lower bounds, two-/three-dimensional assignment,
Option Base per omitted lower bound, last-dimension grow/shrink and rejected
Preserve changes execute in the existing core test target. No new targets.

DAO GetRows constructs genuine zero-based field/record matrices from actual
materialized SQL records; see [DAO_GETROWS_RUNTIME.md](DAO_GETROWS_RUNTIME.md).
Typed/fixed/dynamic array declarations and multidimensional typed arrays remain
open. This extends Variant values and ReDim, not the full declaration grammar
or native Windows SAFEARRAY/COM interoperability. Existing one-million-cell
runtime allocation bounds remain explicit.

Full desktop suite 48/48 (92.36 s), rebuilt GUI-free core 1/1 (1.50 s), Debian
payload/permissions/Build-ID/pinned SQLite and all four fixture hashes pass.
