# Native one-dimensional Variant arrays

`Array(...)` now creates a native immutable-backed Variant array. `Option Base`
0/1 determines its lower bound in the calling module. `LBound` and `UBound`
return Long values; dimension 1 is supported. Empty arrays have upper bound
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

One-dimensional `ReDim a(upper)` and `ReDim a(lower To upper)` allocate Empty
Variant elements. Omitted lower bounds follow Option Base; explicit lower bounds
may be negative. `ReDim Preserve` retains the overlapping prefix, initializes
new elements to Empty and rejects lower-bound changes with error 9. Other value
copies retain their previous elements and bounds. `Erase a` releases this Variant
array and restores Empty; a subsequent ReDim Preserve can allocate it again.

This path targets variables declared As Variant (or implicit Variant variables).
Typed array declarations, Byte-array resize/erase and multiple declarations per
statement remain unsupported. An explicit 1,000,000-element allocation limit
fails with 3251; this limit is a current implementation boundary, not Access parity.

## Remaining boundaries

This is not complete VBA array parity. Typed/fixed/dynamic array declarations,
full typed ReDim/Preserve/Erase semantics, multidimensional arrays, ParamArray, full array For Each semantics,
qualified VBA.Array and native COM arrays remain open. Individual array elements
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
semantics. String-array ReDim remains explicitly unsupported (3251).

Array For Each supports named local/module variables and direct Array/Split/Filter
results, with a Variant control variable. Named sources expose later element
changes during iteration, including element updates through a ByRef helper,
while independent copies remain independent. Empty
arrays skip the body. Direct source replacement, ReDim and Erase are rejected
while traversing the source. Allocation replacement detected at Next also fails
3251. Arbitrary function/property array sources, Byte array enumeration, full
cross-module alias/rebinding and mutation semantics remain open. This does not
claim Windows Access enumeration parity for those boundaries.
