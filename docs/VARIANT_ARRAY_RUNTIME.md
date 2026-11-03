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
Typed array declarations and multiple declarations per statement remain unsupported. An explicit 1,000,000-element allocation limit
fails with 3251; this limit is a current implementation boundary, not Access parity.

## Remaining boundaries

This is not complete VBA array parity. Typed/fixed/dynamic array declarations,
full typed declaration/Erase semantics, full ParamArray semantics, full array For Each semantics,
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
writeback receives the resized value. Non-Preserve resizing now retains the element type and initializes new storage
with the corresponding typed defaults.

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
Typed/fixed/dynamic array declarations, Class-specific/user-defined/fixed-length String
arrays and full Byte enumeration remain open. This extends Variant values and ReDim, not the full declaration grammar
or native Windows SAFEARRAY/COM interoperability. Existing one-million-cell
runtime allocation bounds remain explicit.

Full desktop suite 48/48 (92.36 s), rebuilt GUI-free core 1/1 (1.50 s), Debian
payload/permissions/Build-ID/pinned SQLite and all four fixture hashes pass.

## Typed ReDim inside Variant variables (2026-11-03)

Existing Variant variables accept ReDim bounds As Byte/Integer/Long/Single/
Double/Currency/Date/Boolean/String/Variant (and now Object, below). Storage retains the element type;
VarType includes the correct array flag/subtype and TypeName reports the actual
array type. Element assignment applies existing checked scalar conversion.
Default cells are typed zero, False, empty text or Date serial zero; Variant
cells remain Empty. Overflow/failed assignment preserves the old element.

Without As, resize retains the previous element type. Non-Preserve clears old
contents; Preserve retains values and rejects element-type changes (native 13)
as well as the existing rank/lower/earlier-dimension restrictions. Variant
container copies remain independent. Zero-based one-dimensional Byte values
retain compact byte storage and native byte interoperability.

This does not implement declaring a new typed array variable: typed ReDim on a
fresh undeclared name refuses 3251. Existing scalar variables cannot become
arrays through this route. Class-specific/user-defined/fixed-length String arrays and
full typed declaration grammar remain open. Join/Filter retain their one-
dimensional String/Variant scope; newly supported numeric typed array kinds
refuse 3251 on those operations. Exact Access error precedence remains open.

Existing core tests execute typed integer matrices/defaults/coercion/overflow,
Preserve/type guards, String reset, Byte/Currency/Date/Boolean values and metadata.
Full desktop 48/48 (89.81 s), rebuilt GUI-free core 1/1 (1.45 s), Debian package
and all four original fixture hashes pass. No new test targets.

## Object references and Variant Byte-array Erase (2026-11-03)

Set array(indices)=object/Nothing now assigns native references into Variant
array cells. Whole-array value copies share the referenced objects while retaining
independent array entries. Replacing one entry or erasing one container does not
clear a different copy's entry. Subsequent member access uses the existing native
object implementation and its normal connection/document lifetime checks.

Existing Variant variables also accept ReDim As Object, including dimensions,
Nothing defaults, VarType 8201/TypeName Object() and Preserve retaining reference
identity. Set accepts object references; scalar Set rejects 13 without replacing
the old entry. Plain value/default-member writes to an Object array remain 3251;
explicit object members use the normal binding. Class-specific arrays, typed
array declarations and general COM interoperability remain open.

Erase of a Byte array held in a Variant now releases that container to Empty,
including supported ByRef Variant writeback. Other value copies retain their
bytes and bounds. This does not implement fixed/dynamic typed declaration Erase.

The existing DAO target executes actual SQLite writes through Database objects
retrieved from multidimensional array cells, moves an actual stored Recordset,
checks copy/reference/Nothing/Preserve and failure behavior, and verifies SQL after
container Erase. Core verifies ByRef Byte Erase and an independent byte copy.
Full desktop suite 48/48 (92.45 s), rebuilt GUI-free core 1/1 (1.79 s), Debian
package and all four original fixture SHA-256 checks pass. No new test targets.
