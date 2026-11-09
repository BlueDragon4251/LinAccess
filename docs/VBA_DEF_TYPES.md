# Native module Def-type directives

The native compiler resolves DefBool, DefByte, DefCur, DefDate, DefDbl, DefInt,
DefLng, DefObj, DefSng, DefStr and DefVar from the active module declaration
section. Supported letter specifications are ASCII letters, ascending/descending
ASCII ranges and comma-separated lists, case-insensitive. A-Z supplies the native
universal default, including identifiers beginning with a non-ASCII character.
Specific non-ASCII letter ranges remain an explicit3251 boundary. Overlapping
specifications fail5 before procedure effects. Directives in procedure sections,
malformed specifications and DefLngLng/DefLngPtr fail before execution. No 64-bit
Access-2003 types or undocumented default types are fabricated.

The same mapping is used by source-only procedure signatures, module and local
scalar/array declarations without As, function/getter result bindings, property
setter values, ordinary parameters and fresh/static implicit locals. Explicit As
wins; Const retains its existing expression-inferred type. Typed optional defaults,
ByRef compatibility and array element assignments use the existing type binder.
Option Explicit still requires variable declarations. Type suffixes, full
compile-time name binding and general classes/UDTs remain outside this scope.

Implicit native loop counters obtain the declared default before choosing their
runtime arithmetic; Currency counters use exact ticks. Array For Each controls
whose implicit type is scalar are refused by the existing pre-body target guard.
A new runtime ReDim target with an applicable Def mapping becomes a typed dynamic
array, with an explicit ReDim element type taking precedence. Existing bindings,
fixed-array/alias/allocation guards and Option Explicit restrictions are preserved;
this does not claim full VBA ReDim declaration hoisting or arbitrary declaration
and use-order equivalence.

Verification: application built. First GUI pass16.40s; the initial core block
found an empty DefInt specification reaching a body. That exact parser defect was
fixed, and implicit-loop/ReDim type binding completed. Final existing core/GUI
bundle2/2 passed17.68s (core4.33s), with no new CTest IDs. Checks cover all native
default types, source signatures, ordinary/optional/ByRef/property calls, arrays,
object identity, module/static/fresh state, type overrides, Const inference,
Currency steps, overflow, malformed/overlapping directives, conditional source,
universal Unicode names, descending ranges and Option Explicit. All seven actual
upload/fixture hashes remain unchanged. No additional Original/Windows/full-suite
or package proof. F041/F042/F043 and catalog456/386/70 remain open/unchanged.

Language basis: MS-VBAL5.2.2,5.2.3.1.5,5.2.3.2,5.3.1.4 and5.3.1.5. These are native
implementation checks, not new observations from Windows Access.
