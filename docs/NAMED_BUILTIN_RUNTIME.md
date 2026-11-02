# Named native builtin arguments and Nz context

Selected existing native builtins now bind documented parameter names before
executing their real C++ implementation: MsgBox, InputBox/InputBox$, Replace,
Split, Join, Filter, InStr/InStrRev/StrComp/StrReverse, RGB/QBColor, DateDiff,
Weekday, Left/Right/Mid (including $ forms), Nz, IsMissing, DSum and DCount.
Written expression evaluation order is retained; unknown names fail 448,
duplicate/missing/positional-after-named bindings fail 450. Optional slots keep
the function's existing defaults and unsupported options. InStr retains its
four-slot shape when named arguments are used to avoid overload ambiguity.

Other builtin signatures, Access Eval named syntax, full Windows parameter
metadata and compile-time binding remain open. Existing method/procedure named
support is documented in NAMED_ARGUMENT_RUNTIME.md.

## Nz correction

The Microsoft Access reference specifies that Nz(Null) without ValueIfNull in
VBA returns Empty; numeric and string context then decide the value. The native
VBA runtime now preserves that subtype rather than forcing an empty String.
Examples Nz(Null)+2 and Nz(Null)&2 work accordingly. The Access query-expression
path keeps its documented zero-length String default. Explicit fallback values
retain their subtype. This does not establish all Object/SCODE/coercion parity.

Reference: https://learn.microsoft.com/en-us/office/vba/api/access.application.nz

Existing core and DAO targets cover named/reordered calls, overload/default/error
cases, the documented Nz examples and actual SQL persistence through named Join.
No new CTest target is added.
