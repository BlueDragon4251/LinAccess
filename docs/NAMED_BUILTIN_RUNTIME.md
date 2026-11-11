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


Native SQL Nz (2026-11-11)
-------------------------
Nz now has actual registered SQL callbacks for one or two arguments. Its omitted
SQL fallback is explicitly bound as a zero-length String before using the shared
native evaluator; VBA's Empty default remains unchanged. Evaluation remains eager,
including the explicit fallback expression. This is not SQLite COALESCE laziness.

Proven matching branch kinds (Date, Boolean, Single, integral and bounded Currency)
retain result metadata and actual DAO values/types. A literal NULL branch can
combine with the other proven kind. TEXT results remain without guessed Jet
Text/Memo declarations. Mixed/unknown branch kinds have unknown metadata; in
particular Nz(DateField) can contain Date-encoded text or the SQL empty String
and does not acquire a fake Date declaration. Full heterogeneous Variant identity
through nested SQL consumers is not implemented by this transport. Source-kind,
object/array, precision and locale guards remain; no new Windows oracle follows.
