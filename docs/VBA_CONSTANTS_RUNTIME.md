# Native scalar constants and constant array bounds (2026-11-04)

Local Const and module Const/Private Const/Public Const declarations bind real
immutable scalar values. Multiple comma-separated declarations, explicit
primitive/String/Date/Currency/Variant types and references to previously
declared constants use the native checked expression/coercion implementation.
Module constants default to private; public constants enter the existing public
module-symbol path. Local declarations shadow module bindings under the current
scope rules. Module values persist with their host state.

Constant expression evaluation uses a separate host without application/database/
filesystem context. A bounded token gate permits literals, known constants,
supported builtin constant names and arithmetic/logical operators. Functions,
mutable variables, object access and Is are refused; static declaration checking
does not execute FileLen or other runtime APIs. Scalar Null/Empty/Error/Object/
array constants are outside the supported path. Source-order dependencies and
the existing bounded expression/type/locale rules remain applicable.

Direct assignment/loop/resize targets that name constants refuse 3251 before the
procedure body executes. Supported scalar constants passed ByRef are temporary
values; callee changes do not overwrite the caller constant. These are bounded
native binding guarantees, not complete VBA compiler diagnostic/error precedence.
Full inferred constant type parity, qualified/cross-module compile-time linking,
class Public-Const visibility validation, cross-module default linking, conditional
#Const/Enum compilation and full module/class grammar remain open.

Fixed array declarations additionally accept integral constant expressions and
previous module/local constants in their dimensions, including parenthesized and
hexadecimal expressions. Bounds are folded into the statement representation
before runtime allocation and remain subject to the original rank/cell limits.
Mutable/dynamic expressions do not silently become fixed bounds; ReDim remains
the separate dynamic runtime operation. VARIANT_ARRAY_RUNTIME.md defines fixed
binding, Erase, ByRef and other array boundaries.

Native VBA character constants now include vbCr/vbLf/vbBack/vbFormFeed/
vbVerticalTab/vbNullChar/vbNullString. vbNewLine uses Windows CRLF for the Access
2003 target. vbNullString is represented as an empty native String value; this
does not claim distinct null-BSTR pointer or generic COM interoperability.

The existing core target checks actual typed local/module constants, expression
composition, comparison mode, module retention, pure rejection of runtime calls,
caller immutability under ByRef, pre-body write refusal and real constant-bound
array allocation. No new CTest target is introduced.

Primary VBA references (current documentation, not an Access 2003 oracle):
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/const-statement
https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/constant-expression-required

Checkpoint validation: full desktop 48/48 (95.27 s), freshly rebuilt GUI-free
core 1/1 (2.11 s), Debian payload/permissions/Build-ID/pinned SQLite and all four
original fixture SHA-256 checks pass.


## Folded optional parameter defaults (2026-11-04)

Optional scalar parameter defaults accept pure constant expressions and
previously declared constants from the same module. Their native values are
folded into parameter metadata during declaration parsing; omitted arguments
copy those values instead of evaluating an expression against the runtime host.
Named/positional overrides retain the existing binding path, and explicit Empty
remains distinct from omission. Existing Null/Empty and Collection Nothing
defaults retain their separate values. Each omitted ByRef default is a fresh
local value; changes do not alter the metadata or future calls.

Runtime calls, mutable variables and bare clock reads are refused during
constant/default evaluation. Cross-module/qualified constant linking, general
object defaults, full inferred type/compiler diagnostic precedence and the
existing optional/ParamArray boundaries remain open. Existing core integration
checks arithmetic/string/date defaults, named overrides, explicit Empty,
independent repeated ByRef defaults and pre-execution function-call refusal.
No new CTest target is introduced.

Default-expression checkpoint: full desktop 48/48 (92.51 s), freshly rebuilt
GUI-free 1/1 (1.84 s), Debian package and four original fixture hashes pass.
A subsequent test-only UTC constructor adjustment uses Qt-6.4-compatible API;
the query-model target passes 1/1 (0.04 s) on the existing Qt 6.8.2 environment.
This is not a separately executed Qt 6.4 distribution matrix.
