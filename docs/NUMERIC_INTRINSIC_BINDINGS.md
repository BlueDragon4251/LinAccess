# Native numeric intrinsic result declarations

MS-VBAL declares Abs, Fix, Int, Round and Sgn as Variant-returning functions;
Sqr, Exp, Log, Sin, Cos, Tan and Atn return Double. Choose, Switch and IIf also
return Variant. The existing Access Nz returns a Variant. The native evaluator
now preserves these declaration categories separately from the unchanged value
subtype. In particular, a selected CInt value does not make IIf/Choose/Switch/Nz
into a fixed Integer expression. Supported subsequent arithmetic can therefore
widen Variant Integer to Long rather than report a false fixed-type overflow.

The actual numeric work, precision, subtype checks, Currency/Decimal paths,
existing errors and eager argument evaluation remain in their existing helpers.
The declaration category comes from the function contract, never from an
incidental internal numeric coercion. Generic host/domain/report aggregate
results remain unbound, so cached source-expression categories cannot silently
become declarations of the host function. Access Eval retains its existing
observed numeric behavior and receives no new parity claim.

This extends BOUND_NUMERIC_ARITHMETIC.md. It does not complete all intrinsic or
member return bindings, Variant Byte overflow, mixed Currency/Decimal/Date
semantics, full AST/SQL typing, IEEE exceptional assignment or Windows parity.
F040/F041/F043/F122 and the fixed 456/386/70 catalog totals remain unchanged.

Application/core/DAO/query model targets built. The first existing bundled run
passed 3/3 in8.36 seconds (core5.18/DAO2.98/model0.20). Actual result TypeName and
arithmetic cover Variant-returning numeric and selection functions, Integer/Long
boundaries and Double-function Date overflow. Existing core error, precision and
Access reference checks remain in the same target. No new CTest/catalog IDs.
The checked5df3cba2 DEB predates the later literal and binding changes.

## Named arguments in native VBA

Native VBA binds `Number` for `Abs`, `Fix`, `Int`, `Sgn` and `Round`.
`Round` also binds the optional `NumDigitsAfterDecimal`. Names are
case-insensitive and may be reordered. `VBA.` qualification selects the
intrinsic when a procedure with the same name exists in the module. The
binding follows the formal function declarations in
[MS-VBAL](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74).

The core regression cases exercise omitted precision, a positional number with
named precision, reordered named arguments, Currency and Decimal results,
name and argument-count errors, scalar default access and procedure shadowing.
Verification covers the native runtime. Windows Access behavior and the broader
F040/F043/F118 requirements remain open.
