# Native declared versus Variant arithmetic

MS-VBAL 5.6.9.3 specifies that arithmetic overflow depends on the operand's
declared type as well as its value subtype. The native evaluator now carries a
transient declaration category through proven numeric expression bindings.
This category is separate from VarType and is not stored in database/clipboard
formats. It is established from actual local/argument/module-alias declarations,
actual function return declarations, native array element declarations, native
literals and explicit supported conversion calls. Runtime helper coercions alone
do not establish a source expression's declared type.

For supported primitive +, -, * and / paths, fixed Integer/Long/Single overflow
raises 6. Variant Integer results widen to Long when possible; Variant Long and
Single overflow can widen to Double. Unary minus distinguishes fixed signed
minima from Variant minima. Proven fixed Date range overflow raises 6; a Variant
Date can produce Double outside Date range. Existing checked Currency/Decimal
paths remain. CVar(Expression) returns the actual argument data value with a
Variant declaration category, including Object/Error/Empty/Null; it does not
invent a scalar-default read. Named Expression binding is supported. CVar's
Access Eval availability remains explicitly unvalidated/refused.

Array reads use the declared element type, including Variant ParamArray elements,
rather than a previous assignment's expression category. The first bundle found
this actual bug: a ParamArray slot containing Integer99 was incorrectly treated
as fixed Integer, so its existing *1000 property fixture overflowed. Rebinding at
the real array read fixed the core and real Qt property paths.

Generic host/control/member results and raw local dictionaries without a source
binding resolver are unbound. Cached values cannot silently become fixed source
declarations. Generic Object scalar/default paths also remain unbound. Proven
native runtime local/function/array bindings retain their actual category.

This is bounded support, not a complete AST or declared-type binder. Variant Byte
overflow, complete intrinsic constant/function/member result declarations,
logical/integer-division result provenance, mixed Currency arithmetic, full
suffix binding, all class/late-bound interfaces, IEEE exceptional assignment,
locale/Windows equivalence and full Access Eval remain open. Unbound contexts
retain their previous bounded behavior. F040/F041/F043/F122 stay open; fixed
catalog totals remain 456/386/70.

Application/core/DAO/query-model/GUI targets built. Initial DAO2.93s and query
model0.19s passed. After the array correction and host-boundary refinement,
core5.67s and GUI19.70s passed together (2/2,25.37s). After the final raw-dictionary
boundary refinement, only affected core was rerun and passed5.47s. Existing checks
cover actual local/function/ParamArray/array/property execution, CVar value and
object identity, static errors, Variant widening, Date limits and absence of
false overflow from cached host/helper values. An initial test helper attempted
to derive from a final fixture; that compile error was fixed before the successful
run. No new CTest targets/catalog IDs; no fresh full-suite or Windows proof.

The most recent checked DEB is source5df3cba2 at
/workspace/work/report-running-design/package/linaccess-0.7.0-Linux.deb,
with build-ID/SQLite3.54.1/MDBTools1.0.1/content checks. Its inventory46.35s and
separately secured original-path120.51s passed with100/78/22 and seven unchanged
hashes, retaining the missing macro-source boundary. That package and original
proof predate the later decimal-literal and declared-arithmetic changes.
