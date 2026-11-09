# Native ParamArray aliases and active array locks

MS-VBAL5.3.1.5 defines ParamArray elements as implicit ByRef Variant positional
parameters;5.3.1.11 retains actual variable types when writing those aliases.
When actual VBA argument bindings are available, native packs now bind scalar/
object variables and supported named array elements to their original sources.
Literal/parenthesized/value arguments retain local pack slots. Let/Set element
writes, typed source coercion, immediate observation and ordinary ByRef forwarding
operate through actual sources. Duplicate arguments alias the same variable;
forwarding a pack element reuses its original binding rather than a stale second
array cell. Actual temporary-cell lifetimes and parent alias refresh are retained.

Alias metadata identifies dependent slots without inventing VBA variables. Failed
source conversion/writes restore the affected slot without replacing the pack.
Ordinary aliases borrow their existing values, avoiding a new full Byte()/Memo
copy solely for change comparison. Active ByRef element roots reject array
replacement with native error10 (MS-VBAL5.4.3.3), retaining the actual protected
storage and releasing the lock when the real invocation ends.

Array-valued source-slot writes, unresolved compound addresses, whole-pack
replacement/resize/Erase, full argument expansion/holes/named ParamArray binding
and complete COM/class facade binder semantics remain open. Facades/direct value
APIs lacking real argument classification retain the explicit3251 write guard;
they do not pretend that changing a copy changes a caller variable. Readonly pack
uses and existing object-reference behavior remain. Full live enumeration/array/
Missing/lifecycle behavior and Windows Access equivalence remain F043.

Existing shared core/native-project fixtures exercise real typed variable writes,
parenthesized/literal local slots, duplicate aliases, nested ByRef forwarding,
immediate module reads, object Set replacement/identity, Long/Byte failures with
unchanged sources, explicit array-valued write refusal, actual source array-cell
writes and lock release, empty pack bounds9 and unbound value-API write refusal.
Existing property fixtures now allow local literal-slot mutation through an
owning VBA call. Existing active-cell protection now expects10 instead of3251.

App built. Initial existing bundle: GUI18.94s/DAO2.78s/Original118.62s passed;
core exposed an older literal-pack fixture still expecting mutation refusal.
Its expected local value/error are updated to bad|b|c/0. The ordinary alias read
was also changed to borrow values rather than copying Byte()/Memo. Only affected
core/DAO were repeated, passing2/2 in7.97s (5.06/2.90); GUI/Original were not
needlessly rerun after that internal borrowing/fixture adjustment. Seven original/
upload hashes remain unchanged. Original still has its known missing-macro limit.
No new CTest ID, fresh full suite or Windows oracle. Catalog456/386/70 unchanged.

Latest checked package is9affc5a, before this ParamArray/lock change:
/workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb.
Build-ID/payload/rights/no-fixtures, MDBTools1.0.1 and SQLite3.54.1 JOIN/NULL/counter
checks passed. Its DAO/Original integration passed2/2 in123.02s (2.76/120.25).
