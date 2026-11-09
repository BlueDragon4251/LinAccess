# Native named Property Let/Set indices

Actual native setter signatures now bind named index arguments for owning-module
PropertyLet/Set calls, public form/report members, public project-class members
and declared class default properties. Setter metadata comes from the real Let/
Set accessor, including getterless setters; no getter is executed to discover it.
The final RHS value parameter is excluded from the named index namespace.

Index expressions are evaluated once in supplied order; names reorder the values
and actual variable/array-cell references together. Omitted optional indices retain
Missing/default binding. Typed object indices preserve references. The existing
final-value ByVal rule also applies when the source writes ByRef. Unknown/duplicate
index names fail448, absent required names449 and invalid positional-after-named
or excessive arguments450 before invoking a setter. This bounded runtime binding
is not a full VBA static compiler or assignment evaluation-order oracle.

Owning-module ByRef indices and supported ParamArray aliases retain their original
addresses. External object facades still require explicit ByVal indices; there is
no guessed cross-object ByRef writeback. Named ParamArray setters, native objects
without actual setter metadata, named array subscripts and getter/default-member
forwarding without a real setter remain explicitly outside this path. Full COM/
reference/visibility, evaluation/reentrancy and Windows semantics remain F043.

Existing core source checks reordered typed variable/array/temporary indices,
optional getterless setters, Set/object identity, object-valued indices, final
value isolation, ParamArray forwarding and invalid names without setter/getter/
RHS calls. Real project class/default and real form facades cover the new native
metadata providers and preserve their external ByRef guard. Getter counters remain
zero. Core passed5.02s in the first bundle. GUI initially waited on a QMessageBox:
the introduced NamedWriter class was not cleaned up, so a later TEMP-catalog
fixture correctly reached the module-shadow guard. The class is now removed at
its own scope boundary; the production catalog guard is retained. Final affected
GUI passed18.54s. Only GUI was repeated after its isolation/diagnostic adjustments.

The GUI target keeps a precise20-second watchdog and reports only the modal C++
class on timeout. Temporary message-content logging was removed. Generic language
negative matrices, closed object-array computation, Enum computation and owning
named-setter matrices remain in core rather than being duplicated in GUI. Real Qt
class/form/array and SQL Enum/provider checks remain in GUI. No new CTest IDs or
inflated timeout; no new Original/Windows/full-suite/package claim. Catalog456/386/70
unchanged. Last Original118.62s/DAO2.90s and seven hashes unchanged precede this
named setter scope. Latest checked DEB9affc5a precedes ParamArray/named setters:
/workspace/work/error-text-functions/package/linaccess-0.7.0-Linux.deb.
