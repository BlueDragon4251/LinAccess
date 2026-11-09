# Native ParamArray argument evaluation

The final parameter in a native VBA Sub/Function can be declared ParamArray
values() As Variant (As Variant may be omitted). Additional evaluated arguments
are packed into a real zero-based Variant array, independent of Option Base.
Required preceding scalar arguments retain their ordinary checked binding.
Zero additional arguments yield an empty array with bounds 0/-1.

Procedures can inspect bounds, index values, enumerate the array and pass it to
Join or to a ByVal Variant helper. Arrays passed as one argument remain nested
array values. Existing core and DAO execution targets exercise multiple/empty
argument lists, typed function results, ByVal forwarding and actual insertion
through a ParamArray function combined with an Optional-default helper.

## Remaining boundaries

This first implementation evaluates ParamArray values; it does not claim full
VBA argument-slot/reference mutation semantics. Slot replacement, resize, whole
array reassignment and ByRef forwarding are rejected explicitly with 3251.
Object argument references retain their existing native object behavior.
Combined Optional/ParamArray declarations, named arguments, argument holes,
array expansion and full Windows binder behavior remain open. ParamArray must
be last; unsupported declarations fail during procedure validation.


The subsequent real variable-binding path is documented in PARAMARRAY_ALIASES.md.
It supersedes the initial blanket slot-write refusal for actual VBA calls with
resolved argument bindings, preserving explicit guards for unbound facades and
unsupported array-valued/compound aliases. It does not complete F043.
