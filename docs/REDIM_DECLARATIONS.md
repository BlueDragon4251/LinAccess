# ReDim declaration binding

MS-VBAL5.4.3.3 requires an unmatched ReDim name to declare a resizable local array.
The native runtime now binds these names for the whole procedure, including
OptionExplicit and skipped branches, without evaluating bounds before execution.
Explicit element types or real module Def directives determine the element type;
otherwise it is Variant. Existing parameters, locals, module/public fields and
constants retain their original bindings. Unresolved instance/procedure targets
are refused rather than silently shadowed by invented locals.

Static procedures retain these declared array cells in their actual procedure
state. Existing typed-array assignment, exact ByRef, return, Erase and Preserve
rules apply; a later ReDim cannot change a declared array's element type. Existing
Variant variables retain their prior ability to hold differently typed arrays.
No member/With ReDim target support, AsNew arrays, type-name suffix declarations,
full external reference resolution or complete compiler parity is claimed.

Shared technical sources exercise pre-statement type inspection, declaration in
an inactive branch, bounds evaluated once, real ByRef mutation, returned native
object arrays, retained module arrays and Static calls, and rejected type changes
without changing retained bounds. The former negative test for an undeclared
ReDim AsLong now expects success. A Static fixture intentionally needing a Variant
explicitly declares that Variant rather than relying on the previous missing
ReDim declaration binding.

App built; existing core/GUI2/2 passed22.96s (4.74/18.22). Common native-object and
ReDim integration checked together: DAO2.71s and original event path119.91s,
2/2 passed122.62s. Original known missing-macro limits remain, no initialized-sale
claim. All seven original/upload fixture hashes remain unchanged. No new CTest ID,
Windows oracle, full-suite or package refresh. Catalog456/386/70 unchanged;
complete F040/F042/F043 remain open.
