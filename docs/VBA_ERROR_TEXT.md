# Native Error and Error$ functions

MS-VBAL6.1.2.3.1.15 defines Error/Error$ with an optional ErrorNumber. The native
VBA expression engine now supports parenthesized and bare calls, VBA-qualified
calls and named ErrorNumber. A supplied argument is Long-Let-coerced; values above
65535 raise6, negative Long values are accepted. Zero returns an empty String.
Unknown codes use the specified Application-defined or object-defined error.
fallback. Common native codes have a bounded invariant English description table;
this is not the complete Windows/Access locale or DAO error-message catalog.
Err.Raise with no retained/provided Description uses that same native table.
Existing retained/custom descriptions and other Raise arguments remain unchanged.

With no argument the actual project's last body-raised error number is used.
This separate history changes when a runtime statement raises an error, including
before propagation/debugger handling. Assigning Err.Number, Err.Clear, OnError,
Resume or explicit Exit clearing does not replace that history. Project reset
starts a new history through the actual project error environment. Full histories
for host binding/initialization failures and complete diagnostics remain F051.
No private VBA-facing history property or fake error object is introduced.

Local variables/arrays and actual project procedures retain unqualified-name
precedence. VBA.Error uses the intrinsic despite an unqualified Error array or
procedure; qualified intrinsic calls now avoid incorrectly selecting that local
array. In Access calculated expressions bare Error remains an actual field name;
Error(0) explicitly calls the function. No new standalone SQL builtin is registered.
Actual standard-module SQL functions use the real project error environment.

Shared existing core/GUI cases exercise text/zero/negative/unknown codes, rounding,
named/qualified/bare calls, retained custom Raise descriptions, default native
Raise text, history after Number assignment/Clear/OnError/ExitSub, real parameter
errors6/94/13/91, procedure/array shadowing and qualified escape. The actual SQL
function reads Err.Number907 while Error() describes the previously raised91,
proving the distinction through the SQL host. Existing core/GUI first passed2/2
23.15s (4.81/18.33). Source review then protected the Access field named Error;
only affected core was repeated, passing4.87s; app rebuilt. No new CTest ID,
Original rerun, Windows/full-suite/package claim from these checks. Last Original
118.10s and seven hashes unchanged precede CVErr/Error-text changes. Latest checked
DEB55096a7 precedes project Err/CVErr/text functions. Catalog456/386/70 unchanged;
complete F041/F043/F051/F118 remain open.
