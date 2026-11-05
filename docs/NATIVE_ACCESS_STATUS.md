# Native Access status messages and progress meter

`Application.SysCmd` and unqualified `SysCmd` now execute the status actions
acSysCmdInitMeter1, UpdateMeter2, RemoveMeter3, SetStatus4 and ClearStatus5.
Both function expressions and statement calls use the real native Application
object; named Action/Argument2/Argument3 use the normal binder. Local and external
VBA procedure declarations retain priority for statement calls; local functions
retain expression-call priority. Syntax validation does not invoke GUI actions.
Calculated-expression/SQL SysCmd is not added.

The real QMainWindow status bar displays the message and QProgressBar; a standalone
native document obtains an actual status bar in its layout. Maximum200/update100
produces an actual half-filled meter. SetStatus removes the meter, as documented.
RemoveMeter removes its widget and message; ClearStatus clears the message.
Successful status actions return a Null Variant. No fake Access version, registry,
ActiveX, compiler or object-state result is returned for other actions3251.

Bounded validation happens before changing existing message/meter values: correct
arity450, required arguments, normal VBA String/Long conversions, Null94,
nonempty single-line/NUL-free text5, positive maximum5, update0..maximum5.
Updating without an initialized native meter is unsupported3251. Empty text is
rejected; a single space can replace the status text. Native Qt font/width,
clipping, maximum/range and exact Windows error precedence lack an Access2003
oracle. Status widgets belong to their Qt surface; Application objects still
check the originating document's weak context and fail2450 after closure.
These are actual status actions, not a complete Application/SysCmd implementation.

Existing core/layout2/2 (13.17 s), fresh GUI-free1/1 (1.75 s), current Debian
package/build-ID/SQLite3.54.1 and four original fixture hashes passed. The actual
VBA/UI case observes the maximum, initial/update values, unchanged value after an
invalid update, SetStatus meter removal, ClearStatus and closed-context failure.
No new CTest target or catalog ID. Last full baseline48/48 (83.85 s); no full
rerun claimed after this bounded extension. Existing F068 remains open;456/384/72.

References: [Application.SysCmd](https://learn.microsoft.com/en-us/office/vba/api/access.application.syscmd),
[AcSysCmdAction](https://learn.microsoft.com/en-us/office/vba/api/access.acsyscmdaction).
Only the five longstanding status actions are exposed; newer Office constants
from the evolving online documentation are not added as Access2003 features.
