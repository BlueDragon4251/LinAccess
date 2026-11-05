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
ActiveX or compiler result is returned for other unimplemented actions3251.

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


## Native GetObjectState

SysCmd action10 / acSysCmdGetObjectState now reports actual native object state
for table0, query1, form2 and report3. Open bit1 comes from the existing project
registry and actual registered tab/dialog membership, including hidden native
runtime documents. Modified bit2 is set only by tracked design editors; pending
data-record changes do not become design changes. New bit4 is available only for
tracked registered design identities. Unsaved anonymous native editors do not
reserve an Access object name, so complete Access temporary naming/new-object
state semantics remain open. Closed or nonexistent native object names return0.
The named action/object-type/object-name arguments use Argument2/Argument3.

The shared checked state reader also backs existing AccessObject.IsLoaded;
connection, database identity, project-generation and Qt owner guards are retained.
Other Access object types explicitly3251. No generic all-object COM/VBE state,
external-instance state, original Access designer bits or complete Windows error
precedence claim. Native acObjStateOpen/Dirty/New constants expose1/2/4.

Existing real mainwindow case checks runtime form open, closed report/nonexistent
object, real query designer open1/modified3/saved1/closed0, and pending editable
query data remaining open1. Targeted core/mainwindow/layout3/3 (15.13 s), fresh
headless1/1 (1.76 s), current package/build-ID/SQLite3.54.1/four hashes pass.
Last full baseline48/48 (83.85 s); no full rerun claimed.456/384/72 unchanged.
