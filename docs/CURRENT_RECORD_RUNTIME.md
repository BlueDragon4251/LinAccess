# Bound form CurrentRecord

Native bound forms expose CurrentRecord as a read-only VBA Long. Existing rows
use their current one-based position in the loaded, filtered recordset. The
supported new-record row uses loaded record count + 1, without creating or saving
a record. This is a position, not the row's database identity or bookmark.

The value is available through Me, Forms, Screen.ActiveForm, unqualified form VBA
and calculated control expressions. Actual field/control expression names retain
their existing lookup precedence. Reading the property does not save pending
changes. Navigation/Current events and filters use the updated position.

The existing native unavailable/closed/read-only guards remain in force. Reports
reject the member (438); closed form references report 2450. Unbound/unavailable
sources and a new row whose ordinal cannot be established (including truncated
sources) explicitly report unsupported operation 3251. These bounded cases do
not claim complete Windows error-number identity. No writable design-view
CurrentRecord or unbound-form behavior is claimed.

`form_current_record` uses real Qt controls, VBA, DAO/SQL event auditing and
navigation/filter/new-record commands. It checks Long/VarType 3, calculated
position text, pending edits without saves, read-only assignment and lifecycle
guards. The public contract reference is
https://learn.microsoft.com/en-us/office/vba/api/access.form.currentrecord
(position in the displayed recordset; Long; read-only in Form/Datasheet view).
No Windows Access differential run was available.
