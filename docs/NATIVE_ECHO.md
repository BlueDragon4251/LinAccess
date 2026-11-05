# Native Application.Echo and DoCmd.Echo

Application.Echo(EchoOn,[bstrStatusBarText]) and the backward-compatible
DoCmd.Echo(EchoOn,[StatusBarText]) now suppress/restore actual Qt document-surface
repainting. Named arguments use each documented spelling. Normal Boolean/String
coercions apply, Null94, missing/excess arguments450, multiline/NUL status text5.
No SQL/calculated-expression or unqualified Echo function is introduced.

The application-owned native state filters Show events while suppression is
active. It disables native main-window central content and direct content surfaces
of ordinary top-level AccessFormPreview documents. Newly shown ordinary forms are
included. Embedded forms inherit their enclosing native surface. Modal dialogs
and native modal/PopUp documents are excluded; menus/toolbars and status bars are
not deliberately disabled. VBA execution, data mutations and input/event delivery
continue. Echo True restores each weakly held surface's previous updatesEnabled
value; it does not enable a surface that was already independently disabled.
Destroyed surfaces are skipped. No document/model/host is retained by strong state.

Echo remains off across procedure returns until explicitly turned on, unlike
Hourglass's automatic end-of-execution reset. Optional status text uses the actual
native status surface and preserves an existing progress meter; empty text clears
the message. The adapter operates on the current native Qt surface architecture,
not every possible future Qt child/window/WinAPI drawing path. Exact Windows
painting, invalidation/ordering, popup/designer rendering and error precedence
lack an Access2003 execution oracle. Full UI/Echo compatibility is not claimed.

The existing actual VBA/UI flow checks ordinary surfaces disabled, status bar
still enabled, unchanged progress value, a newly opened form suppressed, an
actual modal QMessageBox enabled, restoration on Echo True and preservation of a
previously disabled main surface. Existing targeted mainwindow/layout2/2 (13.25 s),
then full48/48 (85.02 s), current Debian/build-ID/SQLite3.54.1 and all four original
fixture hashes pass. GUI-free runtime unchanged from fresh1/1 (1.67 s). No new
CTest target/catalog item. F068/F080 remain open;456/384/72 stays fixed.

References: [Application.Echo](https://learn.microsoft.com/en-us/office/vba/api/access.application.echo),
[DoCmd.Echo](https://learn.microsoft.com/en-us/office/vba/api/access.docmd.echo).
