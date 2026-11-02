# Native form window modes and dialog continuation

OpenForm now implements native Qt normal, hidden, icon/minimized and dialog
opening. acDialog sets Modal and PopUp metadata to True and uses application
modality. Its VBA caller waits in a Qt event loop until the new dialog closes or
is hidden. A cancelled Unload does not finish the wait. Closed native documents
retain the existing deferred-disposal protection while outer VBA frames run.

A hidden form can be shown for the first time through Visible=True. Existing
forms can be hidden through OpenForm using the native visibility guards. Modal
and PopUp are readable; imported PopUp is retained. Runtime changes to those two
flags are explicitly unsupported. DoCmd.Minimize/Restore/Maximize act on the
native active document; icon opening uses real minimized window state.

Existing GUI execution audits actual event/SQL order for dialog close, hide and
one cancelled Unload followed by accepted close; it checks Modal/PopUp values,
hidden reveal/re-hide, minimized opening and window commands targeting another
active document. No new test target is added.

## Boundaries

Dialog reuse of an already-open instance, data/view modes, modal visibility with
pending edits or other unsupported lifecycle phases, report modes, full desktop
focus/z-order/frame identity, runtime Modal/PopUp mutation and complete Windows
behavior remain open. Native application-Quit propagation across suspended VBA
frames is a separate follow-up; this segment does not claim complete Quit parity.

Mode reference: https://learn.microsoft.com/en-us/office/vba/api/access.acwindowmode
