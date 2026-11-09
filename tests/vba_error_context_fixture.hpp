#pragma once
#include <string>
inline const std::string nativeErrorContextFixture=R"VBA(Option Explicit
Public retainedEnvironmentError As VBA.ErrObject
Private Sub SeedEnvironmentError(ByVal expected As VBA.ErrObject)
If Not Err Is expected Then Err.Raise 5
Err.Number=812
Err.Source="shared.source"
Err.Description="shared.description"
Err.HelpFile="shared.help"
Err.HelpContext=12
End Sub
Private Sub HandledEnvironmentError(ByVal expected As VBA.ErrObject)
If Not Err Is expected Then Err.Raise 5
On Error Resume Next
Err.Raise 813,"handled.source","handled.description"
End Sub
Private Sub UnhandledEnvironmentError
Err.Raise 814,"nested.source","nested.description"
End Sub
Public Function EnvironmentErrorProbe As Long
Dim observed As VBA.ErrObject, instance As NativeCounter
Set observed=Err
Err.Clear
SeedEnvironmentError observed
If observed.Number<>812 Or observed.Source<>"shared.source" Or observed.Description<>"shared.description" Or observed.HelpFile<>"shared.help" Or observed.HelpContext<>12 Then Err.Raise 5
On Error Resume Next
HandledEnvironmentError observed
If observed.Number<>813 Or observed.Source<>"handled.source" Then EnvironmentErrorProbe=-1: Exit Function
UnhandledEnvironmentError
If observed.Number<>814 Or observed.Source<>"nested.source" Or observed.Description<>"nested.description" Then EnvironmentErrorProbe=-2: Exit Function
On Error GoTo 0
Set instance=New NativeCounter
If Not instance.SameError(observed) Then Err.Raise 5
Set retainedEnvironmentError=Err
EnvironmentErrorProbe=17
End Function
Public Function EnvironmentIdentityProbe As Boolean
EnvironmentIdentityProbe=Err Is retainedEnvironmentError
End Function
)VBA";
inline const std::string nativeFormErrorContextFixture=R"VBA(Option Explicit
Public Function FormEnvironmentError(ByVal expected As VBA.ErrObject) As Boolean
FormEnvironmentError=Err Is expected
End Function
)VBA";
