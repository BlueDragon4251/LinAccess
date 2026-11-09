#pragma once
#include <string>
inline const std::string nativeErrorClearFixture=R"VBA(Option Explicit
Public retainedExitError As VBA.ErrObject
Private Sub ClearExitSub
Set retainedExitError=Err
Err.Number=706
Err.Description="sub pending"
Exit Sub
End Sub
Private Function ClearExitFunction As VBA.ErrObject
Set ClearExitFunction=Err
Err.Number=705
Err.Source="function pending"
Exit Function
End Function
Private Property Get ClearExitProperty As VBA.ErrObject
Set ClearExitProperty=Err
Err.Number=707
Err.HelpFile="property pending"
Exit Property
End Property
Public Function ErrorClearProbe As Long
Dim observed As VBA.ErrObject
Set observed=Err
On Error Resume Next
Err.Raise 701,"source","description","help",12
If observed.Number<>701 Then ErrorClearProbe=-1: Exit Function
On Error Resume Next
If observed.Number<>0 Or observed.Source<>"" Or observed.Description<>"" Or observed.HelpFile<>"" Or observed.HelpContext<>0 Then ErrorClearProbe=-2: Exit Function
Err.Raise 702,"source","description","help",12
If False Then On Error GoTo 0
If observed.Number<>702 Then ErrorClearProbe=-3: Exit Function
On Error GoTo 0
If observed.Number<>0 Or observed.HelpContext<>0 Then Err.Raise 5
Err.Number=703
Err.Description="pending"
On Error GoTo recovered
If observed.Number<>0 Or observed.Description<>"" Then Err.Raise 5
Err.Raise 704
Exit Function
recovered:
ErrorClearProbe=observed.Number
End Function
Public Function ErrorExitProbe As Long
Dim observed As VBA.ErrObject
Set observed=ClearExitFunction()
If observed.Number<>0 Or observed.Source<>"" Then Err.Raise 5
ClearExitSub
If retainedExitError.Number<>0 Or retainedExitError.Description<>"" Then Err.Raise 5
Set observed=ClearExitProperty
If observed.Number<>0 Or observed.HelpFile<>"" Then Err.Raise 5
ErrorExitProbe=17
End Function
)VBA";
