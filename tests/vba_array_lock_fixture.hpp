#pragma once
#include <string>
inline const std::string nativeArrayLockFixture=R"VBA(Option Explicit
Private lockedValues() As Long
Private lockErrors As Long
Private Function OtherLockValues As Long()
Dim items(0 To 1) As Long
items(0)=17
OtherLockValues=items
End Function
Private Sub ReplaceLockedValues(ByRef value As Long)
value=value+1
On Error Resume Next
ReDim lockedValues(0 To 2)
lockErrors=Err.Number
Err.Clear
Erase lockedValues
lockErrors=lockErrors+Err.Number
Err.Clear
lockedValues=OtherLockValues()
lockErrors=lockErrors+Err.Number
On Error GoTo 0
value=value+1
End Sub
Private Sub ForwardLockedValue(ByRef value As Long)
ReplaceLockedValues value
End Sub
Public Function ArrayLockProbe As Long
ReDim lockedValues(0 To 1)
lockedValues(0)=7
ForwardLockedValue lockedValues(0)
If lockedValues(0)<>9 Or UBound(lockedValues)<>1 Or lockErrors<>30 Then Err.Raise 5
ReDim Preserve lockedValues(0 To 2)
ArrayLockProbe=lockedValues(0)+lockErrors+UBound(lockedValues)
End Function
)VBA";
