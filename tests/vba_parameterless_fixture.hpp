#pragma once
#include <string>
inline const std::string nativeParameterlessFixture=R"VBA(Private nativeCalls As Long
Public Sub Touch
nativeCalls=nativeCalls+1
End Sub
Public Function Answer As Long
Touch
Answer=nativeCalls
End Function
Public Function CallBare As Long
CallBare=Answer
End Function
Public Property Get Value As Long
Value=nativeCalls
End Property
Public Property Get EmptyValue
EmptyValue=Empty
End Property
Public Static Function Counter As Long
Dim retained As Long
retained=retained+1
Counter=retained
End Function
Public Function TailCounter As Long Static
Dim retained As Long
retained=retained+1
TailCounter=retained
End Function
Public Function FormProbe As Long
Touch
FormProbe=Me.Value
End Function
)VBA";
