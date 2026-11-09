#pragma once
#include <string>
inline const std::string nativeByRefPropertyFixture=R"VBA(Option Explicit
Private lastWritten As Long
Private getCalls As Long
Private held As Object
Public Property Get Counter(ByRef index As Long) As Long
getCalls=getCalls+1
index=index+1
Counter=index*10
End Property
Public Property Let Counter(ByRef index As Long, ByRef value As Long)
index=index+1
lastWritten=value
value=value+100
End Property
Public Property Get Calls As Long
Calls=getCalls
End Property
Public Function LetProbe As Long
Dim index As Long,value As Long
index=1
value=7
Counter(index)=value
LetProbe=index*100+value*10+lastWritten
End Function
Public Function GetProbe As Long
Dim index As Long,result As Long
index=1
result=Counter(index)
GetProbe=result+index
End Function
Public Function TemporaryProbe As Long
Dim index As Long,result As Long
index=1
result=Counter(index+0)
TemporaryProbe=result+index
End Function
Public Function ArrayProbe As Long
Dim indices(0 To 0) As Long,result As Long
indices(0)=1
Counter(indices(0))=7
result=Counter(indices(0))
ArrayProbe=indices(0)*100+result
End Function
Public Property Get ObjectSlot(ByRef index As Long) As Object
index=index+1
Set ObjectSlot=held
End Property
Public Property Set ObjectSlot(ByRef index As Long,ByRef value As Object)
index=index+1
Set held=value
Set value=Nothing
End Property
Public Function ObjectProbe As Long
Dim index As Long
Dim child As Object,got As Object
index=1
Set child=New Scripting.Dictionary
child.Add "value",19
Set ObjectSlot(index)=child
Set got=ObjectSlot(index)
If got Is child Then ObjectProbe=index*100+got("value")
End Function
Public Property Let Failing(ByRef index As Long,ByVal value As Long)
index=index+1
Err.Raise 11
End Property
Public Function ErrorProbe As Long
Dim index As Long
index=1
On Error Resume Next
Failing(index)=7
ErrorProbe=index*100+Err.Number
End Function
Public Function MismatchProbe As Long
Dim index As Double,result As Long
index=1
On Error Resume Next
result=Counter(index)
MismatchProbe=Err.Number
End Function
)VBA";
