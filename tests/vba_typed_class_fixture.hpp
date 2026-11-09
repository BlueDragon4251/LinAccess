#pragma once
#include <string>
inline const std::string typedClassConsumer=R"VBA(Option Explicit
Public held As NativeCounter
Public Property Get TypedHeld As NativeCounter
Set TypedHeld=held
End Property
Public Property Set TypedHeld(ByVal item As NativeCounter)
Set held=item
End Property
Public Function MakeTyped As NativeCounter
Set MakeTyped=New NativeCounter
End Function
Private Function TypedValue(ByVal item As NativeCounter) As Long
TypedValue=item.Value
End Function
Private Function OptionalTyped(Optional ByVal item As NativeCounter=Nothing) As Long
OptionalTyped=7
If Not item Is Nothing Then OptionalTyped=item.Value
End Function
Private Sub ChangeTyped(ByRef item As NativeCounter)
Set item=New NativeCounter
item.Value=8
End Sub
Public Function TypedProbe As Long
Dim item As NativeCounter
Set item=MakeTyped()
Set TypedHeld=item
If Not TypedHeld Is item Then Err.Raise 5
item.Value=6
If Not TypeOf item Is NativeCounter Then Err.Raise 5
If TypeOf item Is Collection Then Err.Raise 5
TypedProbe=TypedValue(item)+OptionalTyped()+OptionalTyped(item)
ChangeTyped item
TypedProbe=TypedProbe+item.Value+held.Value
Set held=Nothing
End Function
Public Function TypedMismatch As Long
On Error Resume Next
Dim item As NativeCounter
Set item=New Collection
TypedMismatch=Err.Number
End Function
Public Function TypedRefMismatch As Long
Dim item As Object
Set item=New NativeCounter
On Error Resume Next
ChangeTyped item
TypedRefMismatch=Err.Number
End Function
Public Function TypedReturnMismatch As NativeCounter
Set TypedReturnMismatch=New Collection
End Function
)VBA";
