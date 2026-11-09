#pragma once
#include <string>
inline const std::string typedClassConsumer=R"VBA(Option Explicit
Public held As NativeCounter
Public storedClassArray() As NativeCounter
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
Private Sub GrowClassArray(ByRef items() As NativeCounter)
ReDim Preserve items(-1 To 2)
Set items(2)=New NativeCounter
items(2).Value=9
End Sub
Public Function ClassArrayResult As NativeCounter()
Dim items() As NativeCounter
ReDim items(3 To 4)
Set items(3)=New NativeCounter
items(3).Value=7
Set items(4)=New NativeCounter
ClassArrayResult=items
End Function
Public Function EmptyClassArray As NativeCounter()
End Function
Public Function ClassArrayProbe As Long
Dim items() As NativeCounter, fixed(1 To 2) As NativeCounter, copied As Variant, item As Variant
ReDim items(-1 To 0)
Set items(-1)=New NativeCounter
items(-1).Value=5
Set fixed(1)=items(-1)
copied=items
ReDim Preserve items(-1 To 1)
Set items(0)=fixed(1)
Set items(1)=New NativeCounter
For Each item In items
ClassArrayProbe=ClassArrayProbe+item.Value
Next
GrowClassArray items
ClassArrayProbe=ClassArrayProbe+items(2).Value
If UBound(copied)<>0 Then Err.Raise 5
If Not items(-1) Is copied(-1) Then Err.Raise 5
If VarType(items)<>8201 Then Err.Raise 5
If TypeName(items)<>"NativeCounter()" Then Err.Raise 5
If TypeName(EmptyClassArray())<>"NativeCounter()" Then Err.Raise 5
Erase fixed
If Not fixed(1) Is Nothing Then Err.Raise 5
For Each item In ClassArrayResult()
ClassArrayProbe=ClassArrayProbe+item.Value
Next
storedClassArray=items
If storedClassArray(2).Value<>9 Then Err.Raise 5
Erase storedClassArray
End Function
Public Function ClassArrayMismatch As Long
Dim items(0 To 0) As NativeCounter
On Error Resume Next
Set items(0)=New Collection
ClassArrayMismatch=Err.Number
End Function
Public Function ClassArrayRefMismatch As Long
Dim items() As Object
ReDim items(-1 To 0)
On Error Resume Next
GrowClassArray items
ClassArrayRefMismatch=Err.Number
End Function
)VBA";
