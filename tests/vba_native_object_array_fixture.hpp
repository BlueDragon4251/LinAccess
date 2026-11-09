#pragma once
#include <string>
inline const std::string nativeObjectArrayFixture=R"VBA(Option Explicit
Public collections() As VBA.Collection
Public dictionaries() As Scripting.Dictionary
Private Sub GrowCollections(ByRef items() As VBA.Collection)
ReDim Preserve items(-1 To 1)
Set items(1)=New Collection
items(1).Add 9
End Sub
Public Function MakeCollections As VBA.Collection()
Dim items(-1 To 0) As Collection
Set items(-1)=New Collection
items(-1).Add 11
Set items(0)=New Collection
items(0).Add 12
items(0).Add 13
MakeCollections=items
End Function
Public Function NativeObjectArrayProbe As Long
Dim fixed(1 To 2) As Collection, errors() As VBA.ErrObject, fault As ErrObject, item As Variant
ReDim collections(-1 To 0)
Set collections(-1)=New Collection
collections(-1).Add 7
Set collections(0)=New Collection
collections(0).Add 8
collections(0).Add 9
Set fixed(1)=collections(-1)
GrowCollections collections
ReDim dictionaries(2 To 3)
Set dictionaries(2)=New Scripting.Dictionary
dictionaries(2).Add "value",4
ReDim errors(0 To 0)
Set fault=Err
Set errors(0)=fault
If Not TypeOf errors(0) Is VBA.ErrObject Then Err.Raise 5
If Not errors(0) Is fault Then Err.Raise 5
If TypeName(errors)<>"ErrObject()" Then Err.Raise 5
If TypeName(collections)<>"Collection()" Then Err.Raise 5
If TypeName(dictionaries)<>"Dictionary()" Then Err.Raise 5
If VarType(dictionaries)<>8201 Then Err.Raise 5
NativeObjectArrayProbe=collections(-1).Count+collections(0).Count+collections(1).Count+fixed(1).Count+dictionaries(2)("value")
For Each item In MakeCollections()
NativeObjectArrayProbe=NativeObjectArrayProbe+item.Count
Next
On Error Resume Next
Set dictionaries(2)=New Collection
NativeObjectArrayProbe=NativeObjectArrayProbe+errors(0).Number
On Error GoTo 0
Erase collections,dictionaries
End Function
)VBA";
inline const std::string nativeAccessArrayFixture=R"VBA(Option Explicit
Public Function NativeAccessArrayProbe As Long
Dim forms(0 To 0) As Access.Form, boxes() As Access.TextBox, controls() As Access.Control
Set forms(0)=Me
ReDim boxes(1 To 2),controls(1 To 1)
Set boxes(1)=forms(0).Controls("ArrayFirst")
Set boxes(2)=Me.Controls("ArraySecond")
Set controls(1)=boxes(2)
boxes(1).Value=17
controls(1).Value=9
If TypeName(boxes)<>"TextBox()" Or VarType(forms)<>8201 Then Err.Raise 5,,TypeName(boxes) & ":" & CStr(VarType(forms))
If Not TypeOf controls(1) Is Access.TextBox Then Err.Raise 5
NativeAccessArrayProbe=CLng(boxes(1).Value)+CLng(controls(1).Value)
On Error Resume Next
Set boxes(2)=Me
NativeAccessArrayProbe=NativeAccessArrayProbe+Err.Number
On Error GoTo 0
If Not boxes(2) Is controls(1) Then Err.Raise 5
End Function
)VBA";
