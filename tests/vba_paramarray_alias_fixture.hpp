#pragma once
#include <string>
inline const std::string nativeParamArrayAliasFixture=R"VBA(Option Explicit
Private tracked As Long
Private slotErrors As Long
Private Function ReadTracked As Long
ReadTracked=tracked
End Function
Private Sub BumpVariant(ByRef value As Variant)
value=value+2
End Sub
Private Sub AliasTwo(ByRef first As Variant,ByRef second As Variant)
first=first+1
second=second+first
End Sub
Private Function EditPack(ParamArray values()) As String
values(0)=9
values(1)="changed"
values(2)=42
values(3)=7
values(4)=values(0)+1
BumpVariant values(0)
AliasTwo values(0),values(4)
If ReadTracked()<>26 Or values(4)<>26 Then Err.Raise 5
Set values(5)=New Collection
values(5).Add 17
On Error Resume Next
values(0)="invalid"
slotErrors=Err.Number
Err.Clear
values(6)=256
slotErrors=slotErrors+Err.Number
Err.Clear
values(7)=Array(2)
slotErrors=slotErrors+Err.Number
On Error GoTo 0
EditPack=CStr(values(0)) & ":" & CStr(values(4)) & ":" & CStr(values(2)) & ":" & CStr(values(3))
End Function
Public Function ParamArrayAliasProbe As Long
Dim text As String,bag As Object,tiny As Byte,items As Variant,trace As String
tracked=1
text="old"
Set bag=New Scripting.Dictionary
tiny=2
items=Array(1)
trace=EditPack(tracked,text,0,(tracked),tracked,bag,tiny,items)
If trace<>"26:26:42:7" Or text<>"changed" Or TypeName(bag)<>"Collection" Then Err.Raise 5
If tiny<>2 Or items(0)<>1 Or slotErrors<>3270 Then Err.Raise 5
ParamArrayAliasProbe=tracked+tiny+items(0)+bag.Count+slotErrors
End Function
Private Function ElementPack(ParamArray values()) As Long
values(0)=values(0)+3
BumpVariant values(0)
ElementPack=values(0)
End Function
Public Function ParamArrayElementProbe As Long
Dim items() As Long,value As Long
ReDim items(0 To 0)
items(0)=4
value=ElementPack(items(0))
ReDim Preserve items(0 To 1)
ParamArrayElementProbe=value+items(0)+UBound(items)
End Function
Private Function EmptyMutation(ParamArray values()) As Long
On Error Resume Next
values(0)=1
EmptyMutation=Err.Number
End Function
Public Function ParamArrayEmptyProbe As Long
ParamArrayEmptyProbe=EmptyMutation()
End Function
Public Function UnboundPack(ParamArray values()) As Long
values(0)=99
UnboundPack=values(0)
End Function
)VBA";
