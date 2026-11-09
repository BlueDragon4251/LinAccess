#pragma once
#include <string>
inline const std::string nativeStaticFixture=R"VBA(Private propertyBacking As Long
Public Static Function Retained() As String
Dim count As Long, bag As Object
Dim fixed(1 To 2) As Long, dynamic() As Long
Dim label As String * 4
count=count+1
If bag Is Nothing Then Set bag=New Scripting.Dictionary
If Not bag.Exists("calls") Then bag.Add "calls",0
bag("calls")=bag("calls")+1
If count=1 Then ReDim dynamic(1)
fixed(1)=fixed(1)+2
dynamic(0)=dynamic(0)+3
label="A" & CStr(count)
Retained=CStr(count) & ":" & CStr(fixed(1)) & ":" & CStr(dynamic(0)) & ":" & CStr(bag("calls")) & ":" & Trim$(label)
End Function
Private Sub Bump(ByRef value As Variant)
value=value+1
End Sub
Static Function Implicit() As String
count=count+1
If IsEmpty(bag) Then Set bag=New Scripting.Dictionary
If Not bag.Exists("calls") Then bag.Add "calls",0
bag("calls")=bag("calls")+1
Bump byrefOnly
For index=1 To 2
total=total+index
Next
If IsEmpty(values) Then ReDim values(1)
values(0)=values(0)+1
Implicit=CStr(count) & ":" & CStr(bag("calls")) & ":" & CStr(byrefOnly) & ":" & CStr(total) & ":" & CStr(values(0))
End Function
Static Function OtherImplicit() As Long
count=count+100
OtherImplicit=count
End Function
Private Static Function Recursive(ByVal depth As Long) As Long
Dim count As Long
count=count+1
If depth>0 Then ignored=Recursive(depth-1)
Recursive=count
End Function
Public Function Trailing() As Long Static
Dim count As Long
count=count+1
Trailing=count
End Function
Static Function FreshResult(ByVal assign As Boolean) As Long
If assign Then FreshResult=5
End Function
Static Function FreshParameter(ByVal value As Long) As Long
Dim count As Long
count=count+1
value=value+1
FreshParameter=value+count
End Function
Public Static Property Get Accumulated() As Long
Dim reads As Long
reads=reads+1
Accumulated=propertyBacking+reads
End Property
Public Property Let Accumulated(value As Long) Static
Dim writes As Long
writes=writes+1
propertyBacking=value
End Property
Static Function Failure() As Long
Dim count As Long
count=count+1
Err.Raise 11
End Function
)VBA";
