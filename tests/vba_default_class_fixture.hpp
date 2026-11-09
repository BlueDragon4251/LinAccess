#pragma once
#include <string>
inline const std::string nativeDefaultClassFixture=R"VBA(VERSION 1.0 CLASS
Attribute VB_Name = "NativeDefault"
Option Explicit
Private stored As Long
Public Reads As Long
Private Sub Class_Initialize()
stored=4
End Sub
Public Property Get Content(Optional ByVal index As Long=0) As Long
Attribute Content.VB_UserMemId = 0
Reads=Reads+1
Content=stored+index
End Property
Public Property Let Content(Optional ByVal index As Long=0,ByVal amount As Long)
stored=amount-index
End Property
Public Function Item(ByVal index As Long) As Long
Item=999+index
End Function
)VBA";
inline const std::string nativeDefaultClassConsumer=R"VBA(Option Explicit
Public Function DefaultClassProbe As Long
Dim value As NativeDefault, inspected As Long
Set value=New NativeDefault
value=11
DefaultClassProbe=value+value(2)+value(index:=3)
value(2)=17
DefaultClassProbe=DefaultClassProbe+value.Content+value.Reads+value.Item(2)
On Error Resume Next
inspected=value(unknown:=1)
If Err.Number<>448 Then Err.Raise 5
On Error GoTo 0
End Function
)VBA";
