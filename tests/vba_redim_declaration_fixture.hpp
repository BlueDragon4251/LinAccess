#pragma once
#include <string>
inline const std::string nativeReDimDeclarationFixture=R"VBA(Option Explicit
Public redimExisting() As Long
Private redimBoundsCalls As Long
Private Function ReDimUpper As Long
redimBoundsCalls=redimBoundsCalls+1
ReDimUpper=2
End Function
Private Sub ReDimBump(ByRef items() As Long)
items(1)=items(1)+4
End Sub
Private Function ReDimReturned As Collection()
ReDim result(0 To 0) As VBA.Collection
Set result(0)=New Collection
result(0).Add 7
ReDimReturned=result
End Function
Public Function ReDimDeclarationProbe As Long
If TypeName(late)<>"Long()" Or Not IsArray(late) Then Err.Raise 5
If False Then ReDim late(0) As Long
ReDim late(1 To ReDimUpper()) As Long
late(1)=9
ReDim Preserve late(1 To 3)
ReDimBump late
ReDim redimExisting(4 To 5)
redimExisting(4)=17
ReDimDeclarationProbe=late(1)+UBound(late)+redimBoundsCalls
On Error Resume Next
ReDim late(1 To 4) As String
ReDimDeclarationProbe=ReDimDeclarationProbe+Err.Number
On Error GoTo 0
If UBound(late)<>3 Then Err.Raise 5
Dim returned As Variant
returned=ReDimReturned()
If returned(0).Count<>1 Then Err.Raise 5
End Function
Public Function ReDimExistingProbe As Long
ReDimExistingProbe=redimExisting(4)
End Function
Public Static Function ReDimStaticProbe As Long
Dim initialized As Boolean
If Not initialized Then ReDim retained(0 To 0) As Long
initialized=True
retained(0)=retained(0)+7
ReDimStaticProbe=retained(0)
End Function
)VBA";
