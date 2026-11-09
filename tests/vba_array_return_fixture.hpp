#pragma once
#include <string>
inline const std::string nativeArrayReturnFixture=R"VBA(Option Explicit
Private nativeArrayCalls As Long
Private nativeArrayGets As Long
Public Function LinArrayReturnLongs() As Long()
nativeArrayCalls=nativeArrayCalls+1
Dim values(2 To 4) As Long
values(2)=12
values(3)=13
values(4)=14
LinArrayReturnLongs=values
End Function
Public Function LinArrayCallCount As Long
LinArrayCallCount=nativeArrayCalls
End Function
Public Property Get LinArrayReturnMatrix As Long( )
nativeArrayGets=nativeArrayGets+1
Dim values(2 To 3,-1 To 0) As Long
values(2,-1)=21
values(3,-1)=31
values(2,0)=20
values(3,0)=30
LinArrayReturnMatrix=values
End Property
Public Property Get LinArrayGetCount As Long
LinArrayGetCount=nativeArrayGets
End Property
Public Function LinArrayBytes As Byte()
Dim values(0 To 2) As Byte
values(0)=1
values(1)=2
values(2)=255
LinArrayBytes=values
End Function
Public Function LinArrayStrings As String()
Dim values(1 To 2) As String
values(1)="one"
values(2)="ä"
LinArrayStrings=values
End Function
Public Function LinArrayObjects As Object()
Dim values(1 To 1) As Object
Set values(1)=CreateObject("Scripting.Dictionary")
values(1).Add "value",19
LinArrayObjects=values
End Function
Public Function LinArrayVariant As Variant
LinArrayVariant=LinArrayReturnLongs()
End Function
Public Function LinArraySized(ByVal size As Long) As Long()
Dim values() As Long
ReDim values(0 To size)
Dim i As Long
For i=0 To size
values(i)=i*2
Next i
LinArraySized=values
End Function
Public Function LinArrayUnallocated As Long()
End Function
Public Function LinArrayWrong As Long()
Dim values(1 To 2) As String
LinArrayWrong=values
End Function
Public Function LinArrayReadProbe As Long
LinArrayReadProbe=LinArrayReturnLongs(2)+LinArrayReturnMatrix(3,0)+LinArrayBytes(2)
End Function
Public Function LinArrayCopyProbe As Long
Dim first() As Long
Dim second() As Long
first=LinArrayReturnLongs()
second=first
second(2)=99
LinArrayCopyProbe=first(2)*100+second(2)
End Function
Public Function LinArrayQueueProbe As Long
LinArrayQueueProbe=LinArraySized(2)(1)+LinArrayVariant(2)
End Function
Public Function LinArrayBoundsProbe As Long
LinArrayBoundsProbe=LinArrayReturnLongs(9)
End Function
Public Function LinArrayEachProbe As Long
Dim item As Variant
Dim before As Long
before=LinArrayCallCount()
For Each item In LinArrayReturnLongs()
LinArrayEachProbe=LinArrayEachProbe+item
Next
LinArrayEachProbe=LinArrayEachProbe*100+LinArrayCallCount()-before
End Function
Public Function LinArrayEachMatrix As String
Dim item As Variant
For Each item In LinArrayReturnMatrix
LinArrayEachMatrix=LinArrayEachMatrix & CStr(item) & ":"
Next
End Function
Public Function LinArrayEachObjects As Long
Dim item As Variant
For Each item In LinArrayObjects()
LinArrayEachObjects=LinArrayEachObjects+item("value")
Next
End Function
Public Function LinArrayIdentity(ByRef source() As Long) As Long()
LinArrayIdentity=source
End Function
Public Function LinArrayEachCopy As Long
Dim source() As Long
Dim item As Variant
source=LinArrayReturnLongs()
For Each item In LinArrayIdentity(source)
source(3)=99
LinArrayEachCopy=LinArrayEachCopy+item
Next
LinArrayEachCopy=LinArrayEachCopy*100+source(3)
End Function
Public Function LinArrayEachNested As Long
Dim outerItem As Variant
Dim innerItem As Variant
For Each outerItem In LinArraySized(2)
For Each innerItem In LinArrayBytes()
LinArrayEachNested=LinArrayEachNested+outerItem+innerItem
Exit For
Next
Next
End Function
)VBA";
