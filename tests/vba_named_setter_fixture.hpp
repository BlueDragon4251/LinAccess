#pragma once
#include <string>
inline const std::string nativeNamedSetterFixture=R"VBA(Option Explicit
Public backing As Long,getterCalls As Long,writeCalls As Long,rhsReads As Long
Public held As Object
Public Property Get Cell(ByRef first As Long,ByRef second As Long) As Long
getterCalls=getterCalls+1
Cell=backing
End Property
Public Property Let Cell(ByRef first As Long,ByRef second As Long,ByRef value As Long)
first=first+1
second=second+2
backing=value
writeCalls=writeCalls+1
value=999
End Property
Public Property Let OptionalCell(Optional ByVal first As Long=3,Optional ByVal second As Long=4,ByVal value As Long)
backing=first*100+second*10+value
writeCalls=writeCalls+1
End Property
Public Property Set Slot(ByRef first As Long,ByRef second As Long,ByRef value As Object)
first=first+1
second=second+2
Set held=value
writeCalls=writeCalls+1
Set value=Nothing
End Property
Public Property Set PublicSlot(ByVal selector As Long,ByVal value As Object)
Set held=value
writeCalls=writeCalls+1
End Property
Public Property Let Into(ByVal target As Collection,ByVal value As Long)
target.Add value
writeCalls=writeCalls+1
End Property
Private Function Rhs As Long
rhsReads=rhsReads+1
Rhs=19
End Function
Public Function NamedSetterProbe As Long
Dim a As Long,b As Long,value As Long,indices(0 To 1) As Long,child As Object,errors As Long
getterCalls=0
writeCalls=0
rhsReads=0
a=1
b=2
value=7
Cell(second:=b,first:=a)=value
indices(0)=1
indices(1)=2
Cell(second:=indices(1),first:=indices(0))=8
Cell(second:=b+0,first:=a+0)=9
OptionalCell(second:=7)=5
Set child=New Collection
child.Add 1
Set Slot(second:=b,first:=a)=child
Into(target:=child)=11
On Error Resume Next
Cell(unknown:=a)=Rhs()
errors=Err.Number
Err.Clear
Cell(first:=a,first:=b)=Rhs()
errors=errors+Err.Number
Err.Clear
Cell(second:=b)=Rhs()
errors=errors+Err.Number
Err.Clear
Cell(first:=a,b)=Rhs()
errors=errors+Err.Number
On Error GoTo 0
If getterCalls<>0 Or writeCalls<>6 Or rhsReads<>0 Then Err.Raise 5
If Not held Is child Then Err.Raise 5
NamedSetterProbe=a*1000+b*100+indices(0)*10+indices(1)+value+backing+child.Count+errors
End Function
Private Function ChangePacked(ParamArray values()) As Long
Cell(second:=values(1),first:=values(0))=17
ChangePacked=values(0)*10+values(1)
End Function
Public Function NamedSetterPackProbe As Long
Dim a As Long,b As Long,result As Long
a=1
b=2
result=ChangePacked(a,b)
NamedSetterPackProbe=a*100+b*10+result
End Function
Public Function NamedFormSetterProbe As Long
Dim child As Object,a As Long,b As Long,errors As Long
getterCalls=0
writeCalls=0
rhsReads=0
OptionalCell(second:=7)=5
Me.OptionalCell(first:=2,second:=8)=9
Set child=New Collection
Set Me.PublicSlot(selector:=2)=child
Me.Into(target:=child)=11
On Error Resume Next
Me.OptionalCell(unknown:=3)=Rhs()
errors=Err.Number
Err.Clear
Me.Cell(first:=a,second:=b)=Rhs()
errors=errors+Err.Number
On Error GoTo 0
If getterCalls<>0 Or writeCalls<>4 Or rhsReads<>0 Then Err.Raise 5
If Not held Is child Then Err.Raise 5
NamedFormSetterProbe=backing+child.Count+writeCalls+errors
End Function
)VBA";
inline const std::string nativeNamedSetterClassFixture=R"VBA(VERSION 1.0 CLASS
Attribute VB_Name = "NamedWriter"
Option Explicit
Public Backing As Long,Reads As Long,Writes As Long
Public Held As Object
Public Property Get Content(Optional ByVal first As Long=3,Optional ByVal second As Long=4) As Long
Attribute Content.VB_UserMemId = 0
Reads=Reads+1
Content=Backing
End Property
Public Property Let Content(Optional ByVal first As Long=3,Optional ByVal second As Long=4,ByVal amount As Long)
Backing=first*100+second*10+amount
Writes=Writes+1
End Property
Public Property Let WriteOnly(ByVal index As Long,ByVal amount As Long)
Backing=index+amount
Writes=Writes+1
End Property
Public Property Set Slot(ByVal selector As Long,ByVal child As Object)
Set Held=child
Writes=Writes+1
End Property
Public Property Let Into(ByVal target As Collection,ByVal amount As Long)
target.Add amount
Writes=Writes+1
End Property
)VBA";
inline const std::string nativeNamedSetterClassConsumer=R"VBA(Option Explicit
Private rhsReads As Long
Private Function Rhs As Long
rhsReads=rhsReads+1
Rhs=19
End Function
Public Function NamedClassSetterProbe As Long
Dim writer As NamedWriter,child As Object,errors As Long
Set writer=New NamedWriter
writer.Content(second:=7)=5
writer(second:=8,first:=2)=9
writer.WriteOnly(index:=3)=4
Set child=New Collection
Set writer.Slot(selector:=2)=child
writer.Into(target:=child)=11
On Error Resume Next
writer.Content(unknown:=3)=Rhs()
errors=Err.Number
On Error GoTo 0
If writer.Reads<>0 Or writer.Writes<>5 Or rhsReads<>0 Then Err.Raise 5
If Not writer.Held Is child Then Err.Raise 5
NamedClassSetterProbe=writer.Backing+writer.Writes+child.Count+errors
End Function
)VBA";
