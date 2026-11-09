#pragma once
#include <string>
inline const std::string nativeOptionalPropertyFixture=R"VBA(Option Explicit
Private last As Long
Private held As Object
Private writeCount As Long
Private readCount As Long
Public Property Get Written As Long
Written=last
End Property
Public Property Get Cell(Optional ByVal row As Long=2,Optional ByVal col As Long=3) As Long
Cell=row*10+col
End Property
Public Property Let Cell(Optional ByVal row As Long=4,Optional ByVal col As Long=5,ByRef value As Long)
last=row*100+col*10+value
value=value+100
End Property
Public Function DefaultsProbe As Long
Dim value As Long
value=7
Cell=value
DefaultsProbe=last+Cell()+value
End Function
Public Function HoleProbe As Long
Cell(,8)=7
HoleProbe=last+Cell(,8)+Cell(col:=8)
End Function
Public Function EmptyCallProbe As Long
Cell()=7
EmptyCallProbe=last
End Function
Public Property Get Owned(Optional ByRef first As Long=3,Optional ByRef second As Long=4) As Long
first=first+1
second=second+1
Owned=first*100+second*10
End Property
Public Property Let Owned(Optional ByRef first As Long=5,Optional ByRef second As Long=6,ByRef value As Long)
first=first+1
second=second+1
last=first*1000+second*100+value
value=value+100
End Property
Public Function OwnedProbe As Long
Dim index As Long,value As Long,result As Long
index=1
value=7
Owned(index)=value
result=Owned(first:=index)
OwnedProbe=last+index+value+result
End Function
Public Property Get Many(ParamArray values() As Variant) As Long
Dim sum As Long,part As Variant
sum=0
For Each part In values
sum=sum+part
Next
Many=(UBound(values)+1)*100+sum
End Property
Public Property Let Many(ParamArray values() As Variant,ByRef value As Long)
Dim sum As Long,part As Variant
sum=0
For Each part In values
sum=sum+part
Next
last=(UBound(values)+1)*100+sum*10+value
value=value+100
End Property
Public Function ManyProbe As Long
Many(1,2,3)=7
ManyProbe=last+Many(1,2,3)
End Function
Public Function EmptyManyProbe As Long
Many()=7
EmptyManyProbe=last+Many()
End Function
Public Property Get OptionalObject(Optional ByVal index As Long=2) As Object
readCount=index
Set OptionalObject=held
End Property
Public Property Set OptionalObject(Optional ByVal index As Long=4,ByRef value As Object)
writeCount=index
Set held=value
Set value=Nothing
End Property
Public Function OptionalObjectProbe As Long
Dim child As Object,got As Object
Set child=New Scripting.Dictionary
child.Add "value",19
Set OptionalObject()=child
Set got=OptionalObject()
If got Is child Then OptionalObjectProbe=writeCount*100+readCount*10+child("value")
End Function
Public Property Get ManyObject(ParamArray values() As Variant) As Object
readCount=UBound(values)+1
Set ManyObject=held
End Property
Public Property Set ManyObject(ParamArray values() As Variant,ByRef value As Object)
writeCount=UBound(values)+1
Set held=value
Set value=Nothing
End Property
Public Function ManyObjectProbe As Long
Dim child As Object,got As Object
Set child=New Scripting.Dictionary
child.Add "value",19
Set ManyObject(1,2)=child
Set got=ManyObject(1,2,3)
If got Is child Then ManyObjectProbe=writeCount*100+readCount*10+child("value")
End Function
Public Property Get GuardedValues(ParamArray values() As Variant) As Long
On Error Resume Next
values(0)=99
GuardedValues=values(0)*1000+Err.Number
End Property
Public Function GuardedValuesProbe As Long
GuardedValuesProbe=GuardedValues(7)
End Function
)VBA";
