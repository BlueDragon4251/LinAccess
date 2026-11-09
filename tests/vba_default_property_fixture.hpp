#pragma once
#include <string>
inline const std::string nativeDefaultPropertyFixture=R"VBA(Option Explicit
Private retained As Object
Private getterReads As Long
Private functionReads As Long
Public Property Get Bag As Object
getterReads=getterReads+1
If retained Is Nothing Then Set retained=New Scripting.Dictionary
Set Bag=retained
End Property
Public Property Get ReadCount As Long
ReadCount=getterReads
End Property
Public Property Get FunctionCalls As Long
FunctionCalls=functionReads
End Property
Public Function BagFunction As Object
functionReads=functionReads+1
If retained Is Nothing Then Set retained=New Scripting.Dictionary
Set BagFunction=retained
End Function
Public Function DefaultProbe As Long
Bag("value")=7
Dim child As Object
Set child=New Scripting.Dictionary
child.Add "value",11
Set Bag("object")=child
Dim got As Object
Set got=Bag("object")
got("value")=12
DefaultProbe=Bag("value")+BagFunction("value")+got("value")
End Function
Public Property Get Scalar As Long
getterReads=getterReads+1
Scalar=1
End Property
Public Function RejectScalar As Long
RejectScalar=Scalar(1)
End Function
Public Property Get MissingObject As Object
Set MissingObject=Nothing
End Property
Public Function MissingProbe As Long
MissingProbe=MissingObject("value")
End Function
Public Property Get FailedObject As Object
Err.Raise 11
End Property
Public Function FailedProbe As Long
FailedProbe=FailedObject("value")
End Function
)VBA";
