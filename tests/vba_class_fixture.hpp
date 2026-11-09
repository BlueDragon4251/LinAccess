#pragma once
#include <string>
inline const std::string nativeClassFixture=R"VBA(VERSION 1.0 CLASS
Attribute VB_Name = "NativeCounter"
Option Explicit
Private stored As Long
Private reads As Long
Private storedBag As Object
Public Calls As Long
Public Child As Object
Public Cells() As Long
Public Const ConstantValue=4
Private Sub Class_Initialize()
stored=1
Calls=1
Set storedBag=CreateObject("Scripting.Dictionary")
storedBag.Add "value",5
ReDim Cells(1 To 2)
Cells(1)=4
End Sub
Public Property Get Value As Long
reads=reads+1
Value=stored
End Property
Public Property Let Value(ByVal amount As Long)
stored=amount
End Property
Public Property Get ReadCount As Long
ReadCount=reads
End Property
Public Property Get Bag As Object
Set Bag=storedBag
End Property
Public Property Get Values As Long()
Dim result(2 To 3) As Long
result(2)=8
result(3)=9
Values=result
End Property
Public Property Get Indexed(ByVal index As Long) As Long
Indexed=stored+index
End Property
Public Property Let Indexed(ByVal index As Long,ByVal amount As Long)
stored=amount-index
End Property
Private Function Hidden(ByVal amount As Long) As Long
Hidden=stored+amount
End Function
Public Sub Add(ByVal amount As Long)
Me.Value=Me.Hidden(amount)
Calls=Calls+1
End Sub
Public Function Label(Optional ByVal prefix As String="v") As String
Label=prefix & CStr(stored)
End Function
Public Function Identity As Object
Set Identity=Me
End Function
Public Sub RequiresRef(ByRef amount As Long)
amount=amount+1
Calls=Calls+1
End Sub
Public Function OwnedRef As Long
Dim amount As Long
amount=2
RequiresRef amount
OwnedRef=amount
End Function
Public Function ExternalValue As Long
ExternalValue=ProjectValue
End Function
Public Function ExternalMe As Long
On Error Resume Next
ExternalMe=ProjectMe
ExternalMe=Err.Number
End Function
Public Sub Trigger()
MsgBox "technical callback"
Me.Value=Me.Hidden(1)
End Sub
Public Function SaveAudit(ByVal amount As Long) As Long
CurrentDb.Execute "INSERT INTO native_class_audit(value) VALUES(" & CStr(amount) & ")"
SaveAudit=amount
End Function
)VBA";
inline const std::string nativeClassConsumer=R"VBA(Option Explicit
Public Function ClassProbe As Long
Dim first As Object,second As Object,aliasRef As Object
Set first=New NativeCounter
Set second=New NativeCounter
Set aliasRef=first
first.Value=7
first.Add 2
Set first.Child=second
ClassProbe=first.Value*100+second.Value*10+Abs(CInt(first Is aliasRef))
End Function
Public Function ArrayProbe As Long
Dim value As Object
Set value=New NativeCounter
value.Cells(2)=17
ArrayProbe=value.Values(2)+value.Cells(2)
End Function
Public Function MethodProbe As String
Dim value As Object
Set value=New NativeCounter
value.Indexed(2)=20
MethodProbe=value.Label(prefix:="x") & ":" & CStr(value.Indexed(2))
End Function
Public Function ValueGuard As Long
Dim value As Object
Set value=New NativeCounter
On Error Resume Next
ValueGuard=value.Add(7)
ValueGuard=Err.Number*100+value.Calls
End Function
Public Function RefGuard As Long
Dim value As Object,amount As Long
Set value=New NativeCounter
On Error Resume Next
value.RequiresRef amount
RefGuard=Err.Number*100+value.Calls
End Function
Public Function IdentityProbe As Boolean
Dim value As Object
Set value=New NativeCounter
IdentityProbe=value Is value.Identity
End Function
)VBA";
