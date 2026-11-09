#pragma once
#include <string>
inline const std::string nativeSetTimingClass=R"VBA(VERSION 1.0 CLASS
Attribute VB_Name = "NativeSetBox"
Option Explicit
Public Held As Object
Public LastIndex As Long,Writes As Long,Reads As Long
Public Property Get Item(ByVal index As Long) As Object
Attribute Item.VB_UserMemId = 0
Reads=Reads+1
Set Item=Held
End Property
Public Property Set Item(ByVal index As Long,ByVal child As Object)
LastIndex=index
Set Held=child
Writes=Writes+1
End Property
)VBA";
inline const std::string nativeSetTimingFixture=R"VBA(Option Explicit
Private activeReceiver As NativeSetBox,oldReceiver As NativeSetBox,nextReceiver As NativeSetBox
Private sharedValues() As Object
Private order As String
Private chosen As Long
Public formHeld As Object
Public formIndex As Long,formWrites As Long
Private Sub ResetBoxes
Set oldReceiver=New NativeSetBox
Set nextReceiver=New NativeSetBox
Set activeReceiver=oldReceiver
chosen=0
order=""
End Sub
Private Function Receiver As NativeSetBox
order=order & "T"
Set Receiver=activeReceiver
End Function
Private Function NextIndex As Long
order=order & "I"
NextIndex=chosen
End Function
Private Function MakeSimpleChild As Object
Dim child As Collection
order=order & "R"
chosen=1
Set child=New Collection
child.Add 7
Set MakeSimpleChild=child
End Function
Private Function MakeMemberChild As Object
Set MakeMemberChild=MakeSimpleChild()
Set activeReceiver=nextReceiver
End Function
Private Function MakeArrayChild As Object
Set MakeArrayChild=MakeSimpleChild()
ReDim sharedValues(0 To 1)
End Function
Private Function BadChild As Object
order=order & "R"
Err.Raise 7012,"native.set.timing","timing failure"
End Function
Private Function MakeNullChild As Object
Set MakeNullChild=MakeSimpleChild()
Set activeReceiver=Nothing
End Function
Public Property Set FormSlot(ByVal index As Long,ByVal child As Object)
formIndex=index
Set formHeld=child
formWrites=formWrites+1
End Property
Public Function MemberSetTiming As Long
ResetBoxes
Set Receiver().Item(index:=NextIndex())=MakeMemberChild()
If order<>"RTI" Or oldReceiver.Writes<>0 Or nextReceiver.Reads<>0 Then Err.Raise 5
If Not oldReceiver.Held Is Nothing Then Err.Raise 5
MemberSetTiming=nextReceiver.LastIndex*100+nextReceiver.Writes*10+nextReceiver.Held.Count
End Function
Public Function DefaultSetTiming As Long
ResetBoxes
Set activeReceiver(index:=NextIndex())=MakeMemberChild()
If order<>"RI" Or oldReceiver.Writes<>0 Or nextReceiver.Reads<>0 Then Err.Raise 5
If Not oldReceiver.Held Is Nothing Then Err.Raise 5
DefaultSetTiming=nextReceiver.LastIndex*100+nextReceiver.Writes*10+nextReceiver.Held.Count
End Function
Public Function ArraySetTiming As Long
ReDim sharedValues(0 To 0)
Set sharedValues(0)=New Scripting.Dictionary
chosen=0
order=""
Set sharedValues(NextIndex())=MakeArrayChild()
If order<>"RI" Or UBound(sharedValues)<>1 Then Err.Raise 5
If Not sharedValues(0) Is Nothing Then Err.Raise 5
ArraySetTiming=UBound(sharedValues)*100+sharedValues(1).Count
End Function
Public Function OwnSetTiming As Long
chosen=0
order=""
formWrites=0
Set FormSlot(index:=NextIndex())=MakeSimpleChild()
If order<>"RI" Then Err.Raise 5
OwnSetTiming=formIndex*100+formWrites*10+formHeld.Count
End Function
Public Function FailedSetTiming As Long
ResetBoxes
On Error Resume Next
Set Receiver().Item(index:=NextIndex())=BadChild()
FailedSetTiming=Err.Number
On Error GoTo 0
If order<>"R" Or oldReceiver.Writes<>0 Or nextReceiver.Writes<>0 Then Err.Raise 5
End Function
Public Function NullSetTiming As Long
ResetBoxes
On Error Resume Next
Set activeReceiver.Item(index:=NextIndex())=MakeNullChild()
NullSetTiming=Err.Number
On Error GoTo 0
If order<>"R" Or oldReceiver.Writes<>0 Then Err.Raise 5
End Function
Public Function FormSetTiming As Long
chosen=0
order=""
formWrites=0
Set Me.FormSlot(index:=NextIndex())=MakeSimpleChild()
If order<>"RI" Then Err.Raise 5
FormSetTiming=formIndex*100+formWrites*10+formHeld.Count
End Function
Public Sub ClearTimingReferences
Set activeReceiver=Nothing
Set oldReceiver=Nothing
Set nextReceiver=Nothing
Set formHeld=Nothing
Erase sharedValues
End Sub
)VBA";
