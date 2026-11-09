#pragma once
#include <string>
inline const std::string automaticClassState=R"VBA(Public autoConstructed As Long
Public moduleObject As New AutoTracked
)VBA";
inline const std::string automaticClassFixture=R"VBA(VERSION 1.0 CLASS
Attribute VB_Name = "AutoTracked"
Option Explicit
Private identifier As Long
Public Value As Long
Public Child As New AutoTracked
Private Sub Class_Initialize()
autoConstructed=autoConstructed+1
identifier=autoConstructed
Value=1
End Sub
Public Function OwnAutomaticProbe As Boolean
Dim first As Object
Set first=Me.Child
OwnAutomaticProbe=(first Is Child)
Set Me.Child=Nothing
OwnAutomaticProbe=OwnAutomaticProbe And (Not Child Is first)
End Function
Public Function Serial As Long
Serial=identifier
End Function
)VBA";
inline const std::string automaticClassConsumer=R"VBA(Option Explicit
Private Sub ClearAutomatic(ByRef item As AutoTracked)
Set item=Nothing
End Sub
Public Function AutoLocalProbe As String
Dim item As New AutoTracked, kept As Object, collection As New Collection, dictionary As New Scripting.Dictionary
AutoLocalProbe=CStr(autoConstructed)
Set item=Nothing
AutoLocalProbe=AutoLocalProbe & ":" & CStr(autoConstructed)
Set kept=item
AutoLocalProbe=AutoLocalProbe & ":" & CStr(item.Serial)
If Not kept Is item Then Err.Raise 5
Set item=Nothing
AutoLocalProbe=AutoLocalProbe & ":" & CStr(autoConstructed)
If item Is Nothing Then Err.Raise 5
AutoLocalProbe=AutoLocalProbe & ":" & CStr(item.Serial)
ClearAutomatic item
AutoLocalProbe=AutoLocalProbe & ":" & CStr(item.Serial)
collection.Add 4
Set collection=Nothing
If collection.Count<>0 Then Err.Raise 5
dictionary.Add "key",7
Set dictionary=Nothing
If dictionary.Exists("key") Then Err.Raise 5
End Function
Public Function AutoStaticProbe As Long
Static item As New AutoTracked
item.Value=item.Value+1
AutoStaticProbe=item.Value
End Function
Public Function AutoModuleProbe As Long
AutoModuleProbe=moduleObject.Serial
End Function
Public Sub ResetAutomaticModule
Set moduleObject=Nothing
End Sub
Public Function AutoOwnedProbe As Boolean
Dim item As New AutoTracked
AutoOwnedProbe=item.OwnAutomaticProbe
End Function
Public Function AutoDefaultProbe As Long
Dim item As New NativeDefault
item=11
AutoDefaultProbe=item.Content
End Function
)VBA";
