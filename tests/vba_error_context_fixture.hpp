#pragma once
#include <string>
inline const std::string nativeErrorContextFixture=R"VBA(Option Explicit
Public retainedEnvironmentError As VBA.ErrObject
Private Sub SeedEnvironmentError(ByVal expected As VBA.ErrObject)
If Not Err Is expected Then Err.Raise 5
Err.Number=812
Err.Source="shared.source"
Err.Description="shared.description"
Err.HelpFile="shared.help"
Err.HelpContext=12
End Sub
Private Sub HandledEnvironmentError(ByVal expected As VBA.ErrObject)
If Not Err Is expected Then Err.Raise 5
On Error Resume Next
Err.Raise 813,"handled.source","handled.description"
End Sub
Private Sub UnhandledEnvironmentError
Err.Raise 814,"nested.source","nested.description"
End Sub
Public Function EnvironmentErrorProbe As Long
Dim observed As VBA.ErrObject, instance As NativeCounter
Set observed=Err
Err.Clear
SeedEnvironmentError observed
If observed.Number<>812 Or observed.Source<>"shared.source" Or observed.Description<>"shared.description" Or observed.HelpFile<>"shared.help" Or observed.HelpContext<>12 Then Err.Raise 5
On Error Resume Next
HandledEnvironmentError observed
If observed.Number<>813 Or observed.Source<>"handled.source" Then EnvironmentErrorProbe=-1: Exit Function
UnhandledEnvironmentError
If observed.Number<>814 Or observed.Source<>"nested.source" Or observed.Description<>"nested.description" Then EnvironmentErrorProbe=-2: Exit Function
On Error GoTo 0
Set instance=New NativeCounter
If Not instance.SameError(observed) Then Err.Raise 5
Set retainedEnvironmentError=Err
EnvironmentErrorProbe=17
End Function
Public Function EnvironmentIdentityProbe As Boolean
EnvironmentIdentityProbe=Err Is retainedEnvironmentError
End Function
Public Function ErrorVariantConversionProbe As Long
Dim value As Variant, code As Long
value=CVErr(CVErr(2015))
If Not IsError(value) Or VarType(value)<>10 Or CStr(value)<>"Error 2015" Then Err.Raise 5
value=CVErr(65535.49)
If CStr(value)<>"Error 65535" Then Err.Raise 5
value=CVErr(CCur("2.5"))
If CStr(value)<>"Error 2" Then Err.Raise 5
value=CVErr(CDec("3.5"))
If CStr(value)<>"Error 4" Then Err.Raise 5
value=CVErr(Empty)
If Not IsError(value) Or CStr(value)<>"Error 0" Then Err.Raise 5
On Error Resume Next
value=CVErr(65535.5)
code=Err.Number
Err.Clear
value=CVErr(2147483647.5)
code=code+Err.Number
Err.Clear
value=CVErr(Null)
code=code+Err.Number
Err.Clear
value=CVErr(Nothing)
code=code+Err.Number
ErrorVariantConversionProbe=code
End Function
Private Sub ExitWithRecentError
On Error Resume Next
Err.Raise 9
Exit Sub
End Sub
Public Function ErrorTextProbe As Long
Dim text As String, remembered As String, code As Long
If Error(0)<>"" Or Error$(0)<>"" Then Err.Raise 5
If VBA.Error(13)<>"Type mismatch" Or VBA.Error$(ErrorNumber:=9)<>"Subscript out of range" Then Err.Raise 5
If Error(60000)<>"Application-defined or object-defined error." Or Error(-1)<>Error(60000) Then Err.Raise 5
If Error(13.5)<>Error(14) Then Err.Raise 5
On Error Resume Next
Err.Raise 7011,"native.source","custom.description"
remembered=Error()
If remembered<>Error(7011) Then ErrorTextProbe=-1: Exit Function
Err.Number=13
If Error()<>remembered Then ErrorTextProbe=-2: Exit Function
Err.Clear
If Error$<>remembered Then ErrorTextProbe=-3: Exit Function
On Error GoTo 0
If VBA.Error<>remembered Then Err.Raise 5
On Error Resume Next
Err.Raise 11
If Error()<>Error(11) Or Err.Description<>Error(11) Then ErrorTextProbe=-4: Exit Function
On Error GoTo 0
ExitWithRecentError
If Err.Number<>0 Or Error()<>Error(9) Then Err.Raise 5
On Error Resume Next
text=Error(65536)
code=Err.Number
Err.Clear
text=Error(Null)
code=code+Err.Number
Err.Clear
text=Error(CVErr(13))
code=code+Err.Number
Err.Clear
text=Error(Nothing)
code=code+Err.Number
ErrorTextProbe=code
End Function
)VBA";
inline const std::string nativeFormErrorContextFixture=R"VBA(Option Explicit
Public Function FormEnvironmentError(ByVal expected As VBA.ErrObject) As Boolean
FormEnvironmentError=Err Is expected
End Function
)VBA";

inline const std::string nativeErrorShadowFixture=R"VBA(Option Explicit
Public Function Error As String
Error="own"
End Function
Public Function ErrorProcedureShadow As Long
If Error<>"own" Then Err.Raise 5
If VBA.Error(13)<>"Type mismatch" Then Err.Raise 5
ErrorProcedureShadow=17
End Function
Public Function ErrorArrayShadow As Long
Dim Error(0 To 0) As Long
Error(0)=19
If VBA.Error(0)<>"" Then Err.Raise 5
ErrorArrayShadow=Error(0)
End Function
)VBA";
