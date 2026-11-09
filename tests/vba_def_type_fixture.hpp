#pragma once
#include <string>
inline const std::string nativeDefTypeFixture=R"VBA(DefBool B
DefByte X
DefCur C
DefDate D
DefDbl F
DefInt I
DefLng L,A
DefObj O
DefSng R
DefStr S
DefVar V
Private latent,sModule,oModule,aModule(),sCellStored
Public Function TypeProbe As String
Dim integerValue,longValue,stringValue,boolValue,xValue
Dim currencyValue,dateValue,floatValue,realValue,objectValue,variantValue
TypeProbe=CStr(VarType(integerValue)) & ":" & CStr(VarType(longValue)) & ":" & CStr(VarType(stringValue)) & ":" & CStr(VarType(boolValue)) & ":" & CStr(VarType(xValue))
TypeProbe=TypeProbe & ":" & CStr(VarType(currencyValue)) & ":" & CStr(VarType(dateValue)) & ":" & CStr(VarType(floatValue)) & ":" & CStr(VarType(realValue)) & ":" & CStr(VarType(objectValue)) & ":" & CStr(VarType(variantValue))
End Function
Public Property Get SCell
SCell=sCellStored
End Property
Public Property Let SCell(ByVal sample)
sCellStored=sample
End Property
Public Function PropertyProbe As String
SCell="cell"
PropertyProbe=SCell
End Function
Public Function XTypes As String
XTypes=CStr(VarType(xValue)) & ":" & CStr(VarType(iValue)) & ":" & CStr(VarType(sValue)) & ":" & CStr(VarType(dValue)) & ":" & CStr(VarType(vValue))
End Function
Public Function SValue
SValue="typed"
End Function
Public Function LValue
LValue=7
End Function
Public Function OValue
If oModule Is Nothing Then Set oModule=CreateObject("Scripting.Dictionary")
Set OValue=oModule
End Function
Public Function Parameters(ByVal sample,Optional ByVal index) As String
Parameters=sample & ":" & CStr(index) & ":" & CStr(VarType(index))
End Function
Public Function ParameterProbe As String
ParameterProbe=Parameters("value")
End Function
Public Function ArrayProbe As Long
Dim longs(1 To 2)
longs(1)=12
longs(2)=13
aModule=longs
ArrayProbe=VarType(aModule)*100+UBound(aModule)*10+aModule(2)
End Function
Public Static Function LStatic
local=local+1
LStatic=local
End Function
Public Function LOrdinary
local=local+1
LOrdinary=local
End Function
Public Sub Bump(ByRef index)
index=index+1
latent=latent+1
End Sub
Public Function MismatchProbe As Long
Dim index As Long
On Error Resume Next
Bump index
MismatchProbe=Err.Number*100+latent
End Function
Public Function CurrencyLoop As Currency
For cCounter=CCur("0.0001") To CCur("0.0003") Step CCur("0.0001")
CurrencyLoop=CurrencyLoop+cCounter
Next
End Function
Public Function ImplicitArrayProbe As Long
ReDim aImplicit(1 To 2)
aImplicit(2)=9
ImplicitArrayProbe=VarType(aImplicit)*100+aImplicit(2)
End Function
Public Function OverrideProbe As Long
Dim integerValue As String
Const iLiteral=1.25
OverrideProbe=VarType(integerValue)*100+VarType(iLiteral)
End Function
Public Function ModuleProbe As String
latent=latent+1
sModule=sModule & "x"
ModuleProbe=CStr(latent) & ":" & sModule & ":" & CStr(VarType(latent)) & ":" & CStr(VarType(sModule))
End Function
Public Function OverflowProbe As Long
On Error Resume Next
iLocal=32768
OverflowProbe=Err.Number*100+VarType(iLocal)
End Function
)VBA";
