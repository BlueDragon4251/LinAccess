#pragma once
#include <string>
inline const std::string enumTypeFixture=R"VBA(Option Explicit
Option Base 1
Private Enum NativeState_2
First=2
Second=4
End Enum
Public enumField As NativeState_2
Public enumValues() As NativeState_2
Public Const TypedEnumConstant As NativeState_2=5
Public Function EnumIdentity(ByVal value As NativeState_2) As NativeState_2
EnumIdentity=value
End Function
Private Sub BumpLong(ByRef value As Long)
value=value+1
End Sub
Private Sub BumpEnum(ByRef value As NativeState_2)
value=value+1
End Sub
Private Function OptionalEnum(Optional ByVal value As NativeState_2=First) As NativeState_2
OptionalEnum=value
End Function
Public Property Get CurrentEnum As NativeState_2
CurrentEnum=enumField
End Property
Public Property Let CurrentEnum(ByVal value As NativeState_2)
enumField=value
End Property
Public Function EnumArrayResult As NativeState_2()
Dim items() As NativeState_2
ReDim items(-1 To 0)
items(-1)=2
items(0)=9
EnumArrayResult=items
End Function
Public Function EnumTypeProbe As Long
Dim value As NativeState_2, longValue As Long, fixed(1 To 2) As NativeState_2, returned As Variant, item As Variant, total As Long
Const LocalTyped As NativeState_2=7
value=7
BumpLong value
longValue=9
BumpEnum longValue
CurrentEnum=value
fixed(1)=TypedEnumConstant
fixed(2)=LocalTyped
ReDim enumValues(-1 To 0)
enumValues(-1)=First
enumValues(0)=Second
ReDim Preserve enumValues(-1 To 1)
If enumValues(1)<>0 Then Err.Raise 5
returned=EnumArrayResult()
For Each item In returned
total=total+item
Next
If VarType(value)<>3 Or TypeName(value)<>"Long" Then Err.Raise 5
If VarType(enumValues)<>8195 Or TypeName(enumValues)<>"Long()" Then Err.Raise 5
If EnumIdentity(1234)<>1234 Then Err.Raise 5
If True Then Dim inlineValue As NativeState_2: inlineValue=3
If True Then Dim inlineArray(2) As NativeState_2: inlineArray(1)=First
If LBound(inlineArray)<>1 Or UBound(inlineArray)<>2 Then Err.Raise 5
If True Then Const InlineConstant As NativeState_2=First
If InlineConstant<>2 Or inlineValue<>3 Then Err.Raise 5
EnumTypeProbe=value+longValue+CurrentEnum+fixed(1)+fixed(2)+total+OptionalEnum()+EnumIdentity(7)+enumValues(-1)+enumValues(0)
End Function
Public Function EnumTypeOverflow As Long
Dim value As NativeState_2
On Error Resume Next
value=2147483648
EnumTypeOverflow=Err.Number
End Function
)VBA";

inline const std::string enumTypeSqlFixture=R"VBA(Private Enum SqlState_2
FirstSqlState=2
End Enum
Public Function EnumSqlIdentity(ByVal value As SqlState_2) As SqlState_2
EnumSqlIdentity=value
End Function
)VBA";
