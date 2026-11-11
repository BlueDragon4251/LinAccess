#pragma once
#include "vba_enum_type_fixture.hpp"
inline const std::string projectEnumProvider=R"VBA(Attribute VB_Name = "EnumProvider"
Option Explicit
Public Enum NativeShared_2
SharedFirst=2
SharedSecond=4
End Enum
Private Enum NativeHidden_2
HiddenFirst=7
End Enum
Global Enum LegacyShared_2
LegacyFirst=3
End Enum
)VBA";
inline const std::string projectEnumConsumer=[] {
    auto source=enumTypeFixture;const auto begin=source.find("Private Enum NativeState_2");const auto end=source.find("End Enum",begin)+8;
    source.replace(begin,end-begin,"Private Const First As Long=2\nPrivate Const Second As Long=4");
    std::size_t at=0;while((at=source.find("NativeState_2",at))!=std::string::npos) {source.replace(at,std::string("NativeState_2").size(),"NativeShared_2");at+=std::string("NativeShared_2").size();}
    return std::string("Attribute VB_Name = \"EnumConsumer\"\n")+source;
}();
inline const std::string projectEnumQualifiedConsumer=[] {
    auto source=projectEnumConsumer;const std::string from="As NativeShared_2",to="As EnumProvider.NativeShared_2";
    std::size_t at=0;while((at=source.find(from,at))!=std::string::npos) {source.replace(at,from.size(),to);at+=to.size();}return source;
}();
inline const std::string projectEnumSqlConsumer=R"VBA(Attribute VB_Name = "ProjectEnumSqlScope"
Public Function ProjectEnumSql(ByVal value As EnumProvider.NativeShared_2) As NativeShared_2
ProjectEnumSql=value
End Function
Public Function ProjectLegacySql(ByVal value As LegacyShared_2) As EnumProvider.LegacyShared_2
ProjectLegacySql=value
End Function
Public Function ProjectClassEnumSql(ByVal value As ProjectEnumHolder.ClassShared_2) As ClassShared_2
ProjectClassEnumSql=value
End Function
)VBA";
inline const std::string projectEnumClassFixture=R"VBA(VERSION 1.0 CLASS
Attribute VB_Name = "ProjectEnumHolder"
Option Explicit
Public Enum ClassShared_2
ClassFirst=1
End Enum
Public Value As NativeShared_2
Public Property Get Current As EnumProvider.NativeShared_2
Current=Value
End Property
Public Property Let Current(ByVal nextValue As NativeShared_2)
Value=nextValue
End Property
)VBA";
inline const std::string projectEnumClassConsumer=R"VBA(Function ProjectEnumClassProbe As Long
Dim item As ProjectEnumHolder
Set item=New ProjectEnumHolder
item.Current=9.5
ProjectEnumClassProbe=item.Current+VarType(item.Value)
End Function)VBA";
