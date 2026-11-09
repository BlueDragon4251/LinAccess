#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_array.hpp"
#include <map>
#include <iostream>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
namespace {
class FixedStringHost final:public linaccess::VbaHost {
public:
    std::map<std::string,std::optional<std::string>> fields;std::vector<std::string> messages;
    std::optional<std::string> field(std::string_view name) override {return fields.at(std::string(name));}
    void setField(std::string_view name,std::optional<std::string> value) override {fields[std::string(name)]=std::move(value);}
    void message(std::string_view value) override {messages.emplace_back(value);}
    void saveRecord() override {assert(false);}
};
}
void testVbaFixedStrings()
    {
        FixedStringHost host;
        const std::string fixed=R"VBA(Const FixedWidth=4
Public FixedShared As String * FixedWidth
Sub FixedStrings()
Dim local As String * 5, other As String * FixedWidth
Static saved As String * 3
If Len(local)<>5 Or Len(FixedShared)<>4 Then Err.Raise 7101
If AscW(Left$(local,1))<>0 Then Err.Raise 7102
local="ab"
If local<>"ab   " Or LenB(local)<>10 Then Err.Raise 7103
local="abcdefg"
If local<>"abcde" Then Err.Raise 7104
other=12
If other<>"12  " Or TypeName(other)<>"String" Or VarType(other)<>8 Then Err.Raise 7105
local="ä"
If local<>"ä    " Then Err.Raise 7106
FixedShared="xy"
If FixedShared<>"xy  " Then Err.Raise 7107
LSet local="z"
If local<>"z    " Then Err.Raise 7108
RSet local="q"
If local<>"    q" Then Err.Raise 7109
Mid$(local,2,2)="ab"
If local<>" ab q" Then Err.Raise 7110
If AscW(Left$(saved,1))=0 Then saved=""
saved=Trim$(saved) & "x"
Me.FixedSaved=saved
Me.FixedValue=other
Dim failureCode As Long
On Error Resume Next
local=Null
failureCode=Err.Number
On Error GoTo 0
If failureCode<>94 Then Err.Raise 7111
If local<>" ab q" Then Err.Raise 7112
If FixedArgumentSize(local)<>5 Then Err.Raise 7113
End Sub
Function FixedArgumentSize(ByVal value As String) As Long
FixedArgumentSize=Len(value)
End Function
)VBA";
        const auto state=linaccess::createVbaModuleState(fixed);assert(state->types.at("fixedshared")=="string*4" && state->values.at("fixedshared").text==std::string(4,'\0'));
        auto result=linaccess::executeVbaProcedure(fixed,"FixedStrings",host);
        if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Fixed String fixture: "<<result.errorNumber<<"/"<<result.line<<" "<<result.explanation<<'\n';
        assert(result.status==linaccess::VbaExecutionStatus::Completed && host.fields["FixedValue"]=="12  ");
        assert(host.fields["FixedSaved"]=="x  ");
        assert(linaccess::executeVbaProcedure(fixed,"FixedStrings",host).status==linaccess::VbaExecutionStatus::Completed && host.fields["FixedSaved"]=="xx ");
        const auto forward=linaccess::createVbaModuleState("Public item As String * Width\nConst Width=3\nSub ReadWidth()\nEnd Sub\n");assert(forward->values.at("item").text==std::string(3,'\0'));
        assert(linaccess::coerceVbaModuleValue(linaccess::VbaValue::string("x"),"string*65526").text.size()==65526);
        for(const auto& invalid:std::vector<std::pair<std::string,int>>{{"0",5},{"65527",5},{"MissingWidth",3251}}) {
            const auto code="Sub InvalidFixed()\nMsgBox \"wrong\"\nDim s As String * "+invalid.first+"\nEnd Sub";
            const auto parsed=linaccess::executeVbaProcedure(code,"InvalidFixed",host);assert(parsed.status==linaccess::VbaExecutionStatus::Unsupported && parsed.errorNumber==invalid.second && host.messages.empty());
        }
        assert(linaccess::coerceVbaModuleValue(linaccess::VbaValue::string("😀"),"string*2").text=="😀");
        int splitError=0;try{(void)linaccess::coerceVbaModuleValue(linaccess::VbaValue::string("😀"),"string*1");}catch(const linaccess::VbaRuntimeError& error){splitError=error.number;}assert(splitError==3251);
        FixedStringHost arrays;
        const std::string arrayCode=R"VBA(Const Width=3
Public sharedCells(1 To 2) As String * Width
Sub FixedArrays()
Dim cells() As String * Width
Dim fixedCells(0 To 1) As String * 4
Static history(1 To 1) As String * 2
Dim copy As Variant, typed As Variant, item As Variant
Dim code As Long, total As Long
ReDim cells(1 To 2)
If Len(cells(1))<>3 Or AscW(Left$(cells(1),1))<>0 Then Err.Raise 7201
cells(1)="ab"
cells(2)="abcdef"
If cells(1)<>"ab " Or cells(2)<>"abc" Then Err.Raise 7202
ReDim Preserve cells(1 To 3)
If cells(1)<>"ab " Or cells(2)<>"abc" Then Err.Raise 7203
If Len(cells(3))<>3 Or AscW(Left$(cells(3),1))<>0 Then Err.Raise 7204
If VarType(cells)<>8200 Or TypeName(cells)<>"String()" Then Err.Raise 7205
copy=cells
copy(1)="x"
If copy(1)<>"x  " Or cells(1)<>"ab " Then Err.Raise 7206
For Each item In cells
total=total+Len(item)
Next
If total<>9 Then Err.Raise 7207
fixedCells(0)="q"
Erase fixedCells
If LBound(fixedCells)<>0 Or UBound(fixedCells)<>1 Or AscW(Left$(fixedCells(0),1))<>0 Then Err.Raise 7208
If Len(sharedCells(1))<>3 Then Err.Raise 7209
sharedCells(1)="z"
If sharedCells(1)<>"z  " Then Err.Raise 7210
Erase sharedCells
If AscW(Left$(sharedCells(1),1))<>0 Then Err.Raise 7211
ReDim typed(0 To 1) As String * 4
typed(0)="a"
ReDim Preserve typed(0 To 2)
If typed(0)<>"a   " Or Len(typed(2))<>4 Or AscW(Left$(typed(2),1))<>0 Then Err.Raise 7212
Erase cells
On Error Resume Next
code=LBound(cells)
code=Err.Number
On Error GoTo 0
If code<>9 Then Err.Raise 7213
ReDim cells(1 To 1)
cells(1)="ok"
On Error Resume Next
ReDim Preserve cells(1 To 2) As String * 4
code=Err.Number
On Error GoTo 0
If code<>13 Or UBound(cells)<>1 Or cells(1)<>"ok " Then Err.Raise 7214
On Error Resume Next
cells(1)=Null
code=Err.Number
On Error GoTo 0
If code<>94 Or cells(1)<>"ok " Then Err.Raise 7215
If AscW(Left$(history(1),1))=0 Then history(1)=""
history(1)=Trim$(history(1)) & "x"
Me.ArrayStatic=history(1)
Me.ArrayPassed="yes"
End Sub
)VBA";
        auto arrayResult=linaccess::executeVbaProcedure(arrayCode,"FixedArrays",arrays);
        if(arrayResult.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Fixed String array fixture: "<<arrayResult.errorNumber<<"/"<<arrayResult.line<<" "<<arrayResult.explanation<<'\n';
        assert(arrayResult.status==linaccess::VbaExecutionStatus::Completed && arrays.fields["ArrayPassed"]=="yes" && arrays.fields["ArrayStatic"]=="x ");
        assert(linaccess::executeVbaProcedure(arrayCode,"FixedArrays",arrays).status==linaccess::VbaExecutionStatus::Completed && arrays.fields["ArrayStatic"]=="xx");
        const auto wide=linaccess::executeVbaProcedure("Sub TooWide()\nMsgBox \"wrong\"\nDim cells(0 To 1000) As String * 65526\nEnd Sub","TooWide",arrays);
        assert(wide.status!=linaccess::VbaExecutionStatus::Completed && wide.errorNumber==3251 && arrays.messages.empty());
    }
