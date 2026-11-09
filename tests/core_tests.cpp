#include "csv.hpp"
#include "vba_currency.hpp"
#include "vba_decimal.hpp"
#include "native_report_aggregate.hpp"
#include <charconv>
#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_like.hpp"
#include "vba_array.hpp"

// Tests must stay active in Release builds (which normally define NDEBUG).
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <map>
#include <tuple>
#include <optional>
#include <limits>
#include <cmath>
#include <ctime>
#include <filesystem>
#include <fstream>
#include <cstdlib>
#include <chrono>

namespace {
class TestSequence final : public linaccess::VbaObject {
public:
    std::vector<linaccess::VbaValue> values;
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override {
        throw linaccess::VbaRuntimeError(438,"No members");
    }
    std::vector<linaccess::VbaValue> enumerate() override { return values; }
};
class LoopContext final : public linaccess::VbaObject {
public:
    explicit LoopContext(int number):number_(number) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        assert(name=="Value" && args.empty()); return linaccess::VbaValue::numeric(number_);
    }
private:
    int number_;
};
class ScalarProbe final : public linaccess::VbaObject {
public:
    linaccess::VbaValue value;
    int reads{};
    bool hasScalarDefault() const override { return true; }
    linaccess::VbaValue scalarValue() override { ++reads; return value; }
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No explicit members"); }
};
class FakeForm final : public linaccess::VbaHost {
public:
    std::map<std::string, std::optional<std::string>> fields;
    std::vector<std::string> messages;
    std::map<std::string,linaccess::VbaValue> bracketFields;
    bool bracketHost=false;
    linaccess::VbaValue bracketValue(std::string_view name,bool reference) override {
        if (!bracketHost || reference) return VbaHost::bracketValue(name,reference);
        return bracketFields.at(std::string(name));
    }
    bool dialogHost=false;
    int dialogResult=1,dialogButtons=-1;
    std::optional<std::string> dialogTitle;
    bool inputHost=false;
    std::optional<linaccess::VbaInputBoxArguments> inputRequest;
    std::string inputResult;
    std::optional<std::string> evalModule;
    std::optional<std::string> moduleForProcedure(std::string_view name) override {
        return evalModule ? evalModule : VbaHost::moduleForProcedure(name);
    }
    int saves = 0;
    int maximized = 0;
    int navigatedLast = 0;
    int navigatedNext = 0;
    int navigatedNew = 0;
    int requeries = 0;
    std::vector<std::string> sqlStatements;
    std::vector<std::string> openedForms;
    std::map<std::string, bool> visibility;
    int controlsRefreshed = 0;
    int controlsFocused = 0;
    std::optional<std::string> field(std::string_view name) override {
        auto found = fields.find(std::string(name));
        if (found == fields.end()) throw std::runtime_error("Unknown field");
        return found->second;
    }
    void setField(std::string_view name, std::optional<std::string> value) override {
        fields[std::string(name)] = std::move(value);
    }
    void message(std::string_view text) override { messages.emplace_back(text); }
    int messageBox(std::string_view text,int buttons,const std::optional<std::string>& title) override {
        if (!dialogHost) return VbaHost::messageBox(text,buttons,title);
        messages.emplace_back(text); dialogButtons=buttons; dialogTitle=title; return dialogResult;
    }
    void saveRecord() override { ++saves; }
    std::string inputBox(const linaccess::VbaInputBoxArguments& request) override {
        if (!inputHost) return VbaHost::inputBox(request);
        inputRequest=request; return inputResult;
    }
    bool rejectQuit=false;
    void quitApplication() override {
        if(rejectQuit) throw linaccess::VbaRuntimeError(2501,"Quit cancelled");
        linaccess::requestVbaExecutionStop();
    }
    void maximizeForm() override { ++maximized; }
    void gotoLastRecord() override { ++navigatedLast; }
    void gotoNextRecord() override { ++navigatedNext; }
    void gotoNewRecord() override { ++navigatedNew; }
    void requery() override { ++requeries; }
    void executeSql(std::string_view statement) override { sqlStatements.emplace_back(statement); }
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override {
        class SqlSink final:public linaccess::VbaObject {
        public:
            explicit SqlSink(FakeForm& owner):owner_(owner) {}
            linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& arguments) override {
                assert(name=="Execute" && arguments.size()==1);
                owner_.executeSql(arguments[0].asText()); return {};
            }
        private:
            FakeForm& owner_;
        };
        return std::make_shared<SqlSink>(*this);
    }
    void openForm(std::string_view name) override { openedForms.emplace_back(name); }
    void requeryControl(std::string_view) override { ++controlsRefreshed; }
    void focusControl(std::string_view) override { ++controlsFocused; }
    void setControlVisible(std::string_view name, bool value) override {
        visibility[std::string(name)] = value;
    }
};
}

void testByteValues() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    assert(eval("VarType(CByte(255))=vbByte").truthy() && eval("TypeName(CByte(0))").asText()=="Byte");
    assert(eval("CByte(254.5)").asNumber()==254 && eval("CByte(2.5)").asNumber()==2 && eval("CByte(3.5)").asNumber()==4);
    assert(eval("CByte(-0.5)").asNumber()==0 && eval("CByte(Empty)").asNumber()==0);
    assert(eval("CByte(True)").asNumber()==255 && eval("CByte(False)").asNumber()==0);
    assert(eval("CByte(CCur(12.5000))").asNumber()==12 && eval("CByte(\" 255 \" )").asNumber()==255);
    assert(eval("VarType(Abs(CByte(3)))=vbByte And VarType(Int(CByte(3)))=vbByte And VarType(Fix(CByte(3)))=vbByte").truthy());
    for(const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
        {"CByte(255.5)",6},{"CByte(-0.6)",6},{"CByte(Null)",94},{"CByte(CVErr(11))",13},
        {"CByte(\"invalid\")",13},{"CByte()",450},{"CByte(1,2)",450},{"CByte(Nothing)",91}}) {
        int actual=0; try { (void)eval(expression); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==code);
    }
    const auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::numeric(3.5);
    assert(evaluateVbaExpression("CByte(probe)",host,{{"probe",VbaValue::reference(probe)}}).asNumber()==4 && probe->reads==1);
    const std::string module=R"VBA(Public tracked As Byte
Function DefaultByte() As Byte
End Function
Function RoundByte(ByVal value As Byte) As Byte
RoundByte=value
End Function
Sub Increment(ByRef value As Byte)
value=value+1
End Sub
Function UpdateByte() As Byte
Dim value As Byte
If VarType(value)<>vbByte Then Err.Raise(5)
value=254
Call Increment(value)
tracked=value
UpdateByte=tracked
End Function
Function PreserveOverflow() As Byte
On Error Resume Next
tracked=256
Me!Code=CStr(Err.Number)
PreserveOverflow=tracked
End Function
Sub MismatchByte()
Dim value As Integer
value=4
Call Increment(value)
Me!Body="entered"
End Sub
Function TemporaryByte() As Integer
Dim value As Integer
value=4
Call Increment((value))
TemporaryByte=value
End Function
Sub ByteLoop()
Dim value As Byte
CurrentDb.Execute "INSERT before byte loop"
For value=1 To 2
CurrentDb.Execute "INSERT inside byte loop"
Next value
End Sub
)VBA";
    auto state=host.moduleState(module); assert(state->values.at("tracked").numericType==VbaValue::NumericType::Byte && state->values.at("tracked").number==0);
    auto result=executeVbaProcedure(module,"DefaultByte",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->numericType==VbaValue::NumericType::Byte && result.typedReturnValue->number==0);
    std::vector<VbaValue> arguments{VbaValue::numeric(254.5)};
    result=executeVbaProcedureWithArguments(module,"RoundByte",host,arguments); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==254 && result.typedReturnValue->numericType==VbaValue::NumericType::Byte);
    arguments[0]=VbaValue::numeric(256); result=executeVbaProcedureWithArguments(module,"RoundByte",host,arguments); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==6);
    result=executeVbaProcedure(module,"UpdateByte",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==255 && result.typedReturnValue->numericType==VbaValue::NumericType::Byte);
    result=executeVbaProcedure(module,"PreserveOverflow",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==255 && host.fields["Code"]=="6");
    result=executeVbaProcedure(module,"MismatchByte",host); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==13 && !host.fields.contains("Body"));
    result=executeVbaProcedure(module,"TemporaryByte",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==4);
    result=executeVbaProcedure(module,"ByteLoop",host); assert(result.status==VbaExecutionStatus::Completed && host.sqlStatements.size()==3);
}

void testTypedNot() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    assert(eval("Not Not CByte(3)").number==3 && eval("VarType(Not Not CByte(3))=vbByte").truthy());
    assert(eval("Not CByte(0)").number==255 && eval("Not CByte(255)").number==0);
    assert(eval("VarType(Not CByte(10))=vbByte").truthy() && eval("TypeName(Not CByte(10))").asText()=="Byte");
    assert(eval("Not CInt(-32768)").number==32767 && eval("VarType(Not CInt(1))=vbInteger").truthy());
    assert(eval("Not CLng(-2147483648)").number==2147483647 && eval("VarType(Not CLng(1))=vbLong").truthy());
    assert(eval("Not 10").number==-11 && eval("VarType(Not 10)=vbLong").truthy());
    assert(eval("Not CSng(2.5)").number==-3 && eval("VarType(Not CSng(2.5))=vbLong").truthy());
    assert(eval("Not CCur(3.5)").number==-5 && eval("VarType(Not Empty)=vbInteger").truthy());
    assert(eval("Not Empty").number==-1 && eval("Not Null").kind==VbaValue::Kind::Null);
    assert(eval("Not False").kind==VbaValue::Kind::Boolean && eval("Not False").boolean && !eval("Not True").boolean);
    assert(!eval("Not CByte(3)=CByte(3)").boolean && eval("Not 2^3").number==-9);
    const auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::byte(10);
    const auto value=evaluateVbaExpression("Not probe",host,{{"probe",VbaValue::reference(probe)}});
    assert(value.numericType==VbaValue::NumericType::Byte && value.number==245 && probe->reads==1);
    for(const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
        {"Not 1e100",6},{"Not \"invalid\"",13},{"Not CVErr(11)",13},{"Not Nothing",91}}) {
        int actual=0; try { (void)eval(expression); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==code);
    }
    const std::string module="Function Invert(ByVal value As Byte) As Byte\nInvert=Not value\nEnd Function\n";
    std::vector<VbaValue> args{VbaValue::byte(10)};
    const auto result=executeVbaProcedureWithArguments(module,"Invert",host,args);
    assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->numericType==VbaValue::NumericType::Byte && result.typedReturnValue->number==245);
}

void testDefaultTypeInspection() {
    using namespace linaccess;
    FakeForm host;
    const auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::byte(4);
    const VbaVariables variables{{"probe",VbaValue::reference(probe)}};
    const auto inspect=[&] { return evaluateVbaExpression("VarType(probe)",host,variables); };
    assert(inspect().number==17 && probe->reads==1);
    probe->value=VbaValue::string("synthetic"); assert(inspect().number==8 && probe->reads==2);
    probe->value={}; assert(inspect().number==1 && probe->reads==3);
    class BrokenDefault final : public VbaObject {
    public:
        int reads{};
        bool hasScalarDefault() const override { return true; }
        VbaValue scalarValue() override { ++reads; throw VbaRuntimeError(13,"Synthetic getter failure"); }
        VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw VbaRuntimeError(438,"No member"); }
    };
    const auto broken=std::make_shared<BrokenDefault>(); int actual=0;
    try { (void)evaluateVbaExpression("VarType(value)",host,{{"value",VbaValue::reference(broken)}}); } catch(const VbaRuntimeError& failure) { actual=failure.number; }
    assert(actual==13 && broken->reads==1);
    auto container=std::make_shared<TestSequence>();
    assert(evaluateVbaExpression("VarType(value)",host,{{"value",VbaValue::reference(container)}}).number==9);
    assert(evaluateVbaExpression("VarType(Nothing)",host,{}).number==9);
    const auto nested=std::make_shared<ScalarProbe>(); nested->value=VbaValue::reference(container);
    assert(evaluateVbaExpression("VarType(value)",host,{{"value",VbaValue::reference(nested)}}).number==9 && nested->reads==1);
    probe->value=VbaValue::reference(probe); const auto before=probe->reads;
    int error=0; try { (void)inspect(); } catch(const VbaRuntimeError& failure) { error=failure.number; }
    probe->value=VbaValue::empty(); assert(error==28 && probe->reads==before+16);
    const std::string code="Function ErrorType() As Integer\nOn Error Resume Next\nErr.Raise(11)\nErrorType=VarType(Err)\nEnd Function\nFunction ErrorNumberType() As Integer\nErrorNumberType=VarType(Err.Number)\nEnd Function";
    auto result=executeVbaProcedure(code,"ErrorType",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==3);
    result=executeVbaProcedure(code,"ErrorNumberType",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==3);
}

void testNativeMathFunctions() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    assert(eval("Choose(2.5,10,20,30)").asNumber()==20 && eval("Choose(0,10)").kind==VbaValue::Kind::Null);
    assert(eval("Choose(2,10,Null)").kind==VbaValue::Kind::Null);
    assert(eval("Switch(False,10,True,CCur(\"1.25\"),True,30)").kind==VbaValue::Kind::Currency);
    assert(eval("Switch(Null,10,False,20)").kind==VbaValue::Kind::Null);
    assert(eval("Val(String:=\"    1615 198th Street N.E.\")").asNumber()==1615198);
    assert(eval("Val(\"&HFFFF\")").asNumber()==-1 && eval("Val(\"&HFFFF&\")").asNumber()==65535);
    assert(eval("Val(\"&O177777\")").asNumber()==-1 && eval("Val(\"24 and 57\")").asNumber()==24);
    assert(eval("Val(\"+1.25e2tail\")").asNumber()==125 && eval("Val(\"1,25\")").asNumber()==1);
    assert(eval("Val(\"50%\")").asNumber()==50 && eval("Val(\"none\")").asNumber()==0);
    assert(eval("Sqr(Number:=81)").asNumber()==9 && eval("VarType(Sqr(81))=vbDouble").truthy());
    assert(std::abs(eval("Log(Exp(2))").asNumber()-2)<1e-12);
    assert(std::abs(eval("Sin(4*Atn(1)/2)").asNumber()-1)<1e-12);
    assert(std::abs(eval("Cos(4*Atn(1))").asNumber()+1)<1e-12);
    assert(std::abs(eval("Tan(Atn(1))").asNumber()-1)<1e-12);
    assert(eval("Exp(Empty)").asNumber()==1 && eval("Sqr(\"16\")").asNumber()==4);
    for(const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
        {"Choose(1,10,1/0)",11},{"Switch(True,10,False,1/0)",11},{"Switch(True,1,False)",450},
        {"Choose(Null,1)",94},{"Val(\"50.5%\")",13},{"Val(\"1e9999\")",6},{"Val(Null)",94},{"Val()",450},
        {"Val(\"&HFFFF!\")",3251},{"Val(\"&H100000000\")",6},{"Val(\"&H10000%\")",6},
        {"Sqr(-1)",5},{"Log(0)",5},{"Log(-1)",5},{"Sin(Null)",94},
        {"Exp(1000)",6},{"Cos()",450},{"Atn(1,2)",450},{"Sqr(\"invalid\")",13}}) {
        int actual=0; try { (void)eval(expression); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==code);
    }
}

void testCharacterFunctions() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    assert(eval("UCase$(String:=\"aBc 09!\")").asText()=="ABC 09!" && eval("LCase(\"ABC\")").asText()=="abc");
    assert(eval("LCase(Null)").kind==VbaValue::Kind::Null && eval("UCase(Empty)").asText().empty());
    assert(eval("Len(UCase(Chr(0) & \"ab\"))").asNumber()==3);
    assert(eval("Hex(Number:=459)").asText()=="1CB" && eval("Oct$(459)").asText()=="713");
    assert(eval("Hex$(-1)").asText()=="FFFFFFFF" && eval("Oct(-1)").asText()=="37777777777");
    assert(eval("Hex(Empty)").asText()=="0" && eval("Hex(10.5)").asText()=="A");
    assert(eval("Hex(CInt(-1))").asText()=="FFFF" && eval("Oct(CInt(-1))").asText()=="177777");
    assert(eval("Hex(CLng(-1))").asText()=="FFFFFFFF");
    assert(eval("Oct(Null)").kind==VbaValue::Kind::Null);
    assert(eval("Space(3)").asText()=="   " && eval("Space$(0)").asText().empty());
    assert(eval("String(Character:=\"Ωtail\",Number:=3)").asText()=="ΩΩΩ");
    assert(eval("String$(3,321)").asText()=="AAA" && eval("Len(String(3,0))").asNumber()==3);
    assert(eval("String(Null,65)").kind==VbaValue::Kind::Null && eval("String(2,Null)").kind==VbaValue::Kind::Null);
    assert(eval("Asc(\"Apple\")").number==65 && eval("VarType(Asc(\"A\"))=vbInteger").truthy());
    assert(eval("Asc(Chr(0))").number==0 && eval("Asc(Chr(127))").number==127);
    assert(eval("Chr(65.5)").asText()=="B" && eval("Chr(10)").asText()=="\n");
    assert(eval("ChrW(937)").asText()=="Ω" && eval("AscW(\"Ω\")").number==937);
    assert(eval("AscW(\"耀\")").number==-32768 && eval("AscW(\"😀\")").number==-10179);
    assert(eval("AscW(ChrW(-32768))").number==-32768 && eval("AscW(ChrW(-1))").number==-1);
    assert(eval("VarType(AscW(\"Ω\"))=vbInteger").truthy() && eval("TypeName(ChrW(937))").asText()=="String");
    assert(eval("Asc(12)").number==49 && eval("ChrW(\" 65 \")").asText()=="A");
    const auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::string("A");
    assert(evaluateVbaExpression("Asc(probe)",host,{{"probe",VbaValue::reference(probe)}}).number==65 && probe->reads==1);
    probe->value=VbaValue::numeric(937); assert(evaluateVbaExpression("ChrW(probe)",host,{{"probe",VbaValue::reference(probe)}}).asText()=="Ω" && probe->reads==2);
    for(const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
        {"UCase$(Null)",94},{"LCase()",450},{"UCase(\"ö\")",3251},
        {"Space(Null)",94},{"Space$(Null)",94},{"Hex$(Null)",94},{"Oct(2147483648)",6},{"Hex(1,2)",450},
        {"Space(-1)",5},{"Space(1000001)",3251},{"String(2,\"\")",5},{"String$(2,Null)",94},
        {"String(2,128)",3251},{"String(2)",450},{"Space(2147483648)",6},
        {"Asc(\"\")",5},{"AscW(Empty)",5},{"Asc(Null)",94},{"Chr(Null)",94},
        {"Asc(CVErr(11))",13},{"ChrW(\"invalid\")",13},{"Asc(Nothing)",91},
        {"Asc()",450},{"ChrW(1,2)",450},{"Asc(\"Ω\")",3251},{"Chr(128)",3251},
        {"ChrW(65536)",5},{"ChrW(-32769)",5},{"ChrW(2147483648)",6},
        {"ChrW(55357)",3251},{"ChrW(-10179)",3251}}) {
        int actual=0; try { (void)eval(expression); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==code);
    }
}

void testCurrencyForLoops() {
    using namespace linaccess;
    FakeForm host;
    const std::string module=R"VBA(Public upper As Currency
Public stride As Currency
Function CurrencyLoops() As String
Dim value As Currency
Dim count As Long
Dim total As Currency
upper=CCur("922337203685477.5806")
stride=CCur("0.0001")
For value=CCur("922337203685477.5804") To upper Step stride
count=count+1
upper=0
stride=100
Next value
CurrencyLoops=CStr(count)&":"&CStr(value)&":"&CStr(VarType(value))
count=0
For value=CCur("-922337203685477.5805") To CCur("-922337203685477.5807") Step CCur("-0.0001")
count=count+1
Next value
CurrencyLoops=CurrencyLoops&":"&CStr(count)&":"&CStr(value)
For value=CCur("0.0003") To CCur("0.0001") Step CCur("0.0001")
Me!Unexpected="entered"
Next value
CurrencyLoops=CurrencyLoops&":"&CStr(value)
For value=1 To 2 Step 0
Exit For
Next value
CurrencyLoops=CurrencyLoops&":"&CStr(value)
For value=CCur("0.0001") To CCur("0.0003") Step CCur("0.0001")
total=total+value
Next value
CurrencyLoops=CurrencyLoops&":"&CStr(total)&":"&CStr(value)
End Function
Function CurrencyLoopOverflow() As String
Dim value As Currency
On Error Resume Next
For value=CCur("922337203685477.5807") To CCur("922337203685477.5807") Step CCur("0.0001")
Next value
CurrencyLoopOverflow=CStr(Err.Number)&":"&CStr(value)
End Function
Function CurrencyLoopBodyWrite() As Currency
Dim value As Currency
For value=1 To 3
value=3
Next value
CurrencyLoopBodyWrite=value
End Function
)VBA";
    auto result=executeVbaProcedure(module,"CurrencyLoops",host);
    assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="3:922337203685477.5807:6:3:-922337203685477.5808:0.0003:1:0.0006:0.0004");
    assert(!host.fields.contains("Unexpected"));
    result=executeVbaProcedure(module,"CurrencyLoopOverflow",host);
    assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="6:922337203685477.5807");
    result=executeVbaProcedure(module,"CurrencyLoopBodyWrite",host);
    assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Currency && result.typedReturnValue->currencyUnits==40000);
}

void testNativeTypedForLoops() {
    using namespace linaccess;FakeForm host;
    const std::string module=R"VBA(Public boundary As Double
Public stride As Double
Function ByteCounters() As String
Dim value As Byte
Dim count As Long
boundary=3
stride=1
For value=1 To boundary Step stride
count=count+1
boundary=100
stride=0
Next value
ByteCounters=CStr(count)&":"&CStr(value)&":"&TypeName(value)
count=0
For value=3 To 1 Step -1
count=count+1
Next value
ByteCounters=ByteCounters&":"&CStr(count)&":"&CStr(value)
For value=0 To 1 Step 0
Exit For
Next value
ByteCounters=ByteCounters&":"&CStr(value)
End Function
Function ByteCounterOverflow() As String
Dim value As Byte
On Error GoTo Failed
For value=255 To 255
Next value
Failed:
ByteCounterOverflow=CStr(Err.Number)&":"&CStr(value)&":"&TypeName(value)
End Function
Function IntegralBounds() As String
Dim value As Integer
Dim count As Long
For value=0.5 To 0.25
count=count+1
Next value
IntegralBounds=CStr(count)&":"&CStr(value)
count=0
For value=0 To 2.5
count=count+1
Next value
IntegralBounds=IntegralBounds&":"&CStr(count)&":"&CStr(value)
For value=32766 To 40000
Exit For
Next value
IntegralBounds=IntegralBounds&":"&CStr(value)
End Function
Function SingleCounters() As String
Dim value As Single
Dim count As Long
For value=0 To 0.3 Step 0.1
count=count+1
Next value
SingleCounters=CStr(count)&":"&TypeName(value)
count=0
For value=0 To CSng(0.3) Step CSng(0.1)
count=count+1
Next value
SingleCounters=SingleCounters&":"&CStr(count)
For value=5 To 2
Me!UnexpectedSingle="entered"
Next value
SingleCounters=SingleCounters&":"&CStr(value)
For value=2 To 3
value=3
Next value
SingleCounters=SingleCounters&":"&CStr(value)
End Function
Function SingleCounterOverflow() As String
Dim value As Single
On Error GoTo Failed
For value=CSng(3.4028234663852886E38) To 1E39 Step CSng(3.4E38)
Next value
Failed:
SingleCounterOverflow=CStr(Err.Number)&":"&TypeName(value)&":"&CStr(value>0)
End Function
Function DoubleZeroStep() As Double
Dim value As Double
For value=1 To 2 Step 0
Exit For
Next value
DoubleZeroStep=value
End Function
Sub RejectDateCounter()
Dim value As Date
CurrentDb.Execute "INSERT forbidden date loop prefix"
For value=1 To 2
Next value
End Sub
)VBA";
    auto result=executeVbaProcedure(module,"ByteCounters",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="3:4:Byte:3:0:0");
    result=executeVbaProcedure(module,"ByteCounterOverflow",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="6:255:Byte");
    result=executeVbaProcedure(module,"IntegralBounds",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="1:1:3:3:32766");
    result=executeVbaProcedure(module,"SingleCounters",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="3:Single:4:5:4" && !host.fields.contains("UnexpectedSingle"));
    result=executeVbaProcedure(module,"SingleCounterOverflow",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="6:Single:True");
    result=executeVbaProcedure(module,"DoubleZeroStep",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==1);
    result=executeVbaProcedure(module,"RejectDateCounter",host);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251 && host.sqlStatements.empty());
}

void testCurrencyRealStorage() {
    using namespace linaccess;
    constexpr std::int64_t limit=5497558138880000LL;
    const auto check=[](std::int64_t units) {
        const auto exact=VbaValue::currency(units); const auto text=exact.asText();
        double reference=0; const auto parsed=std::from_chars(text.data(),text.data()+text.size(),reference);
        assert(parsed.ec==std::errc{} && parsed.ptr==text.data()+text.size());
        assert(currencyToRealStorage(exact)==reference && currencyFromRealStorage(reference).currencyUnits==units);
    };
    for(const auto units:std::vector<std::int64_t>{0,1,-1,9999,-9999,10000,-10000,12345,-12345,limit-1,-limit+1}) check(units);
    for(int power=0;power<53;++power) for(int offset=-3;offset<=3;++offset) {
        const auto units=(std::int64_t(1)<<power)+offset;
        if(units>=0 && units<limit) { check(units); check(-units); }
    }
    std::uint64_t state=1234567;
    for(int i=0;i<10000;++i) { state=state*6364136223846793005ULL+1442695040888963407ULL; const auto units=static_cast<std::int64_t>(state%static_cast<std::uint64_t>(limit)); check(i%2 ? units : -units); }
    assert(currencyToRealStorage(VbaValue::string("1.23455"))==1.2346);
    assert(currencyToRealStorage(VbaValue::empty())==0 && currencyToRealStorage(VbaValue::logical(true))==-1);
    const auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::currency(12345);
    assert(currencyToRealStorage(VbaValue::reference(probe))==1.2345 && probe->reads==1);
    for(const auto& [value,expected]:std::vector<std::pair<double,int>>{{0.1+0.2,13},{1.234567,13},{0.00005,13},{std::numeric_limits<double>::denorm_min(),13},{549755813888.0,3251},{-549755813888.0,3251},{std::numeric_limits<double>::infinity(),13},{std::numeric_limits<double>::quiet_NaN(),13}}) {
        int error=0; try { (void)currencyFromRealStorage(value); } catch(const VbaRuntimeError& failure) { error=failure.number; } assert(error==expected);
    }
    for(const auto& [value,expected]:std::vector<std::pair<VbaValue,int>>{{VbaValue::currency(limit),3251},{VbaValue::currency(-limit),3251},{VbaValue::currency(std::numeric_limits<std::int64_t>::max()),3251},{VbaValue::currency(std::numeric_limits<std::int64_t>::min()),3251},{VbaValue{},94},{VbaValue::errorValue(11),13},{VbaValue::reference(nullptr),91}}) {
        int error=0; try { (void)currencyToRealStorage(value); } catch(const VbaRuntimeError& failure) { error=failure.number; } assert(error==expected);
    }
}

void testIsObjectFunction() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    for(const auto* expression:{"IsObject(Nothing)","Not IsObject(Empty)","Not IsObject(Null)","Not IsObject(42)",R"VBA(Not IsObject("text"))VBA","Not IsObject(True)","Not IsObject(CVErr(11))"}) {
        const auto value=eval(expression); assert(value.kind==VbaValue::Kind::Boolean && value.boolean);
    }
    const auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::numeric(42);
    assert(evaluateVbaExpression("IsObject(probe)",host,{{"probe",VbaValue::reference(probe)}}).boolean && probe->reads==0);
    assert(!evaluateVbaExpression("IsObject(probe+0)",host,{{"probe",VbaValue::reference(probe)}}).boolean && probe->reads==1);
    assert(!evaluateVbaExpression("IsObject(bytes)",host,{{"bytes",VbaValue::byteArray({1,2,3})}}).boolean);
    assert(eval("TypeName(IsObject(Nothing))").asText()=="Boolean");
    for(const auto& [expression,code]:std::vector<std::pair<std::string,int>>{{"IsObject()",450},{"IsObject(1,2)",450},{"IsObject(probe.Value)",438}}) {
        int actual=0; try { (void)evaluateVbaExpression(expression,host,{{"probe",VbaValue::reference(probe)}}); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==code);
    }
}

void testStringResultTypes() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    for(const auto* expression:{R"VBA(Len("A😀b"))VBA",R"VBA(InStr("A😀b","b"))VBA",R"VBA(InStr("abc","z"))VBA",R"VBA(InStr(2,"abc",""))VBA"}) {
        const auto result=eval(expression); assert(result.kind==VbaValue::Kind::Number && result.numericType==VbaValue::NumericType::Long);
    }
    assert(eval(R"VBA(Len("A😀b"))VBA").number==4 && eval(R"VBA(InStr("A😀b","b"))VBA").number==4);
    assert(eval(R"VBA(InStr(2147483647,"a",""))VBA").number==2147483647);
    for(const auto* expression:{R"VBA(StrComp("a","b"))VBA",R"VBA(StrComp("b","a"))VBA",R"VBA(StrComp("a","a"))VBA"}) {
        const auto result=eval(expression); assert(result.kind==VbaValue::Kind::Number && result.numericType==VbaValue::NumericType::Integer);
    }
    assert(eval(R"VBA(VarType(Len("a"))=vbLong)VBA").truthy());
    assert(eval(R"VBA(TypeName(InStr("a","a")))VBA").asText()=="Long");
    assert(eval(R"VBA(VarType(StrComp("a","b"))=vbInteger)VBA").truthy());
    assert(eval(R"VBA(Len(Null))VBA").kind==VbaValue::Kind::Null && eval(R"VBA(InStr(Null,"a"))VBA").kind==VbaValue::Kind::Null && eval(R"VBA(StrComp("a",Null))VBA").kind==VbaValue::Kind::Null);
}

void testReplaceFunction() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    for(const auto& [expression,expected]:std::vector<std::pair<std::string,std::string>>{
        {R"VBA(Replace("red/red/red","red","blue"))VBA","blue/blue/blue"},
        {R"VBA(Replace("red/red/red","red","blue",5))VBA","blue/blue"},
        {R"VBA(Replace("red/red/red","red","blue",5,1))VBA","blue/red"},
        {R"VBA(Replace("red/red/red","red","blue",5,0))VBA","red/red"},
        {R"VBA(Replace("red/red/red","","blue",3))VBA","d/red/red"},
        {R"VBA(Replace("red/red/red","red",""))VBA","//"},
        {R"VBA(Replace("aaaaa","aa","b"))VBA","bba"},
        {R"VBA(Replace("😀x😀","😀","Ω"))VBA","ΩxΩ"},
        {R"VBA(Replace("a-A","a","X",,,1))VBA","X-X"},
        {R"VBA(Replace("a-A","a","X",,,0))VBA","X-A"},
        {R"VBA(Replace("abc","b","x",99))VBA",""},
        {R"VBA(Replace("","b","x"))VBA",""},
        {R"VBA(Replace(123,2,4))VBA","143"}}) assert(eval(expression).asText()==expected);
    assert(eval(R"VBA(Replace("a" & Chr(0) & "a","a","x"))VBA").asText()==std::string("x\0x",3));
    assert(eval(R"VBA(VarType(Replace("a","a","x"))=vbString)VBA").truthy());
    assert(eval(R"VBA(Replace("abcd","b","X",2.5))VBA").asText()=="Xcd");
    assert(eval(R"VBA(Replace("aaaa","a","x",1,1.5))VBA").asText()=="xxaa");
    assert(eval(R"VBA(Replace("aaaa","a","x",1,Empty))VBA").asText()=="aaaa");
    const auto start=std::make_shared<ScalarProbe>(),count=std::make_shared<ScalarProbe>(),compare=std::make_shared<ScalarProbe>();
    start->value=VbaValue::integer(2); count->value=VbaValue::integer(1); compare->value=VbaValue::integer(1);
    assert(evaluateVbaExpression(R"VBA(Replace("aAa","a","x",start,count,compare))VBA",host,{{"start",VbaValue::reference(start)},{"count",VbaValue::reference(count)},{"compare",VbaValue::reference(compare)}}).asText()=="xa");
    assert(start->reads==1 && count->reads==1 && compare->reads==1);
    const auto text=std::make_shared<ScalarProbe>(),needle=std::make_shared<ScalarProbe>(),replacement=std::make_shared<ScalarProbe>();
    text->value=VbaValue::string("one-one"); needle->value=VbaValue::string("one"); replacement->value=VbaValue::string("two");
    assert(evaluateVbaExpression("Replace(text,needle,replacement)",host,{{"text",VbaValue::reference(text)},{"needle",VbaValue::reference(needle)},{"replacement",VbaValue::reference(replacement)}}).asText()=="two-two");
    assert(text->reads==1 && needle->reads==1 && replacement->reads==1);
    for(const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
        {R"VBA(Replace())VBA",450},{R"VBA(Replace("a",,"x"))VBA",450},
        {R"VBA(Replace(Null,"a","x"))VBA",94},{R"VBA(Replace("a",Null,"x"))VBA",94},
        {R"VBA(Replace("a","a",Null))VBA",94},{R"VBA(Replace("a","a","x",Null))VBA",94},
        {R"VBA(Replace("a","a","x",1,Null))VBA",94},{R"VBA(Replace("a","a","x",1,-1,Null))VBA",94},
        {R"VBA(Replace("a","a","x",0))VBA",5},{R"VBA(Replace("a","a","x",1,-2))VBA",5},
        {R"VBA(Replace("a","a","x",1,-1,3))VBA",5},{R"VBA(Replace("a","a","x",2147483648))VBA",6},
        {R"VBA(Replace(CVErr(11),"a","x"))VBA",13},{R"VBA(Replace(Nothing,"a","x"))VBA",91},
        {R"VBA(Replace("😀a","x","y",2))VBA",3251}}) {
        int actual=0; try { (void)eval(expression); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==code);
    }
}

void testRoundFunction() {
    using namespace linaccess;
    FakeForm host;
    const auto eval=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,{}); };
    assert(eval("Round(2.5)").number==2 && eval("Round(3.5)").number==4 && eval("Round(-2.5)").number==-2 && eval("Round(-3.5)").number==-4);
    for(const auto& sample:std::vector<std::pair<std::string,double>>{{"Round(0.12335,4)",0.1234},{"Round(0.12345,4)",0.1234},{"Round(0.12355,4)",0.1236},{"Round(0.12365,4)",0.1236},{"Round(0.00005,4)",0}})
        assert(std::abs(eval(sample.first).number-sample.second)<1e-12);
    assert(eval("Round(Null)").kind==VbaValue::Kind::Null && eval("Round(\" 2.5 \")").number==2);
    assert(eval("VarType(Round(CInt(1)))=vbInteger").truthy() && eval("VarType(Round(CLng(1)))=vbLong").truthy() && eval("VarType(Round(CByte(1)))=vbByte").truthy());
    assert(eval("Round(True)").number==-1 && eval("VarType(Round(Empty))=vbInteger").truthy());
    const auto positive=eval("Round(CCur(\"1.25\"),1)"),negative=eval("Round(CCur(\"-1.35\"),1)");
    assert(positive.kind==VbaValue::Kind::Currency && positive.currencyUnits==12000 && negative.currencyUnits==-14000);
    assert(eval("Round(CCur(\"922337203685477.0001\"),0)").asText()=="922337203685477");
    assert(eval("Round(CCur(\"922337203685477.1234\"),3)").asText()=="922337203685477.123");
    assert(eval("Round(2.5,)").number==2);
    assert(eval("Round(CCur(\"922337203685477.5807\"),4)").currencyUnits==std::numeric_limits<std::int64_t>::max());
    const auto number=std::make_shared<ScalarProbe>(),precision=std::make_shared<ScalarProbe>(); number->value=VbaValue::numeric(2.5); precision->value=VbaValue::integer(0);
    assert(evaluateVbaExpression("Round(number,precision)",host,{{"number",VbaValue::reference(number)},{"precision",VbaValue::reference(precision)}}).number==2 && number->reads==1 && precision->reads==1);
    for(const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
        {"Round()",450},{"Round(,1)",450},{"Round(1,2,3)",450},{"Round(1,-1)",5},{"Round(1,Null)",94},
        {"Round(CVErr(11))",13},{"Round(\"invalid\")",13},{"Round(Nothing)",91},
        {"Round(1,16)",3251},{"Round(CSng(1.5))",3251},{"Round(1e300,15)",3251},
        {"Round(CCur(\"922337203685477.5807\"),0)",6},{"Round(CCur(\"-922337203685477.5808\"),3)",6}}) {
        int actual=0; try { (void)eval(expression); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==code);
    }
}

void testNativeCollection() {
    using namespace linaccess; FakeForm host;
    const auto result=executeVbaProcedure(R"VBA(Function CollectionFlow() As Currency
Dim values As Collection
Dim alias As Object
Dim entry As Variant
Dim total As Currency
Set values=New VBA.Collection
values.Add Item:=CCur("1.25"), Key:="first"
values.Add 3,"last"
values.Add 2,,"last"
values.Add Item:=CCur("1.5"), After:="FIRST"
values.Remove "last"
Me!CollectionOrder=CStr(values(1)) & ":" & CStr(values.Item("FIRST")) & ":" & CStr(values(2))
For Each entry In values
 total=total+entry
Next
Set alias=values
On Error Resume Next
values.Add 99,"FiRsT"
Me!CollectionDuplicate=Err.Number
Err.Clear
For Each entry In values
 values.Remove 1
 Me!CollectionMutation=Err.Number
 Exit For
Next
On Error GoTo 0
alias.Remove 1
Me!CollectionCount=values.Count
Me!CollectionType=TypeName(values)
CollectionFlow=total
End Function
)VBA","CollectionFlow",host);
    if(result.status!=VbaExecutionStatus::Completed) std::cerr<<"Collection regression status="<<static_cast<int>(result.status)<<" error="<<result.errorNumber<<" line="<<result.line<<"\n";
    assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Currency && result.typedReturnValue->currencyUnits==47500);
    assert(host.fields["CollectionOrder"]=="1.25:1.25:1.5" && host.fields["CollectionDuplicate"]=="457" && host.fields["CollectionMutation"]=="3251" && host.fields["CollectionCount"]=="2" && host.fields["CollectionType"]=="Collection");
    const std::string module=R"VBA(Private pending As Collection
Sub InitializeItems()
Set pending=New Collection
pending.Add 1
End Sub
Sub Append(ByRef items As Collection)
items.Add 2
End Sub
Sub ReplaceVal(ByVal items As Collection)
items.Add 4
Set items=New Collection
items.Add 5
End Sub
Sub ReplaceRef(ByRef items As Collection)
Set items=New Collection
items.Add 3
End Sub
Sub MutateItems()
Append pending
ReplaceVal pending
End Sub
Sub ReplaceItems()
ReplaceRef pending
End Sub
Function CountItems() As Long
CountItems=pending.Count
End Function
Function MakeItems() As Collection
Set MakeItems=New Collection
MakeItems.Add CCur("1.25")
End Function
Function OptionalCount(Optional ByVal items As Collection=Nothing) As Long
If IsMissing(items) Then Err.Raise 5
If items Is Nothing Then
OptionalCount=-1
Else
OptionalCount=items.Count
End If
End Function
Sub OptionalReplace(Optional ByRef items As Collection)
Set items=New Collection
items.Add 9
End Sub
Sub OptionalFlow()
Dim items As Collection
Set items=New Collection
items.Add 3
Me!OptionalProvided=OptionalCount(Items:=items)
Me!OptionalMissing=OptionalCount()
OptionalReplace items
Me!OptionalReplaced=items(1)
OptionalReplace
Me!OptionalAfterOmission=items(1)
End Sub
)VBA";
    assert(executeVbaProcedure(module,"InitializeItems",host).status==VbaExecutionStatus::Completed);
    assert(executeVbaProcedure(module,"MutateItems",host).status==VbaExecutionStatus::Completed);
    auto count=executeVbaProcedure(module,"CountItems",host); assert(count.status==VbaExecutionStatus::Completed && count.returnValue=="3");
    assert(executeVbaProcedure(module,"ReplaceItems",host).status==VbaExecutionStatus::Completed);
    count=executeVbaProcedure(module,"CountItems",host); assert(count.status==VbaExecutionStatus::Completed && count.returnValue=="1");
    assert(executeVbaProcedure(module,"OptionalFlow",host).status==VbaExecutionStatus::Completed);
    assert(host.fields["OptionalProvided"]=="1" && host.fields["OptionalMissing"]=="-1" && host.fields["OptionalReplaced"]=="9" && host.fields["OptionalAfterOmission"]=="9");
    const auto badDefault=executeVbaProcedure("Sub BadDefault(Optional items As Collection=1)\nMe!BadCollectionDefault=1\nEnd Sub","BadDefault",host);
    assert(badDefault.status==VbaExecutionStatus::Unsupported && !host.fields.contains("BadCollectionDefault"));
    const auto factory=executeVbaProcedure(module,"MakeItems",host);
    assert(factory.status==VbaExecutionStatus::Completed && factory.typedReturnValue && factory.typedReturnValue->kind==VbaValue::Kind::Object && factory.typedReturnValue->object->member("Item",{VbaValue::integer(1)}).currencyUnits==12500);
    const auto value=evaluateVbaObjectExpression("New Collection",host,{});
    value.object->member("Add",{VbaValue::dateSerial(2),VbaValue::string("date")});
    value.object->member("Add",{VbaValue{}});
    assert(value.object->member("Item",{VbaValue::string("DATE")}).kind==VbaValue::Kind::Date && value.object->member("Item",{VbaValue::integer(2)}).kind==VbaValue::Kind::Null);
    const auto before=value.object->member("Count",{}).number;
    for(const auto& [name,args,expected]:std::vector<std::tuple<std::string,std::vector<VbaValue>,int>>{
        {"Item",{VbaValue::integer(0)},5},{"Remove",{VbaValue::string("missing")},5},
        {"Add",{VbaValue::integer(1),VbaValue::missing(),VbaValue::integer(1),VbaValue::integer(2)},5},
        {"Add",{VbaValue::integer(1),VbaValue::string("é")},3251},{"Add",{value},3251},{"Add",{evaluateVbaExpression("Array(v)",host,{{"v",value}})},3251}}) {
        int actual=0; try { (void)value.object->member(name,args); } catch(const VbaRuntimeError& error) { actual=error.number; }
        assert(actual==expected && value.object->member("Count",{}).number==before);
    }
    const auto child=evaluateVbaObjectExpression("New Collection",host,{});
    child.object->member("Add",{VbaValue::integer(7)});
    value.object->member("Add",{child,VbaValue::string("child")});
    value.object->member("Add",{evaluateVbaExpression("Array(v)",host,{{"v",child}})});
    assert(value.object->member("Item",{VbaValue::string("child")}).object==child.object);
    for(const auto& item:std::vector<VbaValue>{value,evaluateVbaExpression("Array(v)",host,{{"v",value}})}) {
        int actual=0; try { child.object->member("Add",{item}); } catch(const VbaRuntimeError& error) { actual=error.number; }
        assert(actual==3251 && child.object->member("Count",{}).number==1);
    }
    const auto nested=executeVbaProcedure(R"VBA(Function NestedFlow() As Long
Dim groups As Collection
Dim child As Collection
Dim item As Variant
Dim total As Long
Set groups=New Collection
Set child=New Collection
child.Add 7
groups.Add child
For Each child In groups
child.Add 8
For Each item In child
total=total+item
Next
Next
NestedFlow=total
End Function
)VBA","NestedFlow",host);
    assert(nested.status==VbaExecutionStatus::Completed && nested.returnValue=="15");
}

void testCalendarFunctions() {
    using namespace linaccess; FakeForm host;
    const auto eval=[&](const std::string& source) { return evaluateVbaExpression(source,host,{}); };
    const auto beforeClock=std::time(nullptr);
    const auto timer=eval("VBA.Timer()");
    const auto afterClock=std::time(nullptr);
    assert(timer.kind==VbaValue::Kind::Number && timer.numericType==VbaValue::NumericType::Single && timer.number>=0 && timer.number<=86400);
    const auto nearLocal=[&](std::time_t instant) {
        std::tm local{}; assert(::localtime_r(&instant,&local));
        const auto difference=std::abs(timer.number-(local.tm_hour*3600+local.tm_min*60+local.tm_sec));
        return std::min(difference,std::abs(difference-86400))<2;
    };
    assert(nearLocal(beforeClock) || nearLocal(afterClock));
    assert(eval("VarType(Timer)=vbSingle And VarType(Timer())=vbSingle").truthy());
    int timerError=0; try { (void)eval("Timer(1)"); } catch(const VbaRuntimeError& error) { timerError=error.number; } assert(timerError==450);
    assert(eval("VarType(#2/29/2024#)=vbDate").truthy());
    assert(eval("CStr(#12/31/2003 12:34#)").asText()=="2003-12-31 12:34:00");
    assert(eval("DateDiff(\"d\",#2/28/2024#,#2024-03-01#)").asNumber()==2);
    const auto literalDefault=executeVbaProcedure("Function LiteralDate(Optional value As Date = #12/31/2003#) As Date\nLiteralDate=value\nEnd Function","LiteralDate",host);
    assert(literalDefault.status==VbaExecutionStatus::Completed && literalDefault.typedReturnValue->kind==VbaValue::Kind::Date && literalDefault.typedReturnValue->asText()=="2003-12-31");
    const auto chained=executeVbaProcedure("Sub ColonFlow()\nDim total As Long: total=0\nFor i=1 To 3: total=total+i: Next i\nIf total=6 Then total=total+1: total=total+1 Else total=0\nGoTo done: total=99\ndone: Me!ColonTotal=total\nMe!ColonComment=1: Rem ignored: Me!ColonBad=99\nIf False Then Rem ignored Else Me!ColonBad=99\nEnd If\nEnd Sub","ColonFlow",host);
    assert(chained.status==VbaExecutionStatus::Completed && host.fields["ColonTotal"]=="8" && host.fields["ColonComment"]=="1" && !host.fields.contains("ColonBad"));
    const auto literalBody=executeVbaProcedure("Sub LiteralBody()\nMe!LiteralStamp=#12/31/2003 12:34#: Me!LiteralClock=#14:24#\nIf True Then Me!LiteralNumber=1\nEnd Sub","LiteralBody",host);
    assert(literalBody.status==VbaExecutionStatus::Completed && host.fields["LiteralStamp"]=="2003-12-31 12:34:00" && host.fields["LiteralClock"]=="14:24:00" && host.fields["LiteralNumber"]=="1");
    const auto invalidLiteral=executeVbaProcedure("Sub InvalidDate()\nMe!LiteralEffect=1\nMe!Other=#2/30/2024#\nEnd Sub","InvalidDate",host);
    assert(invalidLiteral.status==VbaExecutionStatus::Unsupported && !host.fields.contains("LiteralEffect"));
    assert(eval("CStr(DateValue(Date:=\"2024-02-29 12:34:56\"))").asText()=="2024-02-29");
    assert(eval("CStr(TimeValue(Time:=\"2024-02-29T12:34\"))").asText()=="12:34:00");
    assert(eval("CStr(TimeValue(CDate(-1.25)))").asText()=="06:00:00" && eval("TimeValue(Null)").kind==VbaValue::Kind::Null);
    assert(eval("CStr(CDate(\"14:24\"))").asText()=="14:24:00");
    assert(eval("TypeName(CDate(2958465)-(-1))").asText()=="Double" && eval("CDate(2958465)-(-1)").asNumber()==2958466);
    assert(eval("CStr(DateAdd(Date:=\"2024-01-31\",Number:=1,Interval:=\"m\"))").asText()=="2024-02-29");
    assert(eval("CStr(DateAdd(\"yyyy\",1,\"2024-02-29\"))").asText()=="2025-02-28");
    assert(eval("CStr(DateAdd(\"w\",1,\"2024-03-01\"))").asText()=="2024-03-02");
    assert(eval("CStr(DateAdd(\"d\",1,CDate(-1.25)))").asText()=="06:00:00");
    assert(eval("DateAdd(\"d\",1,Null)").kind==VbaValue::Kind::Null);
    assert(eval("CStr(DateSerial(Day:=0,Year:=2024,Month:=3))").asText()=="2024-02-29");
    assert(eval("CStr(DateSerial(2024,14,1))").asText()=="2025-02-01");
    assert(eval("CStr(TimeSerial(12-6,-15,0))").asText()=="05:45:00");
    assert(eval("CStr(TimeSerial(0,0,-1))").asText()=="1899-12-29 23:59:59");
    assert(eval("CDbl(TimeSerial(24,0,0))").asNumber()==1 && eval("VarType(DateSerial(2024,1,1))=vbDate").truthy());
    assert(eval("VarType(CDate(2.5))=vbDate").truthy() && eval("TypeName(CDate(2.5))").asText()=="Date");
    assert(eval("VarType(Now)=vbDate And VarType(Date())=vbDate And VarType(Time)=vbDate").truthy());
    assert(eval("CStr(CDate(0))").asText()=="00:00:00");
    assert(!eval("IsNumeric(CDate(2.5))").truthy() && eval("Len(CDate(2.5))").asNumber()==8);
    assert(eval("CDbl(CDate(2.5))").asNumber()==2.5);
    assert(eval(R"(CDbl(CDate("1899-12-29 06:00:00")))").asNumber()==-1.25);
    assert(eval(R"(CStr(CDate(-1.25)))").asText()=="1899-12-29 06:00:00");
    assert(eval(R"(Year(CDate("2024-02-29")))").asNumber()==2024);
    assert(eval(R"(CStr(CDate("2024-02-28")+1))").asText()=="2024-02-29");
    assert(eval(R"(CDate("2024-03-01")-CDate("2024-02-28"))").asNumber()==2);
    std::vector<VbaValue> dateArguments{eval("CDate(\"2024-02-28\")")};
    const auto typed=executeVbaProcedureWithArguments("Function TypedDate(ByVal input As Date) As Date\nDim value As Date\nvalue=input+1\nTypedDate=value\nEnd Function","TypedDate",host,dateArguments);
    assert(typed.status==VbaExecutionStatus::Completed && typed.typedReturnValue && typed.typedReturnValue->kind==VbaValue::Kind::Date && typed.typedReturnValue->asText()=="2024-02-29");
    for(const auto& [source,expected]:std::vector<std::pair<std::string,int>>{
        {R"(DatePart("q","2024-05-01"))",2},{R"(DatePart("y","2024-12-31"))",366},
        {R"(DatePart("w","2024-03-03",vbMonday))",7},{R"(DatePart("ww","2024-01-07"))",2},
        {R"(DatePart(Date:="2024-02-29 23:59:58",Interval:="s"))",58},
        {R"(Year("2024-02-29 23:59:58"))",2024},{R"(Month("2024-02-29"))",2},{R"(Day("2024-02-29"))",29},
        {R"(Hour("23:59:58"))",23},{R"(Minute("2024-02-29T23:59:58"))",59},{R"(Second("23:59:58"))",58},
        {R"(Weekday("2024-03-03"))",1},{R"(Weekday("2024-03-03",vbMonday))",7},{R"(Year(Empty))",1899}}) {
        const auto result=eval(source); assert(result.asNumber()==expected && result.numericType==VbaValue::NumericType::Integer);
    }
    for(const auto& [interval,first,last,expected]:std::vector<std::tuple<std::string,std::string,std::string,int>>{
        {"yyyy","2023-12-31","2024-01-01",1},{"q","2024-03-31","2024-04-01",1},{"m","2024-01-31","2024-02-01",1},
        {"d","2024-02-28","2024-03-01",2},{"y","2024-03-01","2024-02-28",-2},
        {"h","2024-02-29 23:59:59","2024-03-01 00:00:00",1},{"n","23:59:59","23:58:59",-1},
        {"s","2024-02-29 23:59:59","2024-03-01 00:00:00",1},{"ww","2024-03-02","2024-03-03",1}}) {
        const auto expression="DateDiff(\""+interval+"\",\""+first+"\",\""+last+"\")"; const auto result=eval(expression);
        assert(result.asNumber()==expected && result.numericType==VbaValue::NumericType::Long);
    }
    assert(eval(R"(DateDiff("w","2024-01-01 23:00:00","2024-01-08 01:00:00"))").asNumber()==1);
    assert(eval(R"(DateDiff("w","2024-01-09","2024-01-01",vbMonday))").asNumber()==-1);
    assert(eval(R"(DateDiff("w","2024-01-01","2024-01-07"))").asNumber()==0);
    assert(eval(R"(DateDiff("ww","2024-03-02","2024-03-03",vbMonday))").asNumber()==0);
    assert(eval(R"(DateDiff("ww","2024-03-02","2024-03-03",,vbFirstJan1))").asNumber()==1);
    assert(eval("DatePart(\"m\",Null)").kind==VbaValue::Kind::Null);
    assert(eval("Year(Null)").kind==VbaValue::Kind::Null && eval("Weekday(Null)").kind==VbaValue::Kind::Null);
    for(const auto& [source,wanted]:std::vector<std::pair<std::string,int>>{
        {R"(Year("2023-02-29"))",13},{R"(Hour("24:00:00"))",13},{"Year(1)",3251},{"Year()",450},
        {"DateValue(Null)",94},{"TimeValue(\"2024-02-30 12:00\")",13},{"DateValue(1)",3251},
        {"DateAdd(\"m\",1,\"9999-12-31\")",5},{"DateAdd(\"bad\",1,0)",5},{"DateAdd(\"d\",Null,0)",94},
        {"DateSerial(24,1,1)",3251},{"DateSerial(100,1,0)",5},{"TimeSerial(32768,0,0)",6},
        {"DateSerial(2024,Null,1)",94},{"TimeSerial(1,2)",450},{"CDate(Null)",94},{"CDate(2958466)",6},{"CDate(-657435)",6},{"CDate(-1.25)+1",3251},
        {R"(DatePart("ww","2024-01-01",,vbFirstFourDays))",3251},{R"(DatePart("bad","2024-01-01"))",5},
        {R"(DateDiff("invalid","2024-01-01","2024-01-02"))",5},{R"(DateDiff("d",Null,"2024-01-02"))",94},
        {R"(DateDiff("s","0100-01-01","9999-12-31"))",6},{R"(Weekday("2024-01-01",8))",5},
        {R"(Weekday("2024-01-01",vbUseSystem))",3251}}) {
        int error=0; try { (void)eval(source); } catch(const VbaRuntimeError& failure) { error=failure.number; } assert(error==wanted);
    }
    const auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::string("2024-02-29");
    assert(evaluateVbaExpression("Day(probe)",host,{{"probe",VbaValue::reference(probe)}}).asNumber()==29 && probe->reads==1);
    const auto executed=executeVbaProcedure(R"VBA(Sub CalendarValues()
Me!YearValue=CStr(Year("2024-02-29"))
Me!DayDifference=CStr(DateDiff("d","2024-02-28","2024-03-01"))
Me!CalendarTypes=CStr(VarType(Year("2024-02-29"))) & ":" & CStr(VarType(DateDiff("d","2024-02-28","2024-03-01")))
End Sub
)VBA","CalendarValues",host);
    assert(executed.status==VbaExecutionStatus::Completed && host.fields["YearValue"]=="2024" && host.fields["DayDifference"]=="2" && host.fields["CalendarTypes"]=="2:3");
}

void testNativeDecimalVariants() {
    using namespace linaccess;FakeForm host;const auto eval=[&](const std::string& source) {return evaluateVbaExpression(source,host,{});};
    const std::string maximum="79228162514264337593543950335";
    assert(eval("VarType(CDec(1))=vbDecimal").truthy() && eval("TypeName(CDec(1))").asText()=="Decimal");
    assert(eval("VBA.CDec(Expression:=\"1.25\")").asText()=="1.25");
    assert(eval("CStr(CDec(\""+maximum+"\"))").asText()==maximum);
    assert(eval("Str(CDec(\""+maximum+"\"))").asText()==" "+maximum);assert(eval("Str$(-CDec(\""+maximum+"\"))").asText()=="-"+maximum);assert(eval("Str(CDec(\"0.0000000000000000000000000001\"))").asText()==" .0000000000000000000000000001");

    assert(eval("CDec(\""+maximum+"\")-CDec(1)").asText()=="79228162514264337593543950334");
    assert(eval("CDec(\"0.1\")+CDec(\"0.2\")").asText()=="0.3");
    assert(eval("CDec(\"12345678901234567890\")*CDec(\"0.00000001\")=CDec(\"123456789012.3456789\")").truthy());
    assert(eval("CDec(1)/CDec(3)").asText()=="0.3333333333333333333333333333");
    assert(eval("CLng(CDec(\"2.5000000000000000000000000001\"))").asNumber()==3);
    assert(eval("CInt(CDec(\"-2.5\"))").asNumber()==-2 && eval("CByte(CDec(\"-0.5\"))").asNumber()==0);
    assert(eval("Round(CDec(\"1.245\"),2)").asText()=="1.24" && eval("Round(CDec(\"1.255\"),2)").asText()=="1.26");
    assert(eval("Round(CDec(\"0.0000000000000000000000000001\"),28)").asText()=="0.0000000000000000000000000001");
    assert(eval("Int(CDec(\"-1.9\"))").asText()=="-2" && eval("Fix(CDec(\"-1.9\"))").asText()=="-1");
    assert(eval("CCur(CDec(\"1.23445\"))").asText()=="1.2344" && eval("CCur(CDec(\"1.23455\"))").asText()=="1.2346");
    assert(eval("Abs(-CDec(\""+maximum+"\"))").asText()==maximum);
    for(const auto& [source,wanted]:std::vector<std::pair<std::string,int>>{{"CDec(Null)",94},{"CDec(\"invalid\")",13},{"CDec(\"79228162514264337593543950336\")",6},{"CDec(\""+maximum+"\")+CDec(1)",6},{"CDec(1)/CDec(0)",11},{"Round(CDec(1),29)",3251},{"CLng(CDec(\"2147483647.5000000001\"))",6}}) {int error=0;try {(void)eval(source);}catch(const VbaRuntimeError& failure) {error=failure.number;}assert(error==wanted);}
    const auto tie=1.0/32;assert(convertCurrency(VbaValue::numeric(tie)).currencyUnits==312);assert(convertCurrency(VbaValue::numeric(std::nextafter(tie,1.0))).currencyUnits==313);assert(convertCurrency(VbaValue::numeric(std::nextafter(tie,0.0))).currencyUnits==312);assert(convertCurrency(VbaValue::numeric(3.0/32)).currencyUnits==938);assert(convertCurrency(VbaValue::numeric(std::nextafter(-tie,-1.0))).currencyUnits==-313);
    const double edge=922337203685477.5;assert(convertCurrency(VbaValue::numeric(edge)).currencyUnits==9223372036854775000LL && convertCurrency(VbaValue::numeric(-edge)).currencyUnits==-9223372036854775000LL);assert(convertCurrency(VbaValue::numeric(std::numeric_limits<double>::denorm_min())).currencyUnits==0);
    for(const auto number:{std::nextafter(edge,std::numeric_limits<double>::infinity()),std::nextafter(-edge,-std::numeric_limits<double>::infinity()),std::numeric_limits<double>::infinity(),std::numeric_limits<double>::quiet_NaN()}) {int error=0;try {(void)convertCurrency(VbaValue::numeric(number));}catch(const VbaRuntimeError& failure) {error=failure.number;}assert(error==6);}
    auto probe=std::make_shared<ScalarProbe>();probe->value=VbaValue::string(maximum);assert(evaluateVbaExpression("CDec(probe)",host,{{"probe",VbaValue::reference(probe)}}).asText()==maximum && probe->reads==1);
    probe->value=decimalLiteral(maximum);probe->reads=0;assert(evaluateVbaExpression("Str(probe)",host,{{"probe",VbaValue::reference(probe)}}).asText()==" "+maximum && probe->reads==1);
    const auto result=executeVbaProcedure(R"VBA(Sub DecimalValues()
Dim value As Variant, entries(0 To 0) As Variant
value=CDec("79228162514264337593543950335")
entries(0)=value
entries(0)=entries(0)-CDec(1)
Me!DecimalState=CStr(entries(0)) & ":" & CStr(VarType(entries(0)))
End Sub
)VBA","DecimalValues",host);assert(result.status==VbaExecutionStatus::Completed && host.fields["DecimalState"]=="79228162514264337593543950334:14");
    NativeReportAggregate sum("sum"),minimum("min"),maximumAggregate("max");for(const auto& source:{"12345678901234567890.1","0.2"}) {const auto value=decimalLiteral(source);sum.add(value);minimum.add(value);maximumAggregate.add(value);}assert(sum.result().asText()=="12345678901234567890.3" && minimum.result().asText()=="0.2" && maximumAggregate.result().asText()=="12345678901234567890.1");
    assert(eval("Null Or 2").kind==VbaValue::Kind::Null && eval("Null And 2").kind==VbaValue::Kind::Null);
    assert(host.eval("-2^2").asNumber()==4 && eval("-2^2").asNumber()==-4);
    assert(host.eval("Nz(Null)").kind==VbaValue::Kind::Empty && evaluateAccessExpression("Nz(Null)",host).kind==VbaValue::Kind::Text);
}

int main() {
    testNativeDecimalVariants();
    {
        FakeForm host;
        const auto eval=[&](const std::string& expression) { return linaccess::evaluateVbaExpression(expression,host,{}); };
        double savings=0;
        for (int month=0;month<12;++month) savings=savings*1.01+100;
        const auto future=eval("FV(0.01,12,-100)");
        assert(future.numericType==linaccess::VbaValue::NumericType::Double && std::abs(future.number-savings)<1e-9);
        assert(std::abs(eval("FV(Type:=1,Pmt:=-100,NPer:=12,Rate:=0.01)").number-savings*1.01)<1e-9);
        assert(eval("FV(0,12,-100,,1)").number==1200);
        assert(std::abs(eval("FV(0.000000000000001,12,-100)").number-1200)<1e-9);
        const auto payment=eval("Pmt(Pv:=1000,Rate:=0.01,NPer:=12,Fv:=100)").number;
        double loan=1000;
        for (int month=0;month<12;++month) {
            const auto period=std::to_string(month+1);
            const auto interest=eval("IPmt(0.01,"+period+",12,1000,100)").number;
            const auto principal=eval("PPmt(0.01,"+period+",12,1000,100)").number;
            assert(std::abs(interest+loan*0.01)<1e-9 && std::abs(interest+principal-payment)<1e-9);
            loan=loan*1.01+payment;
        }
        assert(std::abs(loan+100)<1e-9);
        assert(eval("IPmt(Type:=1,Pv:=100,NPer:=2,Per:=1,Rate:=0.1)").number==0);
        const auto earlyPayment=eval("Pmt(0.1,2,100,,1)").number;
        assert(std::abs(eval("IPmt(0.1,2,2,100,,1)").number+(100+earlyPayment)*0.1)<1e-11);
        assert(eval("PPmt(0,1,12,1200)").number==-100);
        assert(eval("Pmt(0,12,1000,200)").number==-100);
        double discounted=0,discountFactor=1;
        for(int month=0;month<12;++month) { discountFactor*=1.01; discounted+=100/discountFactor; }
        assert(std::abs(eval("PV(Type:=0,Pmt:=-100,NPer:=12,Rate:=0.01)").number-discounted)<1e-9);
        assert(eval("PV(0,12,-100)").number==1200);
        assert(eval("NPer(0,-100,1200)").number==12);
        assert(std::abs(eval("NPer(0.1,-110,100)").number-1)<1e-12);
        assert(std::abs(eval("VBA.NPer(Type:=1,Pv:=100,Pmt:=-100,Rate:=0.1)").number-1)<1e-12);
        assert(std::abs(eval("NPer(0.000000000000001,-100,1200)").number-12)<1e-9);
        assert(std::abs(eval("NPer(0.1,0,100,-200)").number-7.272540897341713)<1e-12);
        assert(std::abs(eval("VBA.Rate(NPer:=1,Pmt:=-105,Pv:=100,Guess:=0.2)").number-0.05)<1e-12);
        assert(std::abs(eval("Rate(12,Pmt(0.01,12,1000,100),1000,100)").number-0.01)<1e-11);
        assert(eval("Rate(12,-100,1200,,,0)").number==0);


        const auto cashflows=linaccess::executeVbaProcedure(R"VBA(Sub DiscountCashflows()
Dim flows(-1 To 1) As Double
flows(-1)=-100
flows(0)=60
flows(1)=60
Me!Npv=CStr(VBA.NPV(Values:=flows,Rate:=0.1))
Me!NpvZero=CStr(NPV(0,flows))
Dim financed(5 To 8) As Double
financed(5)=-100
financed(6)=110
financed(7)=-100
financed(8)=110
Me!Mirr=CStr(MIRR(Reinvest_rate:=0.1,Values:=financed,Finance_rate:=0.1))
Me!Irr=CStr(IRR(Values:=flows,Guess:=0.2))
Me!IrrCycle=CStr(VBA.IRR(financed))
End Sub
)VBA","DiscountCashflows",host);
        assert(cashflows.status==linaccess::VbaExecutionStatus::Completed);
        assert(std::abs(std::stod(host.fields["Npv"].value())-3.756574004507888)<1e-9 && host.fields["NpvZero"]=="20");
        assert(std::abs(std::stod(host.fields["Mirr"].value())-0.1)<1e-12);
        const auto investmentRate=std::stod(host.fields["Irr"].value());
        double investment=-100; for(const auto receipt:{60,60}) investment=investment*(1+investmentRate)+receipt;
        assert(std::abs(investment)<1e-8 && std::abs(std::stod(host.fields["IrrCycle"].value())-0.1)<1e-12);

        assert(eval("SLN(Life:=5,Salvage:=1000,Cost:=10000)").number==1800);
        assert(eval("DDB(1000,100,5,1)").number==400);
        assert(std::abs(eval("DDB(1000,100,5,5)").number-29.6)<1e-10);
        assert(eval("DDB(Factor:=1,Period:=2,Life:=5,Salvage:=0,Cost:=1000)").number==160);

        double depreciation=0;
        for(int year=1;year<=5;++year) depreciation+=eval("SYD(10000,1000,5,"+std::to_string(year)+")").number;
        assert(depreciation==9000);
        assert(eval("SYD(Period:=1,Life:=5,Salvage:=1000,Cost:=10000)").number==3000);

        bool invalid=false;
        try { eval("Pmt(0,0,1000)"); } catch(const linaccess::VbaRuntimeError& e) { invalid=e.number==5; }
        assert(invalid);
        invalid=false;
        try { eval("NPer(0,0,1000)"); } catch(const linaccess::VbaRuntimeError& e) { invalid=e.number==5; }
        assert(invalid);
        invalid=false;
        try { eval("Rate(12,0,0)"); } catch(const linaccess::VbaRuntimeError& e) { invalid=e.number==5; }
        assert(invalid);
    }
    {
        FakeForm host;
        const auto eval=[&](const std::string& source) { return linaccess::evaluateVbaExpression(source,host,{}); };
        assert(eval("Partition(0,0,99,5)").asText()=="  0:  4");
        assert(eval("Partition(Interval:=20,Stop:=1010,Start:=100,Number:=1001)").asText()=="1000:1010");
        assert(eval("Partition(-1,0,99,5)").asText()=="   : -1");
        assert(eval("Partition(100,0,99,5)").asText()=="100:   ");
        assert(eval("Partition(4.5,0,99,5)").asText()=="  0:  4");
        assert(eval("Partition(5.5,0,99,5)").asText()=="  5:  9");
        assert(eval("Partition(Null,0,99,0)").kind==linaccess::VbaValue::Kind::Null);
        bool invalid=false;
        try { eval("Partition(1,0,99,0)"); } catch(const linaccess::VbaRuntimeError& e) { invalid=e.number==5; }
        assert(invalid);
    }
    testNativeCollection();
    testCalendarFunctions();
    testCurrencyRealStorage();
    testCurrencyForLoops();
    testNativeTypedForLoops();
    testIsObjectFunction();
    testStringResultTypes();
    testReplaceFunction();
    testRoundFunction();
    testNativeMathFunctions();
    testCharacterFunctions();
    testDefaultTypeInspection();
    testTypedNot();
    {
        FakeForm host;
        const auto inspect=[&](const std::string& expression) { return linaccess::evaluateVbaExpression(expression,host,{}); };
        assert(inspect("VarType(VarType(CByte(1)))=vbInteger").truthy());
        assert(inspect("TypeName(VarType(Null))").asText()=="Integer");
        assert(inspect("VarType(Null)").number==1 && inspect("VarType(CVErr(11))").number==10);
        const auto bytes=linaccess::VbaValue::byteArray({1,2,3});
        const auto code=linaccess::evaluateVbaExpression("VarType(bytes)",host,{{"bytes",bytes}});
        assert(code.number==8209 && code.numericType==linaccess::VbaValue::NumericType::Integer);
    }
    testByteValues();
    using namespace linaccess;

    {
        class Indexed final:public VbaObject,public std::enable_shared_from_this<Indexed> {
        public:
            std::map<int,VbaValue> rows;
            std::map<std::string,VbaValue> properties;
            int indexReads{},flagReads{},childReads{},getterReads{};
            VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
                if(name=="Child") { assert(args.empty()); ++childReads; return VbaValue::reference(shared_from_this()); }
                if(name=="Index") { assert(args.empty()); ++indexReads; return VbaValue::numeric(1); }
                if(name=="Flag") { assert(args.empty()); ++flagReads; return VbaValue::logical(true); }
                if(name=="Selected") { ++getterReads; throw VbaRuntimeError(5,"Getter must not implement assignment"); }
                throw VbaRuntimeError(438,"Unknown member");
            }
            void assign(std::string_view name,const VbaValue& value) override { properties[std::string(name)]=value; }
            void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& args,const VbaValue& value) override {
                if(name!="Selected") throw VbaRuntimeError(438,"Unknown indexed property");
                if(args.size()!=1) throw VbaRuntimeError(450,"Expected one index");
                rows[static_cast<int>(args[0].asNumber())]=value;
            }
        };
        FakeForm host; auto object=std::make_shared<Indexed>();
        std::vector<VbaValue> arguments{VbaValue::reference(object)};
        auto result=executeVbaProcedureWithArguments(R"VBA(Sub Write(target As Object)
target.Child.[Selected](CLng(target.Index())+1)=target.Flag()
With target.Child
.Selected(Len("a=b"))=Not target.Flag()
End With
target.[Label.with(=).dot]="x=y"
target.Plain="ordinary"
End Sub
)VBA","Write",host,arguments);
        assert(result.status==VbaExecutionStatus::Completed && object->rows.size()==2 && object->rows[2].truthy() && !object->rows[3].truthy());
        assert(object->indexReads==1 && object->flagReads==2 && object->childReads==2 && object->getterReads==0);
        assert(object->properties["Label.with(=).dot"].asText()=="x=y" && object->properties["Plain"].asText()=="ordinary");
        for(const auto& target:std::vector<std::string>{"target.Selected(0)+1","target.Selected(0)(1)","target.Plain + target.Plain"}) {
            const auto invalid=executeVbaProcedureWithArguments("Sub Invalid(target As Object)\n"+target+"=True\nEnd Sub","Invalid",host,arguments);
            assert((target=="target.Plain + target.Plain" ? invalid.status!=VbaExecutionStatus::Completed : invalid.status==VbaExecutionStatus::Unsupported)
                && object->rows.size()==2 && object->properties.size()==2 && object->getterReads==0);
        }
        result=executeVbaProcedureWithArguments("Sub Bad(target As Object)\nOn Error Resume Next\ntarget.Selected()=True\nMe!Code=CStr(Err.Number)\nEnd Sub","Bad",host,arguments);
        assert(result.status==VbaExecutionStatus::Completed && host.fields["Code"]=="450");
        arguments={VbaValue::reference(std::make_shared<LoopContext>(1))};
        result=executeVbaProcedureWithArguments("Sub Bad(target As Object)\nOn Error Resume Next\ntarget.Value(0)=True\nMe!Code=CStr(Err.Number)\nEnd Sub","Bad",host,arguments);
        assert(result.status==VbaExecutionStatus::Completed && host.fields["Code"]=="438");
    }

    {
        FakeForm host;
        assert(evaluateVbaExpression("Eval(\"1+2*3\")",host,{}).asNumber()==7);
        assert(evaluateVbaExpression("Eval(5)",host,{}).asNumber()==5);
        assert(host.eval("'O''Brien'").asText()=="O'Brien");
        assert(host.eval("'alpha'='ALPHA'").truthy());
        assert(host.eval("Null").kind==VbaValue::Kind::Null);
        assert(host.eval("Null Is Null").truthy() && !host.eval("0 Is Null").truthy());
        assert(!host.eval("Null Is Not Null").truthy() && host.eval("'text' Is Not Null").truthy());
        assert(host.eval("Null Is Null And 1=1").truthy());
        assert(!validVbaExpression("'not a VBA literal'"));
        for (const auto& [expression,number]:std::vector<std::pair<std::string,int>>{
            {"Eval()",450},{"Eval(1,2)",450},{"Eval(Null)",94}}) {
            try { (void)evaluateVbaExpression(expression,host,{}); assert(false); }
            catch (const VbaRuntimeError& failure) { assert(failure.number==number); }
        }
        host.evalModule=R"VBA(Public Function PlusOne(value As Double) As Double
PlusOne=value+1
End Function
Public Function NeedCancel(Cancel As Integer) As Long
NeedCancel=Cancel
End Function
Public Function EvalFailure() As Long
Err.Raise 7002,"eval.source","eval.detail","eval.chm",42
End Function
Public Function RecurseEval() As Variant
RecurseEval=Eval("RecurseEval()")
End Function
Private Function Hidden() As Long
DoCmd.RunSQL "INSERT INTO markers VALUES(98)"
Hidden=1
End Function
Public Sub NotAFunction()
DoCmd.RunSQL "INSERT INTO markers VALUES(99)"
End Sub
)VBA";
        assert(host.eval("PlusOne(7)").asNumber()==8);
        try {(void)host.eval("EvalFailure()");assert(false);}catch(const VbaRuntimeError& failure) {assert(failure.number==7002 && failure.source=="eval.source" && failure.helpFile=="eval.chm" && failure.helpContext==42 && std::string(failure.what())=="eval.detail");}
        for (const auto& [expression,number]:std::vector<std::pair<std::string,int>>{
            {"PlusOne()",450},{"NeedCancel()",450},{"Hidden()",2425},{"NotAFunction()",2425},{"Unknown()",2425},{"RecurseEval()",28}}) {
            try { (void)host.eval(expression); assert(false); }
            catch (const VbaRuntimeError& failure) { assert(failure.number==number); }
        }
        assert(host.sqlStatements.empty() && host.eval("1+1").asNumber()==2);
        const auto statements=executeVbaProcedure(R"VBA(Sub Test()
Eval("PlusOne(1)")
Call Eval("PlusOne(2)")
Eval "PlusOne(3)"
End Sub
)VBA","Test",host);
        assert(statements.status==VbaExecutionStatus::Completed);
    }

    {
        FakeForm host;
        for (const auto& [expression,result]:std::vector<std::pair<std::string,std::string>>{
            {"Str(12)"," 12"},{"Str$(-12)","-12"},{"Str(0)"," 0"},{"Str(-0)"," 0"},
            {"Str(0.5)"," .5"},{"Str(-0.5)","-.5"},{"Str(Empty)"," 0"},{"Str(True)","-1"},
            {"Str(CCur(1.2345))"," 1.2345"}})
            assert(evaluateVbaExpression(expression,host,{}).asText()==result);
        for (const auto& [expression,number]:std::vector<std::pair<std::string,int>>{
            {"Str()",450},{"Str(1,2)",450},{"Str(Null)",94},{"Str(\"bad\")",13}}) {
            try { (void)evaluateVbaExpression(expression,host,{}); assert(false); }
            catch (const VbaRuntimeError& failure) { assert(failure.number==number); }
        }
        const auto bytes=VbaValue::byteArray({1,2,3});
        assert(evaluateVbaExpression("IsArray(token)",host,{{"token",bytes}}).truthy());
        assert(evaluateVbaExpression("VarType(token)=vbArray+vbByte",host,{{"token",bytes}}).truthy());
        assert(evaluateVbaExpression("TypeName(token)",host,{{"token",bytes}}).asText()=="Byte()");
        try { (void)bytes.asText(); assert(false); } catch (const VbaRuntimeError& failure) { assert(failure.number==13); }
        try { (void)bytes.asNumber(); assert(false); } catch (const VbaRuntimeError& failure) { assert(failure.number==13); }
    }

    {
        FakeForm host;
        for (const auto& literal:std::vector<std::string>{"(",")","-","+","not","and","or",".","!",",",";"}) {
            const auto source="\""+literal+"\"";
            assert(validVbaExpression(source));
            const auto value=evaluateVbaExpression(source,host,{});
            assert(value.kind==VbaValue::Kind::Text && value.asText()==literal);
        }
        assert(evaluateVbaExpression("\"(\" & 2465 & \")\"",host,{}).asText()=="(2465)");
        assert(!validVbaExpression("1 \"and\" 2"));
        assert(validVbaExpression("MsgBox(\")\")"));
        assert(evaluateVbaExpression("vbKeyF",host,{}).asNumber()==70);
        assert(evaluateVbaExpression("vbKeyF1",host,{}).asNumber()==112);
        assert(evaluateVbaExpression("vbKeyF24",host,{}).asNumber()==135);
        assert(evaluateVbaExpression("vbKey9",host,{}).asNumber()==57);
        assert(evaluateVbaExpression("vbKeyReturn",host,{}).asNumber()==13);
        assert(evaluateVbaExpression("acShiftMask Or acCtrlMask Or acAltMask",host,{}).asNumber()==7);
    }

    {
        FakeForm host;
        const std::string module=R"VBA(Function Inline(x As Long,y As Long) As Long
Dim value As Long
If x Then value=1: value=value+1 Else value=7
If x Then If y Then value=3 Else value=4
If value=7 Then GoTo done Else value=value+10
done:
Inline=value
End Function
Function Quoted() As String
If "Then Else"="Then Else" Then Quoted="Then:Else" Else Quoted="bad"
End Function
Function Recover() As Long
On Error GoTo failed
Err.Raise 2465
Recover=23
Exit Function
failed:
If Err=2465 Then Resume Next
Recover=99
End Function
)VBA";
        for (const auto& [x,y,expected]:std::vector<std::tuple<int,int,int>>{{0,0,7},{0,1,7},{1,0,14},{1,1,13}}) {
            std::vector<VbaValue> args{VbaValue::numeric(x),VbaValue::numeric(y)};
            const auto result=executeVbaProcedureWithArguments(module,"Inline",host,args);
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==expected);
        }
        const auto quoted=executeVbaProcedure(module,"Quoted",host);
        assert(quoted.status==VbaExecutionStatus::Completed && quoted.typedReturnValue->asText()=="Then:Else");
        const auto recovered=executeVbaProcedure(module,"Recover",host);
        assert(recovered.status==VbaExecutionStatus::Completed && recovered.typedReturnValue->asNumber()==23);
    }

    {
        FakeForm host;
        const std::string module=R"VBA(Option Compare Text
Private reads As Variant
Function Probe(x As Variant) As Variant
reads=reads+1
Probe=x
End Function
Function Choose(x As Variant) As Long
reads=0
Select Case Probe(x)
Case 1, 3, 5
Choose=10
Case "A"
Choose=30
Case "A,B", "To"
Choose=35
Case 6 To 9, Is >= 20
Choose=20
Case Else
Choose=40
End Select
Choose=Choose+reads*100
End Function
Function Nested() As Long
Dim i As Long
For i=1 To 4
Select Case i
Case 1, 2
If i=1 Then
Select Case 7
Case Is < 8
Nested=Nested+1
Case Else
Nested=999
End Select
Else
Nested=Nested+2
End If
Case 3
Exit For
Case Else
Nested=999
End Select
Next
End Function
Function Lazy() As Long
reads=0
Select Case Probe(3)
Case Probe(3), Probe(1/0)
Lazy=1
Case Probe(1/0)
Lazy=99
End Select
Select Case 1
Case 3 To Probe(1/0)
Lazy=99
Case Else
Lazy=Lazy+1
End Select
If False Then
Select Case Probe(1/0)
Case Probe(1/0)
Lazy=99
End Select
End If
Lazy=Lazy+reads*100
End Function
Function RecoverCase() As Long
Dim divisor As Long
divisor=0
On Error GoTo failed
Select Case 2
Case 4/divisor
RecoverCase=7
Case Else
RecoverCase=99
End Select
Exit Function
failed:
divisor=2
Resume
End Function
Function RecoverSelector() As Long
Dim divisor As Long
On Error GoTo failed
Select Case 2/divisor
Case 2
RecoverSelector=23
Case Else
RecoverSelector=99
End Select
Exit Function
failed:
divisor=1
Resume
End Function
Function SkipErrors() As Long
On Error Resume Next
Select Case 1/0
Case Else
SkipErrors=999
End Select
Select Case 2
Case 1/0
SkipErrors=999
Case 2
SkipErrors=3
End Select
SkipErrors=SkipErrors+Err.Number
End Function
Function EmptyBlock() As Long
Select Case 1
End Select
EmptyBlock=8
End Function
)VBA";
        for (const auto& [input,expected]:std::vector<std::pair<VbaValue,int>>{
                {VbaValue::numeric(1),110},{VbaValue::numeric(7),120},{VbaValue::numeric(20),120},
                {VbaValue::numeric(10),140},{VbaValue::string("a"),130},{VbaValue{},140},
                {VbaValue::empty(),140},{VbaValue::string("1"),120},{VbaValue::string("a,b"),135}}) {
            std::vector<VbaValue> args{input};
            auto result=executeVbaProcedureWithArguments(module,"Choose",host,args);
            if (result.status!=VbaExecutionStatus::Completed || result.typedReturnValue->asNumber()!=expected)
                std::cerr<<"Select input kind="<<static_cast<int>(input.kind)<<", status="<<static_cast<int>(result.status)<<", line="<<result.line<<", reason="<<result.explanation<<", expected="<<expected<<", actual="<<(result.typedReturnValue ? result.typedReturnValue->asNumber() : -1)<<"\n";
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==expected);
        }
        for (const auto& [procedure,expected]:std::vector<std::pair<std::string,int>>{
                {"Nested",3},{"Lazy",202},{"RecoverCase",7},{"RecoverSelector",23},{"SkipErrors",14},{"EmptyBlock",8}}) {
            auto result=executeVbaProcedure(module,procedure,host);
            if (result.status!=VbaExecutionStatus::Completed || result.typedReturnValue->asNumber()!=expected)
                std::cerr<<"Select Case regression: "<<procedure<<", status="<<static_cast<int>(result.status)<<", error="<<result.errorNumber<<"\n";
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==expected);
        }
        for (const auto& body:std::vector<std::string>{
                "Case 1", "End Select", "Select Case 1\nMe!Name=1\nCase 1\nEnd Select", "Select Case 1\nIf True Then Me!Name=1\nCase 1\nEnd Select",
                "Select Case 1\nCase Else\nCase 1\nEnd Select", "Select Case 1\nCase 1",
                "Select Case 1\nCase 1,\nEnd Select", "Select Case 1\nCase Is Like 1\nEnd Select",
                "Select Case 1\nCase 1 To\nEnd Select", "Select Case 1\nIf True Then\nCase 1\nEnd If\nEnd Select"}) {
            host.fields.clear();
            auto result=executeVbaProcedure("Sub Invalid()\nMe!Name=42\n"+body+"\nEnd Sub","Invalid",host);
            assert(result.status!=VbaExecutionStatus::Completed && host.fields.empty());
        }
        auto probe=std::make_shared<ScalarProbe>(); probe->value=VbaValue::numeric(3);
        std::vector<VbaValue> objectArgs{VbaValue::reference(probe)};
        auto once=executeVbaProcedureWithArguments(module,"Choose",host,objectArgs);
        assert(once.status==VbaExecutionStatus::Completed && once.typedReturnValue->asNumber()==110 && probe->reads==1);
        host.bracketHost=true; host.bracketFields["a,b"]=VbaValue::numeric(2);
        auto bracket=executeVbaProcedure("Function BracketCase() As Long\nSelect  Case  2\nCase 1, [a,b]\nBracketCase=1\nEnd  Select\nEnd Function","BracketCase",host);
        assert(bracket.status==VbaExecutionStatus::Completed && bracket.typedReturnValue->asNumber()==1);
        for (const auto& relation:std::vector<std::string>{"= 2","<> 3","< 3","<= 2","> 1",">= 2","Is = 2","Is <> 3","Is < 3","Is <= 2","Is > 1","Is >= 2","Is>=2","Is<3","2 To 2"}) {
            auto result=executeVbaProcedure("Function Relation() As Long\nSelect Case 2\nCase "+relation+"\nRelation=1\nEnd Select\nEnd Function","Relation",host);
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==1);
        }
        auto binary=executeVbaProcedure("Function BinaryCase() As Long\nSelect Case \"A\"\nCase \"a\"\nBinaryCase=1\nCase Else\nBinaryCase=2\nEnd Select\nEnd Function","BinaryCase",host);
        assert(binary.status==VbaExecutionStatus::Completed && binary.typedReturnValue->asNumber()==2);
    }

    {
        FakeForm host;
        for (const auto& [opening,closing,expected]:std::vector<std::tuple<std::string,std::string,int>>{
                {"Do While count<3","Loop",3},{"Do Until count=3","Loop",3},
                {"Do","Loop While count<3",3},{"Do","Loop Until count=3",3},
                {"While count<3","Wend",3},{"Do While False","Loop",0},
                {"Do Until True","Loop",0},{"Do","Loop While False",1},
                {"Do","Loop Until True",1},{"Do While Null","Loop",0},
                {"Do","Loop While Null",1}}) {
            auto result=executeVbaProcedure("Function LoopCount() As Long\nDim count As Long\n"+opening+"\ncount=count+1\n"+closing+"\nLoopCount=count\nEnd Function","LoopCount",host);
            if (result.status!=VbaExecutionStatus::Completed) std::cerr<<"Loop regression status="<<static_cast<int>(result.status)<<", line="<<result.line<<"\n";
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==expected);
        }
        const std::string module=R"VBA(Function Nested(outer As Object, inner As Object, items As Object) As Long
Dim i As Long
Dim item As Variant
For i=1 To 2
With outer
Do
For Each item In items
With inner
Select Case item
Case 1
If True Then Exit Do
Case Else
Err.Raise 5
End Select
End With
Next
Err.Raise 5
Loop
Nested=Nested+.Value
End With
Next
End Function
Function Innermost() As Long
Dim count As Long
Do Until count=3
Do
Innermost=Innermost+1
Exit Do
Loop
count=count+1
Loop
End Function
Function ExitForInsideDo() As Long
Dim i As Long
Do
For i=1 To 3
Select Case i
Case 1
Exit For
End Select
Next
ExitForInsideDo=ExitForInsideDo+1
If ExitForInsideDo=2 Then Exit Do
Loop
End Function
Function NullUntil() As Long
Do Until Null
NullUntil=NullUntil+1
If NullUntil=3 Then Exit Do
Loop
Do
NullUntil=NullUntil+1
If NullUntil=5 Then Exit Do
Loop Until Null
End Function
Function ResumeBody() As Long
Dim count As Long
On Error GoTo failed
Do Until count=2
count=count+1
Err.Raise 11
ResumeBody=ResumeBody+1
Loop
Exit Function
failed:
Resume Next
End Function
Function ResumeCondition() As Long
Dim divisor As Long
On Error GoTo failed
Do
ResumeCondition=ResumeCondition+1
Loop Until 2/divisor=2
Exit Function
failed:
divisor=1
Resume
End Function
Function WhileInsideDo() As Long
Do
While True
Select Case 1
Case 1
Exit Do
End Select
Wend
Loop
WhileInsideDo=8
End Function
)VBA";
        auto items=std::make_shared<TestSequence>(); items->values={VbaValue::numeric(1),VbaValue::numeric(2)};
        std::vector<VbaValue> arguments{VbaValue::reference(std::make_shared<LoopContext>(10)),VbaValue::reference(std::make_shared<LoopContext>(20)),VbaValue::reference(items)};
        auto nested=executeVbaProcedureWithArguments(module,"Nested",host,arguments);
        assert(nested.status==VbaExecutionStatus::Completed && nested.typedReturnValue->asNumber()==20);
        for (const auto& [procedure,expected]:std::vector<std::pair<std::string,int>>{
                {"Innermost",3},{"ExitForInsideDo",2},{"NullUntil",5},{"ResumeBody",2},{"ResumeCondition",1},{"WhileInsideDo",8}}) {
            auto result=executeVbaProcedure(module,procedure,host);
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==expected);
        }
        for (const auto& body:std::vector<std::string>{"Exit Do","If True Then Exit Do","Loop","Wend",
                "While True\nExit Do\nWend","Do While True\nWend","While False\nLoop",
                "Do Until False\nLoop Until True","Do\nIf True Then\nLoop\nEnd If","Do","Do Until (\nLoop"}) {
            host.fields.clear();
            auto result=executeVbaProcedure("Sub Invalid()\nMe!Name=42\n"+body+"\nEnd Sub","Invalid",host);
            assert(result.status!=VbaExecutionStatus::Completed && host.fields.empty());
        }
        auto budget=executeVbaProcedure("Sub Forever()\nDo\nLoop\nEnd Sub","Forever",host);
        assert(budget.status==VbaExecutionStatus::Error && budget.explanation.find("10.000")!=std::string::npos);
    }

    {
        FakeForm host;
        const std::string module="Function CurrentProject() As String\nCurrentProject=\"user-function\"\nEnd Function\nFunction CurrentData() As String\nCurrentData=\"user-data\"\nEnd Function\nFunction ReadProject() As String\nReadProject=CurrentProject() & \"|\" & CurrentData()\nEnd Function\n";
        const auto result=executeVbaProcedure(module,"ReadProject",host);
        assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="user-function|user-data");
    }

    {
        FakeForm host;
        const std::string module=R"VBA(Private reads As Variant
Function Probe(value As Variant) As Variant
reads=reads+1
Probe=value
End Function
Function Choose(a As Variant,b As Variant,c As Variant) As Long
reads=0
If Probe(a) Then
Choose=10
ElseIf Probe(b) Then
Choose=20
ElseIf Probe(c) Then
Choose=30
Else
Choose=40
End If
Choose=Choose+reads*100
End Function
Function Lazy() As Long
If True Then
Lazy=1
ElseIf Probe(1/0) Then
Lazy=99
End If
If False Then
If Probe(1/0) Then
Lazy=99
ElseIf Probe(1/0) Then
Lazy=99
Else
Lazy=99
End If
ElseIf True Then
Lazy=Lazy+1
ElseIf Probe(1/0) Then
Lazy=99
End If
End Function
Function Nested() As Long
Dim i As Long
For i=1 To 4
Do
If i=1 Then
Nested=Nested+1
ElseIf i=2 Then
Select Case i
Case 2
If True Then Exit Do
End Select
ElseIf i=3 Then
Exit For
Else
Nested=999
End If
Nested=Nested+10
Exit Do
Loop
Next
End Function
Function RecoverClause() As Long
Dim divisor As Long
On Error GoTo failed
If False Then
RecoverClause=99
ElseIf 2/divisor=2 Then
RecoverClause=7
Else
RecoverClause=99
End If
Exit Function
failed:
divisor=1
Resume
End Function
Function RecoverBody() As Long
On Error GoTo failed
If False Then
RecoverBody=99
ElseIf True Then
Err.Raise 11
RecoverBody=7
ElseIf 1/0 Then
RecoverBody=99
Else
RecoverBody=99
End If
Exit Function
failed:
Resume Next
End Function
Function NoMatch() As Long
If False Then
NoMatch=99
ElseIf Null Then
NoMatch=99
ElseIf False Then
NoMatch=99
End If
NoMatch=NoMatch+1
End Function
)VBA";
        for (const auto& [a,b,c,expected]:std::vector<std::tuple<VbaValue,VbaValue,VbaValue,int>>{
                {VbaValue::logical(true),VbaValue::logical(true),VbaValue::logical(true),110},
                {VbaValue::logical(false),VbaValue::logical(true),VbaValue::logical(true),220},
                {VbaValue::logical(false),VbaValue::logical(false),VbaValue::logical(true),330},
                {VbaValue::logical(false),VbaValue::logical(false),VbaValue::logical(false),340},
                {VbaValue{},VbaValue::logical(true),VbaValue{},220}}) {
            std::vector<VbaValue> args{a,b,c}; auto result=executeVbaProcedureWithArguments(module,"Choose",host,args);
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==expected);
        }
        for (const auto& [procedure,expected]:std::vector<std::pair<std::string,int>>{
                {"Lazy",2},{"Nested",11},{"RecoverClause",7},{"RecoverBody",7},{"NoMatch",1}}) {
            auto result=executeVbaProcedure(module,procedure,host);
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==expected);
        }
        for(const auto& body:std::vector<std::string>{"ElseIf True Then\nEnd If",
                "If True Then\nElse\nElseIf True Then\nEnd If","If True Then\nElseIf ( Then\nEnd If",
                "Select Case 1\nCase 1\nElseIf True Then\nEnd Select"}) {
            host.fields.clear(); auto result=executeVbaProcedure("Sub Invalid()\nMe!Name=42\n"+body+"\nEnd Sub","Invalid",host);
            assert(result.status!=VbaExecutionStatus::Completed && host.fields.empty());
        }
    }

    {
        const std::string module=R"VBA(Private counter As Variant
Private flag As Boolean
Public text As String
Function NextValue() As Variant
counter=counter+1
Call Touch()
NextValue=counter
End Function
Sub Touch()
counter=counter+1
End Sub
Function Recursive(level As Long) As Variant
If level>0 Then
counter=counter+1
Recursive=Recursive(level-1)+counter
Else
Recursive=counter
End If
End Function
Function CurrentValue() As Variant
CurrentValue=counter
End Function
Function Shadow() As Variant
If False Then
Dim counter As Variant
End If
counter=9
Shadow=counter
End Function
Function ForwardLocal() As Variant
ForwardLocal=counter
Dim counter As Variant
End Function
Function TypedFailure() As String
flag=False
On Error Resume Next
Call BreakFlag(flag)
TypedFailure=CStr(Err.Number) & ":" & CStr(VarType(flag))
End Function
Sub BreakFlag(ByRef argument As Variant)
argument=Null
End Sub
Function TextValue() As Variant
text=True
TextValue=text
End Function
)VBA";
        FakeForm first,second;
        const auto call=[&](FakeForm& host,const char* procedure) {
            const auto result=executeVbaProcedure(module,procedure,host);
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue);
            return *result.typedReturnValue;
        };
        assert(call(first,"NextValue").asNumber()==2 && call(first,"NextValue").asNumber()==4);
        const auto event=executeVbaEventExpression(module,"=NextValue()",first);
        assert(event.status==VbaExecutionStatus::Completed && event.typedReturnValue->asNumber()==6);
        assert(executeVbaEventExpression(module,"=NextValue()+)",first).status==VbaExecutionStatus::Unsupported);
        assert(call(first,"CurrentValue").asNumber()==6);
        assert(call(second,"CurrentValue").kind==VbaValue::Kind::Empty);
        std::vector<VbaValue> args{VbaValue::numeric(2)};
        const auto recursive=executeVbaProcedureWithArguments(module,"Recursive",second,args);
        assert(recursive.status==VbaExecutionStatus::Completed && recursive.typedReturnValue->asNumber()==6);
        assert(call(first,"Shadow").asNumber()==9 && call(first,"CurrentValue").asNumber()==6);
        assert(call(first,"ForwardLocal").kind==VbaValue::Kind::Empty);
        assert(call(first,"TypedFailure").asText()=="94:11");
        assert(call(first,"TextValue").kind==VbaValue::Kind::Text && call(first,"TextValue").asText()=="True");
        const auto declarations=createVbaModuleState("Public first,second As String\nPrivate flag As Boolean\n");
        assert(declarations->values.at("first").kind==VbaValue::Kind::Empty && declarations->values.at("second").asText().empty());
        assert(declarations->publicVariables.size()==2 && !declarations->publicVariables.contains("flag"));
    }

    {
        const auto declarations=createVbaModuleState("Public untyped,small As Integer,wide As Long\nPrivate hidden As Long\n");
        assert(declarations->values.at("untyped").kind==VbaValue::Kind::Empty && declarations->values.at("small").asNumber()==0 && declarations->values.at("wide").asNumber()==0);
        assert(declarations->publicVariables.size()==3 && !declarations->publicVariables.contains("hidden"));
        const std::string module=R"VBA(Private small As Integer
Public wide As Long
Function StepWide() As Variant
wide=wide+40000
StepWide=wide
End Function
Sub RoundValues()
small=2.5
wide=-3.5
Me!Rounded=CStr(small)&":"&CStr(wide)
End Sub
Function OverflowSmall() As Variant
On Error Resume Next
small=32767
small=small+1
Me!Code=CStr(Err.Number)
OverflowSmall=small
End Function
Function OverflowWide() As Variant
On Error Resume Next
wide=2147483647
wide=wide+1
Me!Code=CStr(Err.Number)
OverflowWide=wide
End Function
Function Shadow() As Variant
Dim small As Long
small=90000
Shadow=small
End Function
)VBA";
        FakeForm first,second;
        assert(first.moduleState(module)->values.at("small").numericType==VbaValue::NumericType::Integer);
        assert(first.moduleState(module)->values.at("wide").numericType==VbaValue::NumericType::Long);
        assert(executeVbaProcedure(module,"RoundValues",first).status==VbaExecutionStatus::Completed && first.fields["Rounded"]=="2:-4");
        auto result=executeVbaProcedure(module,"StepWide",first); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==39996);
        result=executeVbaProcedure(module,"StepWide",second); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==40000);
        result=executeVbaProcedure(module,"OverflowSmall",first); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==32767 && first.fields["Code"]=="6");
        result=executeVbaProcedure(module,"OverflowWide",first); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==2147483647 && first.fields["Code"]=="6");
        result=executeVbaProcedure(module,"Shadow",first); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==90000 && first.moduleState(module)->values.at("small").asNumber()==32767);
        assert(coerceVbaModuleValue(VbaValue::numeric(-32768.5),"integer").asNumber()==-32768);
        assert(coerceVbaModuleValue(VbaValue::numeric(-2147483648.5),"long").asNumber()==-2147483648.0);
        assert(coerceVbaModuleValue(VbaValue::logical(true),"integer").asNumber()==-1 && coerceVbaModuleValue(VbaValue::empty(),"long").asNumber()==0);
        assert(coerceVbaModuleValue(VbaValue::currency(25000),"integer").asNumber()==2);
        for(const auto& [type,value,expected]:std::vector<std::tuple<std::string,VbaValue,int>>{
            {"integer",VbaValue::numeric(32767.5),6},{"long",VbaValue::numeric(2147483647.5),6},
            {"long",VbaValue::numeric(std::numeric_limits<double>::infinity()),6},{"integer",VbaValue{},94},
            {"long",VbaValue::string("bad"),13},{"integer",VbaValue::errorValue(11),13}}) {
            int actual=0; try { coerceVbaModuleValue(value,type); } catch(const VbaRuntimeError& failure) { actual=failure.number; } assert(actual==expected);
        }
        const auto dynamicState=createVbaModuleState("Public rows() As Long\n");
        assert(dynamicState->types.at("rows")=="long()" && isVbaArray(dynamicState->values.at("rows")) && vbaArrayRank(dynamicState->values.at("rows"))==0);
        const auto fixedState=createVbaModuleState("Public rows(10) As Long\n");
        assert(fixedState->fixedArrays.contains("rows") && vbaArrayBounds(fixedState->values.at("rows"))==std::vector<VbaArrayDimension>({{0,10}}));
        int actual=0; try { createVbaModuleState("Public rows(SIZE) As Long\n"); } catch(const VbaRuntimeError& failure) { actual=failure.number; } assert(actual==3251);
    }

    {
        const std::string module=R"VBA(Sub RoundLocals()
Dim small As Integer
Dim wide As Long
small=2.5
wide=-3.5
Me!Rounded=CStr(small)&":"&CStr(wide)
On Error Resume Next
small=32768
Me!SmallCode=CStr(Err.Number)
wide=2147483648
Me!WideCode=CStr(Err.Number)
Me!Preserved=CStr(small)&":"&CStr(wide)
End Sub
Function RoundedResult() As Integer
RoundedResult=3.5
End Function
Function InvalidResult() As Integer
InvalidResult=32768
End Function
Sub Consume(ByVal value As Integer)
Me!Seen=CStr(value)
End Sub
Sub ChangeLong(ByRef value As Long)
Me!Body="entered"
value=value+1
End Sub
Sub Mismatch()
Dim narrow As Integer
narrow=3
On Error Resume Next
Call ChangeLong(narrow)
Me!Code=CStr(Err.Number)
Me!Seen=CStr(narrow)
End Sub
Sub Temporary()
Dim narrow As Integer
narrow=3
Call ChangeLong((narrow))
Me!Seen=CStr(narrow)
End Sub
Function ReferenceOverflow() As Variant
Dim wide As Long
wide=2147483647
On Error Resume Next
Call ChangeLong(wide)
Me!Code=CStr(Err.Number)
ReferenceOverflow=wide
End Function
Function CountTyped() As Long
Dim i As Integer
For i=1 To 3
CountTyped=CountTyped+i
Next i
Me!Counter=CStr(i)
End Function
Sub OverflowCounter()
Dim i As Integer
For i=32767 To 32767
Me!Body="entered"
Next i
End Sub
Sub FractionalCounter()
Dim i As Long
For i=1 To 2.5
Me!Body="entered"
Next i
End Sub
Sub InvalidEnumerator(items As Object)
Dim item As Integer
Me!Body="entered"
For Each item In items
Me!Seen=CStr(item)
Next item
End Sub
Sub InvalidSet(value As Object)
Dim number As Long
Me!Body="entered"
Set number=value
End Sub
)VBA";
        FakeForm host;
        assert(executeVbaProcedure(module,"RoundLocals",host).status==VbaExecutionStatus::Completed && host.fields["Rounded"]=="2:-4" && host.fields["Preserved"]=="2:-4");
        assert(host.fields["SmallCode"]=="6" && host.fields["WideCode"]=="6");
        auto result=executeVbaProcedure(module,"RoundedResult",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==4);
        result=executeVbaProcedure(module,"InvalidResult",host); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==6);
        std::vector<VbaValue> args{VbaValue::numeric(1.5)};
        result=executeVbaProcedureWithArguments(module,"Consume",host,args); assert(result.status==VbaExecutionStatus::Completed && host.fields["Seen"]=="2" && args[0].asNumber()==1.5);
        args={VbaValue{}}; result=executeVbaProcedureWithArguments(module,"Consume",host,args); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==94 && host.fields["Seen"]=="2");
        FakeForm mismatch; result=executeVbaProcedure(module,"Mismatch",mismatch); assert(result.status==VbaExecutionStatus::Completed && mismatch.fields["Code"]=="13" && mismatch.fields["Seen"]=="3" && !mismatch.fields.contains("Body"));
        FakeForm temporary; result=executeVbaProcedure(module,"Temporary",temporary); assert(result.status==VbaExecutionStatus::Completed && temporary.fields["Seen"]=="3" && temporary.fields["Body"]=="entered");
        result=executeVbaProcedure(module,"ReferenceOverflow",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==2147483647 && host.fields["Code"]=="6");
        result=executeVbaProcedure(module,"CountTyped",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==6 && host.fields["Counter"]=="4");
        result=executeVbaProcedure(module,"OverflowCounter",host); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==6);
        FakeForm fractional; result=executeVbaProcedure(module,"FractionalCounter",fractional); assert(result.status==VbaExecutionStatus::Completed && fractional.fields["Body"]=="entered");
        auto sequence=std::make_shared<TestSequence>(); sequence->values={VbaValue::numeric(1)}; args={VbaValue::reference(sequence)};
        for(const auto* procedure:{"InvalidEnumerator","InvalidSet"}) {
            FakeForm invalid; result=executeVbaProcedureWithArguments(module,procedure,invalid,args);
            assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251 && !invalid.fields.contains("Body"));
        }
    }

    {
        const std::string module=R"VBA(Sub LocalValues()
Dim flag As Boolean
Dim label As String
flag="False"
label=42
Me!Converted=CStr(flag)&":"&label
On Error Resume Next
flag=Null
Me!FlagCode=CStr(Err.Number)
label=Null
Me!TextCode=CStr(Err.Number)
Me!Preserved=CStr(flag)&":"&label
End Sub
Function BooleanResult() As Boolean
BooleanResult=1.5
End Function
Function StringResult() As String
StringResult=True
End Function
Function Combine(ByVal flag As Boolean,ByVal label As String) As String
Me!Body="entered"
If flag Then
Combine=label
Else
Combine="off"
End If
End Function
Sub FlipBoolean(ByRef value As Boolean)
value=Not value
End Sub
Sub ExtendString(ByRef value As String)
value=value&"!"
End Sub
Function References() As String
Dim flag As Boolean
Dim label As String
flag=0
label=7
Call FlipBoolean(flag)
Call ExtendString(label)
References=CStr(flag)&":"&label
End Function
Sub BadFor()
Dim flag As Boolean
Me!Body="entered"
For flag=0 To 1
Next flag
End Sub
Sub BadEnumerator(items As Object)
Dim item As String
Me!Body="entered"
For Each item In items
Next item
End Sub
)VBA";
        FakeForm host; auto result=executeVbaProcedure(module,"LocalValues",host);
        assert(result.status==VbaExecutionStatus::Completed && host.fields["Converted"]=="False:42" && host.fields["Preserved"]=="False:42" && host.fields["FlagCode"]=="94" && host.fields["TextCode"]=="94");
        result=executeVbaProcedure(module,"BooleanResult",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Boolean && result.typedReturnValue->boolean);
        result=executeVbaProcedure(module,"StringResult",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Text && result.typedReturnValue->text=="True");
        std::vector<VbaValue> args{VbaValue::string("False"),VbaValue::numeric(42)};
        result=executeVbaProcedureWithArguments(module,"Combine",host,args);
        assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="off" && args[0].kind==VbaValue::Kind::Text && args[1].kind==VbaValue::Kind::Number);
        args={VbaValue::numeric(1),VbaValue::logical(true)}; result=executeVbaProcedureWithArguments(module,"Combine",host,args); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="True");
        FakeForm invalid; args={VbaValue{},VbaValue::numeric(42)}; result=executeVbaProcedureWithArguments(module,"Combine",invalid,args); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==94 && !invalid.fields.contains("Body"));
        args={VbaValue::logical(true),VbaValue::errorValue(11)}; result=executeVbaProcedureWithArguments(module,"Combine",invalid,args); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==13 && !invalid.fields.contains("Body"));
        result=executeVbaProcedure(module,"References",host); assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="True:7!");
        result=executeVbaProcedure(module,"BadFor",invalid); assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251 && !invalid.fields.contains("Body"));
        auto sequence=std::make_shared<TestSequence>(); sequence->values={VbaValue::string("value")}; args={VbaValue::reference(sequence)};
        result=executeVbaProcedureWithArguments(module,"BadEnumerator",invalid,args); assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251 && !invalid.fields.contains("Body"));
    }

    {
        FakeForm form;
        const std::string module=R"VBA(Function Calculate() As Variant
Dim amount As Variant
amount=CCur("1")
Call Increment(amount)
Calculate=amount
End Function
Sub Increment(ByRef amount As Variant)
amount=amount+CCur("0.0001")
End Sub
Function ValueCopy() As Variant
Dim amount As Variant
amount=CCur("1")
Call IncrementCopy(amount)
ValueCopy=amount
End Function
Sub IncrementCopy(ByVal amount As Variant)
amount=amount+CCur("0.0001")
End Sub
Function SharedSlot() As Variant
Dim amount As Variant
amount=CCur("1")
Call Transfer(amount,amount)
SharedSlot=amount
End Function
Sub Transfer(ByRef debit As Variant,ByRef credit As Variant)
debit=debit-CCur("0.0001")
credit=credit+CCur("0.0001")
End Sub
Function BeforeError() As Variant
Dim amount As Variant
amount=CCur("1")
On Error Resume Next
Call ChangeThenFail(amount)
BeforeError=amount
End Function
Sub ChangeThenFail(ByRef amount As Variant)
amount=amount+CCur("0.0001")
Err.Raise 5
End Sub
)VBA";
        const auto verify=[&](const char* procedure,std::int64_t units) {
            const auto result=executeVbaProcedure(module,procedure,form);
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue);
            assert(result.typedReturnValue->kind==VbaValue::Kind::Currency && result.typedReturnValue->currencyUnits==units);
        };
        verify("Calculate",10001); verify("ValueCopy",10000); verify("SharedSlot",10000); verify("BeforeError",10001);
    }

    {
        FakeForm host;
        assert(evaluateVbaExpression("MsgBox(\"plain\")",host,{}).asNumber()==1);
        try { (void)evaluateVbaExpression("MsgBox(\"styled\",vbYesNo)",host,{}); assert(false); }
        catch (const VbaRuntimeError& error) { assert(error.number==3251); }
        host.dialogHost=true; host.dialogResult=7;
        assert(evaluateVbaExpression("MsgBox(\"question\",vbYesNo+vbQuestion+vbDefaultButton2)",host,{}).asNumber()==7);
        assert(host.dialogButtons==292 && !host.dialogTitle);
        assert(evaluateVbaExpression("MsgBox(\"optional\",,\"\",,)",host,{}).asNumber()==7);
        assert(host.dialogButtons==0 && host.dialogTitle && host.dialogTitle->empty());
        for (const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
            {"MsgBox()",450},{"MsgBox(,)",450},{"MsgBox(Null)",94},
            {"MsgBox(\"x\",Null)",94},{"MsgBox(\"x\",,Null)",94},
            {"MsgBox(\"x\",,,\"help\")",3251},{"MsgBox(\"x\",,,,,0)",450}}) {
            try { (void)evaluateVbaExpression(expression,host,{}); assert(false); }
            catch (const VbaRuntimeError& error) { assert(error.number==code); }
        }
        try { (void)evaluateVbaExpression("InputBox(\"prompt\")",host,{}); assert(false); }
        catch (const VbaRuntimeError& error) { assert(error.number==3251); }
        host.inputHost=true; host.inputResult="answer";
        assert(evaluateVbaExpression("InputBox(\"prompt\")",host,{}).asText()=="answer");
        assert(host.inputRequest && host.inputRequest->prompt=="prompt" && !host.inputRequest->title && host.inputRequest->defaultValue.empty());
        assert(!host.inputRequest->xTwips && !host.inputRequest->yTwips);
        assert(evaluateVbaExpression("InputBox$(\"prompt\",,\"seed\",1440,720,,)",host,{}).asText()=="answer");
        assert(!host.inputRequest->title && host.inputRequest->defaultValue=="seed" && host.inputRequest->xTwips==1440 && host.inputRequest->yTwips==720);
        (void)evaluateVbaExpression("InputBox(\"prompt\",,,2.5,3.5)",host,{});
        assert(host.inputRequest->xTwips==2 && host.inputRequest->yTwips==4);
        host.inputResult="";
        const auto cancelled=evaluateVbaExpression("InputBox(\"prompt\",\"\")",host,{});
        assert(cancelled.kind==VbaValue::Kind::Text && cancelled.asText().empty() && host.inputRequest->title && host.inputRequest->title->empty());
        for (const auto& [expression,code]:std::vector<std::pair<std::string,int>>{
            {"InputBox()",450},{"InputBox(,)",450},{"InputBox(Null)",94},{"InputBox(\"x\",Null)",94},
            {"InputBox(\"x\",,Null)",94},{"InputBox(\"x\",,,Null)",94},{"InputBox(\"x\",,,,Null)",94},
            {"InputBox(\"x\",,,,,'help')",5},{"InputBox(\"x\",,,,,\"help\")",3251},
            {"InputBox(\"x\",,,,,,1)",3251},{"InputBox(\"x\",,,,,,,)",450},
            {"InputBox(\"x\",,,\"bad\")",13},{"InputBox(\"x\",,,2147483648)",6}}) {
            try { (void)evaluateVbaExpression(expression,host,{}); assert(false); }
            catch (const VbaRuntimeError& error) { assert(error.number==code); }
        }
        host.dialogResult=0;
        try { (void)evaluateVbaExpression("MsgBox(\"x\")",host,{}); assert(false); }
        catch (const VbaRuntimeError& error) { assert(error.number==5); }
    }

    {
        FakeForm form;
        const std::string module = R"VBA(
Sub Iterate(ByVal items As Object)
Dim item As Variant
Dim total As Long
Dim count As Long
item = 7
For Each item In items
count = count + 1
total = total + item
Next item
Me!Total = total
Me!Count = count
Me!Last = item
End Sub
)VBA";
        auto sequence = std::make_shared<TestSequence>();
        std::vector<VbaValue> args{VbaValue::reference(sequence)};
        auto result = executeVbaProcedureWithArguments(module,"Iterate",form,args);
        assert(result.status == VbaExecutionStatus::Completed);
        assert(form.fields["Count"] == "0" && form.fields["Last"] == "7");
        sequence->values = {VbaValue::numeric(2),VbaValue::numeric(5)};
        result = executeVbaProcedureWithArguments(module,"Iterate",form,args);
        assert(result.status == VbaExecutionStatus::Completed);
        assert(form.fields["Total"] == "7" && form.fields["Last"] == "5");
        args[0] = VbaValue::reference(nullptr);
        result = executeVbaProcedureWithArguments(module,"Iterate",form,args);
        assert(result.status == VbaExecutionStatus::Error && result.errorNumber == 91);
        args[0] = VbaValue::numeric(2);
        result = executeVbaProcedureWithArguments(module,"Iterate",form,args);
        assert(result.status == VbaExecutionStatus::Error && result.errorNumber == 13);
    }

    assert(validVbaExpression("Format$(\"2003-12-31\", \"yyyy\")"));
    {
        FakeForm unavailable;
        const auto result=executeVbaProcedure("Sub FormatDate()\nMe!Result=Format$(\"2003-12-31\", \"yyyy\")\nEnd Sub","FormatDate",unavailable);
        assert(result.status==VbaExecutionStatus::Error && result.errorNumber==438);
    }
    const auto csv = parseCsv("id,name\r\n1,\"Egger, Nico\"\r\n2,\"A\"\"B\"\r\n");
    assert(csv.size() == 3);
    assert(csv[1][1] == "Egger, Nico");
    assert(csv[2][1] == "A\"B");
    const auto multiline = parseCsv("x,y\n\"a\nb\",c\n");
    assert(multiline[1][0] == "a\nb");
    assert(parseCsv("").empty());
    const auto tsv=parseTsvDetailed("1\t\"a\tb\"\"c\n\"\r\n2\t\r\n");
    assert(tsv.size()==2 && tsv[0].size()==2 && tsv[0][1].quoted && tsv[0][1].value=="a\tb\"c\n" && tsv[1][1].value.empty());
    for(const auto* invalid:{"\"unclosed","\"closed\"tail","unescaped\"quote"}) {
        bool rejected=false; try { (void)parseTsvDetailed(invalid); } catch(const std::runtime_error&) { rejected=true; } assert(rejected);
    }

    const auto detailed = parseCsvDetailed("value,other\n\"\\x\",\x1e\n\"\x1e\",\x1e\n");
    assert(detailed.size() == 3);
    assert(detailed[2][0].quoted && detailed[2][0].value == "\x1e");
    assert(!detailed[2][1].quoted && detailed[2][1].value == "\x1e");

    const auto code = analyzeVba(R"VBA(
Option Explicit
Private Sub Form_Load()
    MsgBox "Hi"
    MsgBox "CreateObject(""fake"")"
End Sub
Private Sub cmdSpeichern_Click()
    DoCmd.RunCommand acCmdSaveRecord
    Set db = CurrentDb()
    ' CreateObject("ignored")
End Sub
Public Function Total() As Long
    Total = 2
End Function
)VBA");
    assert(code.procedures.size() == 3);
    assert(code.procedures[0].name == "Form_Load");
    assert(code.procedures[0].eventCandidate);
    assert(code.procedures[2].name == "Total");
    assert(code.compatibilityFlags.size() == 2);
    const auto localized = analyzeVba(
        "Private Sub Mengenänderung_Click()\nEnd Sub\n"
        "Function set_Stk_Münzliste_zwingend_auf_Null_wenn_leer()\nEnd Function\n"
        "Private Sub Test()\nSet db = CurrentDb\nEnd Sub\n");
    assert(localized.procedures.size() == 3);
    assert(localized.procedures[0].name == "Mengenänderung_Click");
    assert(localized.procedures[1].name == "set_Stk_Münzliste_zwingend_auf_Null_wenn_leer");
    assert(localized.compatibilityFlags.size() == 1);
    FakeForm localizedHost;
    const auto localizedRun = executeVbaProcedure(
        "Private Sub Mengenänderung_Click()\nMsgBox \"Ereignis\"\nEnd Sub",
        "Mengenänderung_Click", localizedHost);
    assert(localizedRun.status == VbaExecutionStatus::Completed);
    assert(localizedHost.messages.size() == 1 && localizedHost.messages[0] == "Ereignis");

    FakeForm host;
    host.fields["Nachname"] = std::nullopt;
    host.fields["Name"] = std::nullopt;
    const std::string vba = R"VBA(
Private Sub Form_Load()
    MsgBox "Formular geöffnet"
End Sub
Private Sub btnSpeichern_Click()
    If IsNull(Me!Nachname) Then
        MsgBox "Bitte Nachnamen eingeben"
        Exit Sub
    End If
    DoCmd.RunCommand acCmdSaveRecord
    MsgBox "Gespeichert!"
End Sub
Private Sub btnAktion_Click()
    Me!Name = "Nico"
    MsgBox "Hallo " & Me!Name
End Sub
)VBA";
    auto loaded = executeVbaProcedure(vba, "Form_Load", host);
    assert(loaded.status == VbaExecutionStatus::Completed);
    assert(host.messages.back() == "Formular geöffnet");
    auto blocked = executeVbaProcedure(vba, "btnSpeichern_Click", host);
    assert(blocked.status == VbaExecutionStatus::Completed && blocked.exitSub);
    assert(host.saves == 0);
    host.fields["Nachname"] = "Egger";
    auto saved = executeVbaProcedure(vba, "btnSpeichern_Click", host);
    assert(saved.status == VbaExecutionStatus::Completed && host.saves == 1);
    executeVbaProcedure(vba, "btnAktion_Click", host);
    assert(host.fields["Name"].value() == "Nico");
    assert(host.messages.back() == "Hallo Nico");
    assert(executeVbaProcedure(vba, "Unknown", host).status == VbaExecutionStatus::ProcedureMissing);
    const auto unsupported = executeVbaProcedure(
        "Sub Test()\nMsgBox \"One\"\nCreateObject(\"X\")\nEnd Sub", "Test", host);
    assert(unsupported.status == VbaExecutionStatus::Unsupported);
    assert(host.messages.back() == "Hallo Nico"); // preflight stopped side effects
    const auto malformed = executeVbaProcedure(
        "Sub Test()\nIf IsNull(Me!Name) Then\nElse\nElse\nEnd If\nEnd Sub", "Test", host);
    assert(malformed.status == VbaExecutionStatus::Error);

    const auto moreVba = R"VBA(
Option Explicit
Private Sub Berechnen()
    Dim i As Long
    Dim sum As Long
    sum = 0
    For i = 1 To 4
        sum = sum + i
    Next i
    If sum = 10 And Not IsNull(Me![Name]) Then
        Me![Name] = "Sum=" & CStr(sum)
        MsgBox "Fertig: " & Me![Name]
    Else
        MsgBox "Fehler"
    End If
End Sub
Public Function Anzahl() As Long
    Dim index As Long
    Dim counter As Long
    counter = 0
    For index = 5 To 1 Step -2
        counter = counter + 1
    Next index
    Anzahl = counter
End Function
)VBA";
    const auto calc = executeVbaProcedure(moreVba, "Berechnen", host);
    if (calc.status != VbaExecutionStatus::Completed)
        std::cerr << "VBA calculation error: " << calc.explanation << " @" << calc.line << '\n';
    assert(calc.status == VbaExecutionStatus::Completed);
    assert(host.fields["Name"].value() == "Sum=10");
    assert(host.messages.back() == "Fertig: Sum=10");
    const auto count = executeVbaProcedure(moreVba, "Anzahl", host);
    assert(count.status == VbaExecutionStatus::Completed && count.returnValue == "3");
    const auto budget = executeVbaProcedure(
        "Sub Endless()\nDim i As Long\nFor i = 1 To 99999999\nNext i\nEnd Sub",
        "Endless", host);
    assert(budget.status == VbaExecutionStatus::Error);
    assert(budget.explanation.find("10.000") != std::string::npos);
    const auto blockedCall = executeVbaProcedure(
        "Sub Test()\nMsgBox \"Darf nicht laufen\"\nShell(\"test\")\nEnd Sub", "Test", host);
    assert(blockedCall.status == VbaExecutionStatus::Unsupported);
    assert(host.messages.back() == "Fertig: Sum=10");
    const auto accessOpen = executeVbaProcedure(
        "Private Sub Form_Open(Cancel As Integer)\n"
        "DoCmd.Maximize\nDoCmd.GoToRecord , , acLast\nEnd Sub",
        "Form_Open", host);
    assert(accessOpen.status == VbaExecutionStatus::Completed);
    assert(host.maximized == 1 && host.navigatedLast == 1);
    const auto accessCancel = executeVbaProcedure(
        "Private Sub Form_Open(Cancel As Integer)\nCancel=1\nEnd Sub",
        "Form_Open", host);
    assert(accessCancel.status == VbaExecutionStatus::Completed && accessCancel.cancelled);
    const auto accessActions = executeVbaProcedure(R"VBA(
Private Sub cmdActions_Click()
Dim tableName As String
tableName = "recipes"
DoCmd.RunSQL "DELETE FROM " & tableName & " WHERE id = 4"
CurrentDb.Execute "INSERT INTO recipes (id) VALUES (5)"
DoCmd.GoToRecord , , acNext
DoCmd.GoToRecord , , acNewRec
DoCmd.Requery
DoCmd.OpenForm "frmRezepte"
End Sub
)VBA", "cmdActions_Click", host);
    assert(accessActions.status == VbaExecutionStatus::Completed);
    assert(host.sqlStatements.size() == 2);
    assert(host.sqlStatements[0] == "DELETE FROM recipes WHERE id = 4");
    assert(host.sqlStatements[1] == "INSERT INTO recipes (id) VALUES (5)");
    assert(host.navigatedNext == 1 && host.navigatedNew == 1 && host.requeries == 1);
    assert(host.openedForms.size() == 1 && host.openedForms[0] == "frmRezepte");
    assert(executeVbaProcedure("Sub OpenDefault()\nDoCmd.OpenForm \"form\",acNormal,,\"\"\nEnd Sub", "OpenDefault", host).status==VbaExecutionStatus::Completed);
    assert(host.openedForms.back()=="form");
    assert(executeVbaProcedure("Sub OpenFiltered()\nDoCmd.OpenForm \"form\",,,\"id=2\"\nEnd Sub", "OpenFiltered", host).errorNumber==3251);
    const std::string bracketCode="Sub Bracket()\nDim flag As Boolean\nflag=False\nIf [flag]=-1 Then\nMe.Tag=\"selected\"\nElse\nMe.Tag=\"clear\"\nEnd If\nEnd Sub";
    assert(validateVbaProcedure(bracketCode,"Bracket").status==VbaExecutionStatus::Completed);
    assert(executeVbaProcedure(bracketCode,"Bracket",host).errorNumber==438);
    host.bracketHost=true; host.bracketFields["flag"]=VbaValue::numeric(-1);
    assert(executeVbaProcedure(bracketCode,"Bracket",host).status==VbaExecutionStatus::Completed && host.fields["Tag"]=="selected");
    host.bracketFields["flag"]=VbaValue::numeric(0);
    assert(executeVbaProcedure(bracketCode,"Bracket",host).status==VbaExecutionStatus::Completed && host.fields["Tag"]=="clear");
    host.bracketHost=false;
    for (const auto& [declaration,kind]:std::vector<std::pair<std::string,VbaValue::Kind>>{{"As Boolean",VbaValue::Kind::Boolean},{"As String",VbaValue::Kind::Text},{"As Long",VbaValue::Kind::Number},{"",VbaValue::Kind::Empty},{"As Object",VbaValue::Kind::Object}}) {
        const auto outcome=executeVbaProcedure("Function DefaultResult() "+declaration+"\nEnd Function","DefaultResult",host);
        assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue && outcome.typedReturnValue->kind==kind);
        if (kind==VbaValue::Kind::Boolean) assert(!outcome.typedReturnValue->boolean);
        if (kind==VbaValue::Kind::Number) assert(outcome.typedReturnValue->number==0);
        if (kind==VbaValue::Kind::Object) assert(!outcome.typedReturnValue->object);
    }
    assert(executeVbaProcedure("Function UnsupportedResult() As CustomType\nMe.Tag=\"effect\"\nEnd Function","UnsupportedResult",host).status==VbaExecutionStatus::Unsupported);
    const auto inlineExit=executeVbaProcedure(R"VBA(Sub InlineExit()
Dim total As Long
Dim outer As Long
Dim inner As Long
For outer=1 To 3
For inner=1 To 5
If inner=2 Then total=total+10: Exit For
total=total+1
Next
total=total+100
Next
Me!LoopResult=total
End Sub)VBA","InlineExit",host);
    assert(inlineExit.status==VbaExecutionStatus::Completed && host.fields["LoopResult"]=="333");
    assert(executeVbaProcedure("Sub InvalidExit()\nMe!LoopResult=0\nIf True Then Exit For\nEnd Sub","InvalidExit",host).status!=VbaExecutionStatus::Completed);
    assert(host.fields["LoopResult"]=="333");
    const auto nestedCall = executeVbaProcedure(R"VBA(
Private Sub btn_Click()
Call InternerHandler
Me!Anzeige.Visible = False
Me!Anzeige.SetFocus
Me!Anzeige.Requery
End Sub
Private Sub InternerHandler()
MsgBox "Unterprogramm"
End Sub
)VBA", "btn_Click", host);
    assert(nestedCall.status == VbaExecutionStatus::Completed);
    assert(host.messages.back() == "Unterprogramm");
    assert(host.visibility["Anzeige"] == false);
    assert(host.controlsFocused == 1 && host.controlsRefreshed == 1);
    const auto joinedSql = executeVbaProcedure(
        "Sub Test()\nCurrentDb.Execute \"DELETE FROM \" & _\n"
        "  \"recipes WHERE id = 1\"\nEnd Sub", "Test", host);
    assert(joinedSql.status == VbaExecutionStatus::Completed);
    assert(host.sqlStatements.back() == "DELETE FROM recipes WHERE id = 1");
    const auto builtins = executeVbaProcedure(R"VBA(
Sub Builtins()
Me!Name = Left("Keksrezepte", 4) & Right("Brot", 2) & Mid("abcdef", 3, 2)
Me!Nachname = IIf(IsNumeric("123"), CStr(Int(3.9)), "Fehler")
MsgBox Now()
End Sub
)VBA", "Builtins", host);
    assert(builtins.status == VbaExecutionStatus::Completed);
    assert(host.fields["Name"] == "Keksotcd");
    assert(host.fields["Nachname"] == "3");
    assert(host.messages.back().size()==19 || host.messages.back().size()==10);
    const auto bareClocks=executeVbaProcedure(R"VBA(
Sub Clocks()
Me!Name = Date
Me!Nachname = Time
MsgBox Now
End Sub
)VBA","Clocks",host);
    assert(bareClocks.status==VbaExecutionStatus::Completed);
    assert(host.fields["Name"]->size()==10 && host.fields["Name"]->at(4)=='-' && host.fields["Name"]->at(7)=='-');
    assert(host.fields["Nachname"]->size()==8 && host.fields["Nachname"]->at(2)==':' && host.fields["Nachname"]->at(5)==':');
    assert((host.messages.back().size()==19 && host.messages.back()[10]==' ') || host.messages.back().size()==10);
    const auto badClock=executeVbaProcedure(R"VBA(
Sub ClockError()
On Error Resume Next
Me!Name = Date(1)
Me!Nachname = CStr(Err.Number)
End Sub
)VBA","ClockError",host);
    assert(badClock.status==VbaExecutionStatus::Completed && host.fields["Nachname"]=="450");
    const auto errorVariants=executeVbaProcedure(R"VBA(
Sub VariantErrors()
Dim problem As Variant
problem = CVErr(2015)
Me!Name = CStr(IsError(problem)) & ":" & CStr(problem) & ":" & CStr(Err.Number)
Call ReplaceError(problem)
Me!Nachname = CStr(problem)
Me!Returned = CStr(MakeError())
On Error Resume Next
problem = problem + 1
Me!Arithmetic = CStr(Err.Number)
Err.Clear
problem = CVErr(-1)
Me!InvalidError = CStr(Err.Number)
problem = CVErr()
Me!ArgumentError = CStr(Err.Number)
End Sub
Sub ReplaceError(ByRef value As Variant)
value = CVErr(2020)
End Sub
Function MakeError() As Variant
MakeError = CVErr(2042)
End Function
)VBA","VariantErrors",host);
    assert(errorVariants.status==VbaExecutionStatus::Completed);
    assert(host.fields["Name"]=="True:Error 2015:0" && host.fields["Nachname"]=="Error 2020");
    assert(host.fields["Returned"]=="Error 2042" && host.fields["Arithmetic"]=="13" && host.fields["InvalidError"]=="5");
    assert(host.fields["ArgumentError"]=="450");
    const auto labelHandler = executeVbaProcedure(R"VBA(
Sub Labels()
Dim divisor As Long
Dim answer As Long
On Error GoTo Recover
divisor = 0
answer = 12 / divisor
Me!Name = CStr(answer)
GoTo Finished
Recover:
Me!Nachname = CStr(Err.Number)
divisor = 3
Resume
Finished:
Exit Sub
End Sub
)VBA", "Labels", host);
    assert(labelHandler.status == VbaExecutionStatus::Completed);
    assert(host.fields["Name"] == "4" && host.fields["Nachname"] == "11");
    const auto nestedError = executeVbaProcedure(R"VBA(
Sub Outer()
On Error GoTo Handler
Inner
Exit Sub
Handler:
Me!Name = CStr(Err.Number)
Resume Done
Done:
Exit Sub
End Sub
Sub Inner()
Err.Raise(6)
End Sub
)VBA", "Outer", host);
    assert(nestedError.status == VbaExecutionStatus::Completed && host.fields["Name"] == "6");
    assert(executeVbaProcedure("Sub Test()\nResume Next\nEnd Sub", "Test", host).errorNumber == 20);
    const auto functions = executeVbaProcedure(R"VBA(
Sub Functions()
Dim n As Long
n = 2
Me!Name = CStr(Twice(n) + PlusOne(n))
Me!Nachname = CStr(n)
End Sub
Function Twice(ByVal value As Long) As Long
Twice = value * 2
End Function
Function PlusOne(ByRef value As Long) As Long
value = value + 1
PlusOne = value
End Function
)VBA", "Functions", host);
    assert(functions.status == VbaExecutionStatus::Completed);
    assert(host.fields["Name"] == "7" && host.fields["Nachname"] == "3");
    const auto aliased = executeVbaProcedure(R"VBA(
Sub Aliases()
Dim n As Long
n = 1
Touch n, n
Me!Name = CStr(n)
End Sub
Sub Touch(ByRef a As Long, ByRef b As Long)
a = a + 1
b = b + a
End Sub
)VBA", "Aliases", host);
    assert(aliased.status == VbaExecutionStatus::Completed && host.fields["Name"] == "4");
    const auto parameterCalls = executeVbaProcedure(R"VBA(
Sub Parameters()
Dim n As Long
n = 2
Increment n
Call Increment(n)
IgnoreChange n
Me!Name = CStr(n)
KeyCapt "7,8"
End Sub
Sub Increment(ByRef value As Long)
value = value + 1
End Sub
Sub IgnoreChange(ByVal value As Long)
value = 99
End Sub
Sub KeyCapt(C)
Me!Nachname = C
End Sub
)VBA", "Parameters", host);
    assert(parameterCalls.status == VbaExecutionStatus::Completed);
    assert(host.fields["Name"] == "4");
    assert(host.fields["Nachname"] == "7,8");
    const auto handled = executeVbaProcedure(R"VBA(
Sub Errors()
Dim untyped, typed As Long
If Not IsEmpty(untyped) Or IsEmpty(typed) Then
    Err.Raise(5)
End If
On Error Resume Next
typed = 1 / 0
Me!Name = CStr(Err.Number)
Err.Clear
Me!Nachname = CStr(Err.Number)
End Sub
)VBA", "Errors", host);
    assert(handled.status == VbaExecutionStatus::Completed);
    assert(host.fields["Name"] == "11" && host.fields["Nachname"] == "0");
    const std::string detailedErrors=R"VBA(Sub DetailedErrors()
Dim result As Long
On Error GoTo caught
result=DetailedFailure()
Me!ClearedError=CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description & ":" & Err.HelpFile & ":" & CStr(Err.HelpContext)
Exit Sub
caught:
Me!DetailedError=CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description & ":" & Err.HelpFile & ":" & CStr(Err.HelpContext)
Resume Next
End Sub
Function DetailedFailure() As Long
Err.Raise Number:=vbObjectError+513, Source:="native.error", Description:="detail", HelpFile:="local.chm", HelpContext:=42
End Function
Sub AssignedError()
On Error Resume Next
Err.Source="assigned"
Err.Description="stored"
Err.HelpFile="assigned.chm"
Err.HelpContext=7
Err.Number=123
Err.Raise 1001,,"explicit"
Me!AssignedError=CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description & ":" & Err.HelpFile & ":" & CStr(Err.HelpContext)
Err.Clear
Me!AssignedClear=CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description & ":" & Err.HelpFile & ":" & CStr(Err.HelpContext)
End Sub
)VBA";
    const auto detailedErrorRun=executeVbaProcedure(detailedErrors,"DetailedErrors",host);
    assert(detailedErrorRun.status==VbaExecutionStatus::Completed);
    assert(host.fields["DetailedError"]=="-2147220991:native.error:detail:local.chm:42");
    assert(host.fields["ClearedError"]=="0::::0");
    const auto directError=executeVbaProcedure(detailedErrors,"DetailedFailure",host);
    assert(directError.status==VbaExecutionStatus::Error && directError.errorNumber==-2147220991);
    assert(directError.explanation=="detail" && directError.errorSource=="native.error" && directError.errorHelpFile=="local.chm" && directError.errorHelpContext==42);
    const auto eventError=executeVbaEventExpression(detailedErrors,"=DetailedFailure()",host);
    assert(eventError.status==VbaExecutionStatus::Error && eventError.errorNumber==directError.errorNumber && eventError.errorSource==directError.errorSource && eventError.errorHelpFile==directError.errorHelpFile && eventError.errorHelpContext==42);
    assert(executeVbaProcedure(detailedErrors,"AssignedError",host).status==VbaExecutionStatus::Completed);
    assert(host.fields["AssignedError"]=="1001:assigned:explicit:assigned.chm:7" && host.fields["AssignedClear"]=="0::::0");
    assert(evaluateVbaExpression("VarType(vbObjectError)",host,{}).asNumber()==3);
    const auto errorStatement=executeVbaProcedure(R"VBA(Sub ErrorStatement()
On Error Resume Next
Error vbObjectError+514
Me!StatementError=CStr(Err.Number) & ":" & TypeName(Err)
Err.Clear
Error (11)
Me!StatementDivision=CStr(Err.Number)
End Sub
)VBA","ErrorStatement",host);
    assert(errorStatement.status==VbaExecutionStatus::Completed);
    const auto shadowedErr=executeVbaProcedure("Sub ShadowError()\nDim Err As Long\nError 11\nEnd Sub","ShadowError",host);
    assert(shadowedErr.status==VbaExecutionStatus::Error && shadowedErr.errorNumber==11);
    assert(validateVbaProcedure("Sub InvalidError()\nError 11,22\nEnd Sub","InvalidError").status==VbaExecutionStatus::Unsupported);
    assert(host.fields["StatementError"]=="-2147220990:ErrObject" && host.fields["StatementDivision"]=="11");
    for(const auto expression:{"-Null","+Null","-(Null+1)","+(Null+1)"})
        assert(evaluateVbaExpression(expression,host,{}).kind==VbaValue::Kind::Null);
    const auto shadowedFunctions=executeVbaProcedure(R"VBA(Private calls As Long
Function Len(Optional ByVal value As String="fallback") As Long
calls=calls+1
Len=VBA.Len(value)+10
End Function
Function Round(Optional ByVal first As Variant, Optional ByVal second As Variant) As String
calls=calls+1
Round=CStr(IsMissing(first)) & ":" & CStr(IsMissing(second))
End Function
Function Left() As Long
calls=calls+1
Left=7
End Function
Sub ShadowedFunctions()
Me!ShadowedFunctions=CStr(Len()) & ":" & CStr(Len(value:="abc")) & ":" & CStr(VBA.Len("abc")) & ":" & Round(,2) & ":" & Left("xyz",1) & ":" & CStr(Left()) & ":" & CStr(calls)
End Sub
)VBA","ShadowedFunctions",host);
    if(shadowedFunctions.status!=VbaExecutionStatus::Completed) std::cerr<<"Function shadow error="<<shadowedFunctions.errorNumber<<" line="<<shadowedFunctions.line<<" explanation="<<shadowedFunctions.explanation<<"\n";
    assert(shadowedFunctions.status==VbaExecutionStatus::Completed && host.fields["ShadowedFunctions"]=="18:13:3:True:False:x:7:4");
    const auto subAsValue=executeVbaProcedure("Sub ValueSub()\nMe!NeverValueSub=1\nEnd Sub\nSub CheckSubValue()\nOn Error Resume Next\nMe!IgnoredSubValue=ValueSub()\nMe!SubValueError=CStr(Err.Number)\nEnd Sub","CheckSubValue",host);
    assert(subAsValue.status==VbaExecutionStatus::Completed && !host.fields.contains("NeverValueSub") && host.fields["SubValueError"]=="3251");
    const auto nativeTypeChecks=executeVbaProcedure(R"VBA(Private inspected As Long
Private owned As VBA.Collection
Function RetainedObject() As Object
inspected=inspected+1
Set RetainedObject=owned
End Function
Sub NativeTypeChecks()
Dim dictionary As Scripting.Dictionary, slot As Variant, absent As Object
Set owned=New Collection
Set dictionary=New Scripting.Dictionary
dictionary.Add "items",owned
Set slot=owned
If Not TypeOf dictionary("items") Is VBA.Collection Then Err.Raise 5
If Not TypeOf slot Is Collection Then Err.Raise 5
If TypeOf dictionary Is Collection Then Err.Raise 5
If Not TypeOf RetainedObject() Is Object Then Err.Raise 5
slot=7
If TypeOf slot Is Object Then Err.Raise 5
If Not TypeOf Err Is VBA.ErrObject Then Err.Raise 5
On Error Resume Next
slot=TypeOf absent Is Object
Me!NativeTypeNothing=CStr(Err.Number)
On Error GoTo 0
Me!NativeTypeChecks=CStr(TypeOf owned Is Collection) & ":" & CStr(TypeOf dictionary Is Scripting.Dictionary) & ":" & CStr(inspected)
End Sub
)VBA","NativeTypeChecks",host);
    if(nativeTypeChecks.status!=VbaExecutionStatus::Completed) std::cerr<<"Native TypeOf error="<<nativeTypeChecks.errorNumber<<" line="<<nativeTypeChecks.line<<" explanation="<<nativeTypeChecks.explanation<<"\n";
    assert(nativeTypeChecks.status==VbaExecutionStatus::Completed && host.fields["NativeTypeChecks"]=="True:True:1" && host.fields["NativeTypeNothing"]=="91");
    const auto unknownType=executeVbaProcedure("Sub UnknownType()\nMe!NeverTypeOf=1\nDim obj As Object\nIf TypeOf obj Is UnregisteredClass Then Exit Sub\nEnd Sub","UnknownType",host);
    assert(unknownType.status==VbaExecutionStatus::Unsupported && !host.fields.contains("NeverTypeOf"));
    const auto computedBranches=executeVbaProcedure(R"VBA(Private picks As Long
Function PickBranch() As Double
picks=picks+1
PickBranch=1.5
End Function
Sub ComputedBranches()
Dim items As Collection, item As Variant, total As Long, selector As Variant, repaired As Long
Set items=New Collection
items.Add 1
items.Add 2
For Each item In items
On PickBranch() GoSub Wrong, 0200
Next
selector="bad"
On Error GoTo Repair
On selector GoSub Wrong
On Error GoTo 0
On 0.5 GoTo Wrong
On 3 GoSub Wrong, Wrong
On Error Resume Next
On -1 GoTo Wrong
Me!ComputedErrors=CStr(Err.Number)
Err.Clear
On 256 GoSub Wrong
Me!ComputedErrors=Me!ComputedErrors & ":" & CStr(Err.Number)
Err.Clear
On 32768 GoTo Wrong
Me!ComputedErrors=Me!ComputedErrors & ":" & CStr(Err.Number)
Err.Clear
On Null GoTo Wrong
Me!ComputedErrors=Me!ComputedErrors & ":" & CStr(Err.Number)
On Error GoTo 0
For Each item In items
With items
If item=1 Then On 1 GoTo Released
End With
Next
Released:
items.Add 3
Me!ComputedBranches=CStr(total) & ":" & CStr(picks) & ":" & CStr(items.Count) & ":" & CStr(repaired)
Exit Sub
Wrong:
Err.Raise 5
Exit Sub
200 total=total+item
If total=1 Then On 1 GoSub Ten, Wrong
Return
Ten:
total=total+10
Return
Repair:
selector=0
repaired=repaired+1
Resume
End Sub
)VBA","ComputedBranches",host);
    if(computedBranches.status!=VbaExecutionStatus::Completed) std::cerr<<"Computed branch error="<<computedBranches.errorNumber<<" line="<<computedBranches.line<<" explanation="<<computedBranches.explanation<<"\n";
    assert(computedBranches.status==VbaExecutionStatus::Completed && host.fields["ComputedBranches"]=="13:2:3:1" && host.fields["ComputedErrors"]=="5:5:6:94");
    const auto missingComputed=executeVbaProcedure("Sub MissingComputed()\nMe!NeverComputed=1\nOn 0 GoTo Missing\nEnd Sub","MissingComputed",host);
    assert(missingComputed.status==VbaExecutionStatus::Error && !host.fields.contains("NeverComputed"));
    assert(validateVbaProcedure("Sub InboundComputed()\nOn 0 GoTo Inner\nFor i=1 To 2\nInner: Exit For\nNext\nEnd Sub","InboundComputed").status==VbaExecutionStatus::Unsupported);
    const auto goSubResult=executeVbaProcedure(R"VBA(Sub GoSubScopes()
Dim items As Collection, item As Variant, total As Long, held As Long, i As Long
Set items=New Collection
items.Add 1
items.Add 2
For Each item In items
If item>0 Then GoSub 0100
Next
items.Add 4
Me!GoSubValues=CStr(total) & ":" & CStr(items.Count) & ":" & CStr(held)
Exit Sub
100 On Error Resume Next
items.Add 99
held=Err.Number
On Error GoTo 0
For i=1 To 2
GoSub 200
Next
Return
200 total=total+item
GoTo 0300
300: Return
End Sub
)VBA","GoSubScopes",host);
    if(goSubResult.status!=VbaExecutionStatus::Completed) std::cerr<<"GoSub error="<<goSubResult.errorNumber<<" line="<<goSubResult.line<<"\n";
    assert(goSubResult.status==VbaExecutionStatus::Completed && host.fields["GoSubValues"]=="6:3:3251");
    const auto missingGoSub=executeVbaProcedure("Sub MissingGoSub()\nReturn\nEnd Sub","MissingGoSub",host);
    assert(missingGoSub.status==VbaExecutionStatus::Error && missingGoSub.errorNumber==3);
    const auto goSubHandler=executeVbaProcedure(R"VBA(Sub HandledGoSub()
Dim items As Collection,item As Variant,total As Long,handled As Long,text As String
Set items=New Collection
items.Add 1
items.Add 2
text="bad"
On Error GoTo 1000
For Each item In items
GoSub Accumulate
Next
items.Add 4
Me!HandledGoSub=CStr(total) & ":" & CStr(handled) & ":" & CStr(items.Count)
Exit Sub
Accumulate:
total=total+CLng(text)+item
Return
1000:
handled=handled+1
text="2"
GoSub InspectLease
Resume 0
InspectLease:
On Error Resume Next
items.Add 99
Me!GoSubFaultLease=CStr(Err.Number)
On Error GoTo 1000
Return
End Sub
)VBA","HandledGoSub",host);
    if(goSubHandler.status!=VbaExecutionStatus::Completed) std::cerr<<"Handled GoSub error="<<goSubHandler.errorNumber<<" line="<<goSubHandler.line<<"\n";
    assert(goSubHandler.status==VbaExecutionStatus::Completed && host.fields["HandledGoSub"]=="7:1:3" && host.fields["GoSubFaultLease"]=="3251");
    const auto numericResume=executeVbaProcedure("Sub NumericResume()\nOn Error GoTo 900\nError 13\nExit Sub\n700 Me!NumericResume=\"yes\"\nExit Sub\n900: Resume 0700\nEnd Sub","NumericResume",host);
    assert(numericResume.status==VbaExecutionStatus::Completed && host.fields["NumericResume"]=="yes");
    assert(validateVbaProcedure("Sub NumericRange()\n65530: Exit Sub\nEnd Sub","NumericRange").status==VbaExecutionStatus::Unsupported);
    const auto unhandled = executeVbaProcedure(
        "Sub Errors()\nOn Error Resume Next\nErr.Raise(5)\nOn Error GoTo 0\nErr.Raise(6)\nEnd Sub", "Errors", host);
    assert(unhandled.status == VbaExecutionStatus::Error);
    const VbaVariables variables;
    const auto evaluate = [&](std::string_view source) { return evaluateVbaExpression(source, host, variables); };
    assert(evaluate("Empty + 2").asNumber() == 2);
    const auto currencyCheck=[&](std::string_view text,std::string_view expected) {
        const auto value=evaluate(text);
        assert(value.kind==VbaValue::Kind::Currency && value.asText()==expected);
    };
    const std::string typedCurrencyModule=R"VBA(Private amount As Currency
Public balance As Currency
Function TypedCurrency(ByVal input As Currency) As Currency
Dim localAmount As Currency
localAmount=input
amount=localAmount
balance=localAmount
TypedCurrency=localAmount+CCur("0.0001")
End Function
Function DefaultCurrency() As Currency
Dim untouched As Currency
DefaultCurrency=untouched
End Function
)VBA";
    assert(host.moduleState(typedCurrencyModule)->values.at("amount").kind==VbaValue::Kind::Currency);
    assert(host.moduleState(typedCurrencyModule)->values.at("amount").currencyUnits==0);
    std::vector<VbaValue> typedCurrencyArgs{VbaValue::string("922337203685477.5806")};
    auto typedCurrencyResult=executeVbaProcedureWithArguments(typedCurrencyModule,"TypedCurrency",host,typedCurrencyArgs);
    assert(typedCurrencyResult.status==VbaExecutionStatus::Completed && typedCurrencyResult.typedReturnValue->kind==VbaValue::Kind::Currency && typedCurrencyResult.typedReturnValue->currencyUnits==std::numeric_limits<std::int64_t>::max());
    assert(typedCurrencyArgs[0].kind==VbaValue::Kind::Text);
    typedCurrencyResult=executeVbaProcedure(typedCurrencyModule,"DefaultCurrency",host);
    assert(typedCurrencyResult.status==VbaExecutionStatus::Completed && typedCurrencyResult.typedReturnValue->kind==VbaValue::Kind::Currency && typedCurrencyResult.typedReturnValue->currencyUnits==0);
    currencyCheck("CCur(\"922337203685477.5807\")","922337203685477.5807");
    currencyCheck("CCur(\"-922337203685477.5808\")","-922337203685477.5808");
    currencyCheck("CCur(\"1.23445\")","1.2344");
    currencyCheck("CCur(\"1.23455\")","1.2346");
    currencyCheck("CCur(\"-1.23445\")","-1.2344");
    currencyCheck("CCur(\"0.00005001\")","0.0001");
    currencyCheck("CCur(\"4.5e-5\")","0");
    currencyCheck("CCur(True)","-1");
    currencyCheck("CCur(Empty)","0");
    currencyCheck("CCur(1.25)","1.25");
    currencyCheck("Int(CCur(\"-1.25\"))","-2");
    currencyCheck("Int(CCur(\"1.25\"))","1");
    currencyCheck("CCur(\"922337203685477.5806\")+CCur(\"0.0001\")","922337203685477.5807");
    currencyCheck("CCur(\"0.0001\")*CCur(\"0.5\")","0");
    currencyCheck("CCur(\"0.0003\")*CCur(\"0.5\")","0.0002");
    currencyCheck("CCur(\"-0.0003\")*CCur(\"0.5\")","-0.0002");
    currencyCheck("-CCur(\"1.0001\")","-1.0001");
    currencyCheck("Abs(CCur(\"-1.0001\"))","1.0001");
    assert(evaluate("CCur(\"922337203685477.5806\") < CCur(\"922337203685477.5807\")").truthy());
    assert(evaluate("VarType(CCur(1))").asNumber()==6);
    assert(evaluate("TypeName(CCur(1))").asText()=="Currency");
    assert(evaluate("CCur(1)/CCur(2)").asNumber()==0.5);
    assert(evaluate("CCur(1)/CCur(2)").kind==VbaValue::Kind::Number);
    const auto currencyError=[&](std::string_view source,int number) {
        bool caught=false;
        try { (void)evaluate(source); } catch(const VbaRuntimeError& error) { caught=error.number==number; }
        assert(caught);
    };
    currencyError("CCur(\"922337203685477.58075\")",6);
    currencyError("-CCur(\"-922337203685477.5808\")",6);
    currencyError("CCur(\"922337203685477.5807\")+CCur(1)",6);
    currencyError("CCur(Null)",94);
    currencyError("CCur(\"invalid\")",13);
    currencyError("CCur(CVErr(2015))",13);
    currencyError("CCur()",450);
    currencyError("Int(CCur(\"-922337203685477.5808\"))",6);
    assert(evaluate("TypeName(Nothing)").asText()=="Nothing");

    assert(!evaluate("0 = \"\"").truthy());
    assert(!evaluate("2 = \"2\"").truthy());
    assert(evaluate("2 < \"1\"").truthy());
    assert(evaluate("\"\" > -1").truthy());
    assert(evaluate("Empty = \"\"").truthy());
    assert(evaluate("Empty = 0").truthy());
    assert(evaluate("Null = \"\"").kind == VbaValue::Kind::Null);

    assert(evaluate("IsEmpty(Empty)").truthy());
    assert(!evaluate("IsNull(Empty)").truthy());
    assert(evaluate("Null And False").kind == VbaValue::Kind::Boolean);
    assert(!evaluate("Null And False").truthy());
    assert(evaluate("Null Or False").kind == VbaValue::Kind::Null);
    assert(evaluate("Null Or True").truthy());
    assert(evaluate("Not Null").kind == VbaValue::Kind::Null);
    assert(evaluate("Not 1 = 2").truthy());
    assert(evaluate("6 And 3").asNumber() == 2);
    assert(evaluate("4 Or 1").asNumber() == 5);
    assert(evaluate("CBool(Empty)").kind==VbaValue::Kind::Boolean && !evaluate("CBool(Empty)").boolean);
    assert(!evaluate("CBool(0)").boolean && evaluate("CBool(-0.25)").boolean);
    assert(!evaluate("CBool(\"False\")").boolean && evaluate("CBool(\" TrUe \")").boolean);
    assert(evaluate("CBool(\" 2.5 \")").boolean && !evaluate("CBool(\" 0 \")").boolean);
    assert(evaluate("CBool(CCur(\"0.0001\"))").boolean);
    assert(evaluate("VarType(CBool(1))").asNumber()==11 && evaluate("TypeName(CBool(0))").asText()=="Boolean");
    assert(evaluate("CSng(0.1)").asNumber()==static_cast<double>(0.1f));
    assert(evaluate("CSng(16777217)").asNumber()==16777216);
    assert(evaluate("CDbl(16777217)").asNumber()==16777217);
    assert(evaluate("VarType(CSng(1))=vbSingle").truthy() && evaluate("TypeName(CSng(1))").asText()=="Single");
    assert(evaluate("VarType(CDbl(CSng(1)))=vbDouble").truthy());
    assert(evaluate("CStr(CSng(0.1))").asText()=="0.1");
    assert(evaluate("CSng(Empty)").asNumber()==0 && evaluate("CSng(True)").asNumber()==-1);
    assert(evaluate("CInt(\" 2.5 \")").asNumber()==2 && evaluate("CLng(\" 3.5 \")").asNumber()==4);
    assert(evaluate("CInt(2.5)").asNumber()==2 && evaluate("CInt(3.5)").asNumber()==4);
    assert(evaluate("CInt(-2.5)").asNumber()==-2 && evaluate("CInt(True)").asNumber()==-1);
    assert(evaluate("CInt(Empty)").asNumber()==0 && evaluate("CInt(CCur(\"1.5\"))").asNumber()==2);
    assert(evaluate("VarType(CInt(1))=vbInteger").truthy());
    assert(evaluate("VarType(CLng(1))=vbLong").truthy());
    assert(evaluate("VarType(CDbl(CInt(1)))=vbDouble").truthy());
    assert(evaluate("VarType(CInt(1))").asNumber()==2 && evaluate("TypeName(CInt(1))").asText()=="Integer");
    assert(evaluate("VarType(CLng(1))").asNumber()==3 && evaluate("TypeName(CLng(1))").asText()=="Long");
    for(const auto& item:std::vector<std::pair<const char*,double>>{
        {R"(12.6 Mod 5)",3},{R"(12 Mod 4.3)",0},{R"(10 Mod 3 * 2)",4},
        {R"(20 Mod 6 \ 2)",2},{R"(20 \ 6 Mod 3)",0},{R"(20 \ 6 * 2)",1},
        {R"(19 \ 4)",4},{R"(-19 \ 4)",-4},{R"(19 \ -4)",-4},
        {R"(-7 Mod 3)",-1},{R"(7 Mod -3)",1},{R"(4.5 \ 1.5)",2},
        {R"(Empty Mod "5")",0},{R"(Empty \ 5)",0},{R"(-2147483648 Mod -1)",0}})
        assert(evaluate(item.first).asNumber()==item.second);
    assert(evaluate(R"(VarType(CInt(7) Mod CInt(3)))").asNumber()==2);
    assert(evaluate(R"(VarType(CLng(7) \ CLng(3)))").asNumber()==3);
    assert(evaluate(R"(Null Mod 0)").kind==VbaValue::Kind::Null && evaluate(R"(Null \ 0)").kind==VbaValue::Kind::Null);
    for(const auto& item:std::vector<std::pair<const char*,double>>{
        {"3^3^3",19683},{"2^3^2",64},{"-2^2",-4},{"(-2)^2",4},
        {"2^-2",0.25},{"2^-2^3",0.015625},{"-2^-2^3",-0.015625},
        {"2*-3^2",-18},{"(-2)^-3",-0.125},{"0^0",1},{"Empty^2",0}})
        assert(evaluate(item.first).asNumber()==item.second);
    assert(evaluate("VarType(CInt(2)^CInt(3))=vbDouble").truthy());
    assert(evaluate("Null^2").kind==VbaValue::Kind::Null && evaluate("2^Null").kind==VbaValue::Kind::Null);
    for(const auto& item:std::vector<std::pair<const char*,double>>{
        {"10 Xor 8",2},{"10 Eqv 8",-3},{"8 Imp 10",-1},
        {"1 Or 2 Xor 3",0},{"0 Eqv 1 Imp 2",3},{"1.5 Xor 2",0},
        {"0 Imp Null",-1},{"Null Imp -1",-1}})
        assert(evaluate(item.first).asNumber()==item.second);
    for(const auto& item:std::vector<std::pair<const char*,std::optional<bool>>>{
        {"True Xor True",false},{"True Xor False",true},{"False Xor True",true},{"False Xor False",false},
        {"True Eqv True",true},{"True Eqv False",false},{"False Eqv True",false},{"False Eqv False",true},
        {"True Imp True",true},{"True Imp False",false},{"True Imp Null",{}},
        {"False Imp True",true},{"False Imp False",true},{"False Imp Null",true},
        {"Null Imp True",true},{"Null Imp False",{}},{"Null Imp Null",{}},
        {"Null Xor True",{}},{"True Eqv Null",{}},{"1 Imp Null",{}},{"Null Imp 1",{}}}) {
        const auto result=evaluate(item.first);
        if(!item.second) assert(result.kind==VbaValue::Kind::Null);
        else assert(result.kind==VbaValue::Kind::Boolean && result.boolean==*item.second);
    }
    assert(evaluate("VarType(CInt(10) Eqv CInt(8))=vbInteger").truthy());
    assert(evaluate("VarType(CLng(10) Xor CLng(8))=vbLong").truthy());
    assert(evaluate("Abs(Null)").kind==VbaValue::Kind::Null && evaluate("Int(Null)").kind==VbaValue::Kind::Null && evaluate("Fix(Null)").kind==VbaValue::Kind::Null);
    assert(evaluate("Abs(-3.75)").number==3.75 && evaluate("Int(-3.75)").number==-4 && evaluate("Fix(-3.75)").number==-3);
    assert(evaluate("Sgn(-3.75)").number==-1 && evaluate("Sgn(0)").number==0 && evaluate("Sgn(3.75)").number==1);
    assert(evaluate("Abs(Empty)").number==0 && evaluate("Fix(Empty)").number==0 && evaluate("Abs(True)").number==1);
    assert(evaluate("VarType(Abs(CInt(-3)))=vbInteger").truthy());
    assert(evaluate("VarType(Fix(CLng(-3)))=vbLong").truthy());
    assert(evaluate("VarType(Int(CSng(-3.75)))=vbSingle").truthy());
    assert(evaluate("VarType(Sgn(1))=vbInteger").truthy());
    assert(evaluate("Fix(CCur(\"-922337203685477.5808\"))").asText()=="-922337203685477");
    assert(evaluate("VarType(Fix(CCur(1.2)))=vbCurrency").truthy());
    assert(evaluate("CLng(2.5)").asNumber() == 2);
    assert(evaluate("CLng(3.5)").asNumber() == 4);
    assert(evaluate("CLng(-2.5)").asNumber() == -2);
    assert(evaluate("Null & Null").kind == VbaValue::Kind::Null);
    assert(evaluate("Nothing Is Nothing").truthy());
    assert(evaluate("Len(Null)").kind == VbaValue::Kind::Null);
    bool invalidNull = false;
    try { (void)evaluate("CStr(Null)"); } catch (...) { invalidNull = true; }
    assert(invalidNull);
    assert(evaluate(".5+2e1").asNumber()==20.5);
    assert(evaluate("1E-2").asNumber()==0.01);
    assert(evaluate("Len(\"Grüße😀\")").asNumber()==7);
    assert(evaluate("VBA.LenB(\"Grüße😀\")").asNumber()==14 && evaluate("VarType(LenB(\"a\"))=vbLong").truthy());
    assert(evaluate("LenB(Null)").kind==VbaValue::Kind::Null && evaluate("LenB(Empty)").asNumber()==0);
    for(const auto& [source,expected]:std::vector<std::pair<std::string,int>>{{"LenB(1)",3251},{"LenB()",450},{"LenB(1,2)",450}}) {
        int actual=0; try { (void)evaluate(source); } catch(const VbaRuntimeError& error) { actual=error.number; } assert(actual==expected);
    }

    auto scalarProbe=std::make_shared<ScalarProbe>(); scalarProbe->value=VbaValue::string("ö😀");
    const VbaVariables scalarValues{{"probe",VbaValue::reference(scalarProbe)}};
    assert(evaluateVbaExpression("Len(probe)",host,scalarValues).asNumber()==3 && scalarProbe->reads==1);
    scalarProbe->value=VbaValue{};
    assert(evaluateVbaExpression("Len(probe)",host,scalarValues).kind==VbaValue::Kind::Null && scalarProbe->reads==2);
    const VbaVariables unicodeValues{{"embedded",VbaValue::string(std::string("A\0B",3))},{"invalid",VbaValue::string(std::string(1,static_cast<char>(0xff)))}};
    assert(evaluateVbaExpression("Len(embedded)",host,unicodeValues).asNumber()==3);
    assert(evaluateVbaExpression("LenB(embedded)",host,unicodeValues).asNumber()==6);
    assert(evaluateVbaExpression("Mid(embedded,2,1)",host,unicodeValues).asText()==std::string(1,'\0'));
    int invalidUnicode=0; try { (void)evaluateVbaExpression("Len(invalid)",host,unicodeValues); } catch(const VbaRuntimeError& failure) { invalidUnicode=failure.number; } assert(invalidUnicode==13);

    assert(evaluate("Left(\"Grüße\",3)").asText()=="Grü");
    assert(evaluate("Right$(\"Grüße\",2)").asText()=="ße");
    assert(evaluate("Mid(\"A😀Z\",2,2)").asText()=="😀");
    assert(evaluate("Mid$(\"A😀Z\",4)").asText()=="Z");
    assert(evaluate("Right(\"A😀Z\",3)").asText()=="😀Z");
    assert(evaluate("Left$(\"abc\",1.5)").asText()=="ab");
    assert(evaluate("Mid(\"abc\",99)").asText().empty());
    assert(evaluate("Left(Null,2)").kind==VbaValue::Kind::Null);
    assert(evaluate("Right(Null,2)").kind==VbaValue::Kind::Null);
    assert(evaluate("Mid(Null,1)").kind==VbaValue::Kind::Null);
    const auto substringError=[&](const char* expression,int expected) {
        int number=0; try { (void)evaluate(expression); } catch(const VbaRuntimeError& failure) { number=failure.number; } if(number!=expected) std::cerr<<"Substring error: expected="<<expected<<", actual="<<number<<", expression="<<expression<<"\n"; assert(number==expected);
    };
    substringError("CSng(Null)",94); substringError("CSng(CVErr(2015))",13);
    substringError("CSng(3.5e38)",6); substringError("CSng(1e-50)",3251);
    substringError("CSng(\"invalid\")",13); substringError("CSng()",450);
    substringError("CDbl()",450); substringError("CLng(1,2)",450);
    scalarProbe->value=VbaValue::numeric(0.1); const auto beforeSingle=scalarProbe->reads;
    const auto singleValue=evaluateVbaExpression("CSng(probe)",host,scalarValues);
    assert(singleValue.numericType==VbaValue::NumericType::Single && singleValue.number==static_cast<double>(0.1f) && scalarProbe->reads==beforeSingle+1);
    substringError(R"(7 \ 0.5)",11); substringError("7 Mod 0.5",11);
    substringError(R"(-2147483648 \ -1)",6); substringError(R"(CInt(-32768) \ CInt(-1))",6);
    substringError(R"("invalid" Mod 2)",13); substringError(R"(1e100 \ 2)",6);
    scalarProbe->value=VbaValue::numeric(12.6); const auto beforeDivision=scalarProbe->reads;
    assert(evaluateVbaExpression(R"(probe \ 5)",host,scalarValues).number==2 && scalarProbe->reads==beforeDivision+1);
    substringError("(-2)^0.5",5); substringError("1e200^2",6);
    substringError("0^-1",3251); substringError("1e-200^2",3251);
    substringError(R"("invalid"^2)",13);
    scalarProbe->value=VbaValue::numeric(3); const auto beforePower=scalarProbe->reads;
    assert(evaluateVbaExpression("probe^2",host,scalarValues).number==9 && scalarProbe->reads==beforePower+1);
    substringError("1e100 Xor 2",6); substringError(R"("invalid" Eqv 2)",13);
    scalarProbe->value=VbaValue::numeric(10); const auto beforeExclusive=scalarProbe->reads;
    assert(evaluateVbaExpression("probe Xor 8",host,scalarValues).number==2 && scalarProbe->reads==beforeExclusive+1);
    substringError("Abs(CInt(-32768))",6); substringError("Abs(CLng(-2147483648))",6);
    substringError("Sgn(Null)",94); substringError("Fix(CVErr(2015))",13);
    substringError("Abs()",450); substringError("Int(1,2)",450); substringError("Fix()",450); substringError("Sgn()",450);
    scalarProbe->value=VbaValue::numeric(-3.75); const auto beforeFix=scalarProbe->reads;
    assert(evaluateVbaExpression("Fix(probe)",host,scalarValues).number==-3 && scalarProbe->reads==beforeFix+1);
    substringError("CInt(Null)",94); substringError("CInt(CVErr(2015))",13);
    substringError("CInt(\"invalid\")",13); substringError("CInt(32767.5)",6);
    substringError("CInt(-32768.6)",6); substringError("CInt()",450); substringError("CInt(1,2)",450);
    assert(evaluate("CInt(-32768.5)").asNumber()==-32768);
    scalarProbe->value=VbaValue::numeric(2.5); const auto beforeInteger=scalarProbe->reads;
    assert(evaluateVbaExpression("CInt(probe)",host,scalarValues).numericType==VbaValue::NumericType::Integer && scalarProbe->reads==beforeInteger+1);
    substringError("CBool(Null)",94); substringError("CBool(CVErr(2015))",13);
    substringError("CBool(\"\")",13); substringError("CBool(\"invalid\")",13);
    substringError("CBool()",450); substringError("CBool(1,2)",450);
    substringError("CBool(Nothing)",91);
    scalarProbe->value=VbaValue::string("False"); const auto beforeBoolean=scalarProbe->reads;
    const auto convertedBoolean=evaluateVbaExpression("CBool(probe)",host,scalarValues);
    assert(convertedBoolean.kind==VbaValue::Kind::Boolean && !convertedBoolean.boolean && scalarProbe->reads==beforeBoolean+1);
    substringError("CLng(Null)",94); substringError("CDbl(Null)",94);
    substringError("Left$(Null,2)",94); substringError("Mid$(Null,1)",94);
    substringError("Left(\"abc\",Null)",94); substringError("Left(\"abc\",-1)",5);
    substringError("Mid(\"abc\",0)",5); substringError("Left(\"abc\",1e100)",6);
    substringError("Left(\"abc\")",450);
    substringError("1e309",6); substringError("1e+",5);
    substringError("Left(\"😀\",1)",3251); substringError("Mid(\"😀\",2,1)",3251);
    const auto unicodeProcedure=executeVbaProcedure(R"VBA(Sub UnicodeStrings()
Me!Name=Mid$("A😀Z",2,2) & Left$("ö1",1)
On Error Resume Next
Me!Nachname=Left$(Null,2)
Me!UnicodeError=CStr(Err.Number)
End Sub
)VBA","UnicodeStrings",host);
    assert(unicodeProcedure.status==VbaExecutionStatus::Completed && host.fields["Name"]=="😀ö" && host.fields["UnicodeError"]=="94");
    assert(evaluate("Trim(\"  ö  \")").asText()=="ö");
    assert(evaluate("LTrim$(\"  ö  \")").asText()=="ö  ");
    assert(evaluate("RTrim(\"  ö  \")").asText()=="  ö");
    assert(evaluate("Trim$(\"   \")").asText().empty());
    assert(evaluate("Trim(Null)").kind==VbaValue::Kind::Null);
    assert(evaluate("LTrim(Null)").kind==VbaValue::Kind::Null && evaluate("RTrim(Null)").kind==VbaValue::Kind::Null);
    substringError("Trim$(Null)",94); substringError("LTrim$(Null)",94); substringError("RTrim$(Null)",94);
    substringError("Trim()",450);
    assert(evaluate("RGB(255,0,0)").asNumber()==255);
    assert(evaluate("RGB(0,255,0)").asNumber()==65280);
    assert(evaluate("RGB(0,0,255)").asNumber()==16711680);
    assert(evaluate("RGB(255,255,255)").asNumber()==16777215);
    assert(evaluate("RGB(300,127.5,0.5)").asNumber()==33023);
    assert(evaluate("RGB(1e100,0,0)").asNumber()==255);
    assert(evaluate("RGB(Empty,0,0)").asNumber()==0);
    substringError("RGB(Null,0,0)",94); substringError("RGB(-1,0,0)",5);
    substringError("RGB(\"bad\",0,0)",13); substringError("RGB(CVErr(5),0,0)",13);
    substringError("RGB(1,2)",450); substringError("RGB(1,2,3,4)",450);
    scalarProbe->value=VbaValue::numeric(12); const auto beforeColor=scalarProbe->reads;
    assert(evaluateVbaExpression("RGB(probe,probe,probe)",host,scalarValues).asNumber()==789516 && scalarProbe->reads==beforeColor+3);
    const int legacyColors[]={0x000000,0x800000,0x008000,0x808000,0x000080,0x800080,0x008080,0xc0c0c0,0x808080,0xff0000,0x00ff00,0xffff00,0x0000ff,0xff00ff,0x00ffff,0xffffff};
    for(int index=0;index<16;++index) assert(evaluate(("QBColor("+std::to_string(index)+")").c_str()).asNumber()==legacyColors[index]);
    assert(evaluate("QBColor(7.5)").asNumber()==0x808080 && evaluate("QBColor(Empty)").asNumber()==0);
    assert(evaluate("QBColor(\"12\")").asNumber()==255);
    assert(evaluate("acTable").asNumber()==0 && evaluate("acQuery").asNumber()==1 && evaluate("acModule").asNumber()==5);
    substringError("QBColor(Null)",94); substringError("QBColor(\"bad\")",13);
    substringError("QBColor(-1)",5); substringError("QBColor(16)",5);
    substringError("QBColor()",450); substringError("QBColor(1,2)",450);
    const auto beforeLegacy=scalarProbe->reads;
    assert(evaluateVbaExpression("QBColor(probe)",host,scalarValues).asNumber()==255 && scalarProbe->reads==beforeLegacy+1);
    const VbaVariables whitespaceValues{{"text",VbaValue::string(" \t ö\r\n ")}};
    assert(evaluateVbaExpression("Trim(text)",host,whitespaceValues).asText()=="\t ö\r\n");
    scalarProbe->value=VbaValue{};
    const int beforeTrim=scalarProbe->reads;
    assert(evaluateVbaExpression("Trim(probe)",host,scalarValues).kind==VbaValue::Kind::Null && scalarProbe->reads==beforeTrim+1);
    const auto trimProcedure=executeVbaProcedure(R"VBA(Sub TrimStrings()
Me!Name=LTrim$("  ö  ") & ":" & RTrim$("  x  ")
On Error Resume Next
Me!Nachname=Trim$(Null)
Me!TrimError=CStr(Err.Number)
End Sub
)VBA","TrimStrings",host);
    assert(trimProcedure.status==VbaExecutionStatus::Completed && host.fields["Name"]=="ö  :  x" && host.fields["TrimError"]=="94");
    assert(evaluate("InStr(\"A😀Z\",\"Z\")").asNumber()==4);
    assert(evaluate("InStr(2,\"A😀Z\",\"😀\",vbBinaryCompare)").asNumber()==2);
    assert(evaluate("InStr(3,\"A😀Z\",\"Z\")").asNumber()==4);
    assert(evaluate("InStr(,\"abc\",\"b\")").asNumber()==2);
    assert(evaluate("InStr(1,\"ABC\",\"b\",vbTextCompare)").asNumber()==2);
    assert(evaluate("InStr(\"abc\",\"Z\")").asNumber()==0);
    assert(evaluate("InStr(\"\",\"\")").asNumber()==0);
    assert(evaluate("InStr(8,\"abc\",\"\")").asNumber()==8);
    assert(evaluate("InStr(Null,\"a\")").kind==VbaValue::Kind::Null);
    assert(evaluate("StrComp(\"A\",\"a\",vbBinaryCompare)").asNumber()==-1);
    assert(evaluate("StrComp(\"A\",\"a\",vbTextCompare)").asNumber()==0);
    assert(evaluate("StrComp(\"😀\",\"\",0)").asNumber()==-1 && evaluate("\"😀\"<\"\"").truthy());
    assert(evaluate("StrComp(Null,\"a\")").kind==VbaValue::Kind::Null);
    substringError("InStr(0,\"a\",\"a\")",5); substringError("InStr(Null,\"a\",\"a\")",94);
    substringError("InStr(1,\"a\",\"a\",1031)",3251);
    substringError("StrComp(\"a\",\"b\",Null)",94); substringError("StrComp(\"a\",\"b\",3)",5);
    substringError("InStr(\"abc\",\"b\",1)",13);
    const auto textComparison=executeVbaProcedure(R"VBA(Option Compare Text
Sub CompareStrings()
Me!Name=CStr(InStr("ABc","b"))
Me!Nachname=CStr(StrComp("A","a",vbUseCompareOption))
End Sub
)VBA","CompareStrings",host);
    assert(textComparison.status==VbaExecutionStatus::Completed && host.fields["Name"]=="2" && host.fields["Nachname"]=="0");
    {
        namespace fs=std::filesystem;
        char directoryPattern[]="/tmp/linaccess-file-facts-XXXXXX";
        const auto* directoryName=::mkdtemp(directoryPattern); assert(directoryName);
        const fs::path directory(directoryName), file=directory/fs::path("native-ä.dat");
        const auto child=directory/"created-ä";
        host.fields["NativeDirectory"]=child.string();
        const auto directoryRun=[&](const std::string& body) {
            return executeVbaProcedure("Sub NativeDirectoryOperations()\n"+body+"\nEnd Sub", "NativeDirectoryOperations",host);
        };
        assert(directoryRun("MkDir Me!NativeDirectory").status==VbaExecutionStatus::Completed && fs::is_directory(child));
        assert(directoryRun("MkDir Me!NativeDirectory").errorNumber==75);
        const auto savedDirectory=fs::current_path();
        assert(directoryRun("ChDir Me!NativeDirectory").status==VbaExecutionStatus::Completed);
        assert(fs::equivalent(fs::current_path(),child));
        assert(evaluate("CurDir()").asText()==child.string());
        assert(evaluate("CurDir").asText()==child.string());
        assert(directoryRun("Me!CurrentNativeDirectory=CurDir").status==VbaExecutionStatus::Completed);
        assert(host.fields["CurrentNativeDirectory"]==child.string());
        const auto shadow=executeVbaProcedure("Function CurDir() As String\nCurDir=\"module function\"\nEnd Function\nSub ReadNativeDirectory()\nMe!CurrentNativeDirectory=CurDir\nEnd Sub", "ReadNativeDirectory",host);
        assert(shadow.status==VbaExecutionStatus::Completed && host.fields["CurrentNativeDirectory"]=="module function");

        assert(evaluate("VBA.CurDir$(Drive:=\"\")").asText()==child.string());
        substringError("CurDir(Null)",94);
        substringError("CurDir(\"C\")",3251);
        substringError("CurDir(1,2)",450);
        VbaVariables directoryLocals{{"curdir",VbaValue::string("local directory variable")}};
        assert(evaluateVbaExpression("CurDir",host,directoryLocals).asText()=="local directory variable");

        assert(directoryRun("RmDir \".\"").errorNumber==75 && fs::exists(child));
        assert(directoryRun("ChDir \"missing-child\"").errorNumber==76);
        assert(fs::equivalent(fs::current_path(),child));
        assert(directoryRun("ChDir Null").errorNumber==94);
        assert(directoryRun("ChDir \"\"").errorNumber==76);
        fs::current_path(savedDirectory);

        { std::ofstream output(child/"retained.txt"); output<<"synthetic"; }
        assert(directoryRun("RmDir Me!NativeDirectory").errorNumber==75 && fs::exists(child/"retained.txt"));
        host.fields["NativeDirectory"]=(child/"retained.txt").string();
        assert(directoryRun("RmDir Me!NativeDirectory").errorNumber==75 && fs::exists(child/"retained.txt"));
        host.fields["NativeDirectory"]=child.string();
        fs::remove(child/"retained.txt");
        assert(directoryRun("RmDir Me!NativeDirectory").status==VbaExecutionStatus::Completed && !fs::exists(child));
        assert(directoryRun("RmDir Me!NativeDirectory").errorNumber==76);
        assert(directoryRun("MkDir Null").errorNumber==94);
        assert(directoryRun("MkDir \"\"").errorNumber==52);

        { std::ofstream output(file,std::ios::binary); output.write("A\0B",3); assert(output.good()); }
        const auto copy=directory/"copied-ä.dat";
        host.fields["CopySource"]=file.string(); host.fields["CopyDestination"]=copy.string();
        assert(directoryRun("FileCopy Me!CopySource, Me!CopyDestination").status==VbaExecutionStatus::Completed);
        const auto copiedBytes=[&] { std::ifstream input(copy,std::ios::binary); return std::string(std::istreambuf_iterator<char>(input),{}); };
        assert(copiedBytes()==std::string("A\0B",3));
        { std::ofstream output(copy,std::ios::binary); output << "old destination content"; }
        assert(directoryRun("FileCopy Me!CopySource, Me!CopyDestination").status==VbaExecutionStatus::Completed && copiedBytes()==std::string("A\0B",3));
        host.fields["CopySource"]=(directory/"absent.dat").string();
        assert(directoryRun("FileCopy Me!CopySource, Me!CopyDestination").errorNumber==53 && copiedBytes()==std::string("A\0B",3));
        host.fields["CopySource"]=file.string(); host.fields["CopyDestination"]=file.string();
        assert(directoryRun("FileCopy Me!CopySource, Me!CopyDestination").errorNumber==3251 && fs::file_size(file)==3);
        const auto renamed=directory/"renamed-ä As file.dat";
        host.fields["RenameSource"]=copy.string(); host.fields["RenameTarget"]=renamed.string();
        assert(directoryRun("Name Me!RenameSource As Me!RenameTarget").status==VbaExecutionStatus::Completed);
        assert(!fs::exists(copy) && fs::file_size(renamed)==3);
        assert(directoryRun("Name CStr(\""+renamed.string()+"\") As Me!RenameSource").status==VbaExecutionStatus::Completed);
        assert(fs::exists(copy) && !fs::exists(renamed));
        host.fields["RenameTarget"]=file.string();
        assert(directoryRun("Name Me!RenameSource As Me!RenameTarget").errorNumber==58 && fs::exists(copy) && fs::file_size(file)==3);
        assert(directoryRun("Dim Name As String\nDim FileCopy As String\nDim MkDir As String\nName = \"kept As text\"\nFileCopy = \"kept copy\"\nMkDir = \"kept directory\"\nMe!NameStatementVariable=Name & \"|\" & FileCopy & \"|\" & MkDir").status==VbaExecutionStatus::Completed && host.fields["NameStatementVariable"]=="kept As text|kept copy|kept directory");
        const auto folder=directory/"folder-source",movedFolder=directory/"folder-target";
        assert(fs::create_directory(folder)); { std::ofstream output(folder/"kept.dat"); output << "kept"; }
        host.fields["RenameSource"]=folder.string(); host.fields["RenameTarget"]=movedFolder.string();
        assert(directoryRun("Name Me!RenameSource As Me!RenameTarget").status==VbaExecutionStatus::Completed);
        assert(!fs::exists(folder) && fs::file_size(movedFolder/"kept.dat")==4);
        host.fields["KillTarget"]=copy.string();
        assert(directoryRun("If False Then\nKill Me!KillTarget\nEnd If").status==VbaExecutionStatus::Completed && fs::exists(copy));
        assert(directoryRun("Kill CStr(Me!KillTarget)").status==VbaExecutionStatus::Completed && !fs::exists(copy) && fs::file_size(file)==3);
        assert(directoryRun("Kill Me!KillTarget").errorNumber==53);
        host.fields["KillTarget"]=(directory/"missing-parent"/"missing.dat").string();
        assert(directoryRun("Kill Me!KillTarget").errorNumber==76);
        host.fields["KillTarget"]=directory.string();
        assert(directoryRun("Kill Me!KillTarget").errorNumber==75 && fs::exists(directory));
        assert(directoryRun("Kill Null").errorNumber==94 && directoryRun("Kill \"*.dat\"").errorNumber==3251);
        assert(directoryRun("Dim Kill As String\nKill = \"kept\"\nMe!KillVariable=Kill").status==VbaExecutionStatus::Completed && host.fields["KillVariable"]=="kept");
        host.fields["FsoFile"]=file.string(); host.fields["FsoFolder"]=directory.string();
        const auto fsoProbe=executeVbaProcedure(R"VBA(Function MakeFs() As Scripting.FileSystemObject
Set MakeFs=New Scripting.FileSystemObject
End Function
Function TypedFileExists(ByVal fs As Scripting.FileSystemObject,ByVal spec As String) As Boolean
TypedFileExists=fs.FileExists(spec)
End Function
Sub ProbeFs()
Dim fs As Scripting.FileSystemObject
Set fs=MakeFs()
Me!FsoTyped=CStr(TypeOf fs Is Scripting.FileSystemObject)
Me!FsoIdentity=TypeName(fs) & ":" & CStr(VarType(fs))
Me!FsoHasFile=CStr(fs.FileExists(FileSpec:=Me!FsoFile))
Me!FsoTypedCall=CStr(TypedFileExists(fs,Me!FsoFile))
Me!FsoHasFolder=CStr(fs.FolderExists(FolderSpec:=Me!FsoFolder))
Me!FsoWrongKinds=CStr(fs.FileExists(Me!FsoFolder)) & ":" & CStr(fs.FolderExists(Me!FsoFile))
Me!FsoMissing=CStr(fs.FileExists(fs.BuildPath(Name:="missing.dat",Path:=Me!FsoFolder)))
Me!FsoAbsolute=fs.GetAbsolutePathName(Path:=".")
Me!FsoLeaf=fs.GetFileName(Path:="C:\docs\alpha-ä.tar.gz")
Me!FsoBase=fs.GetBaseName(Path:="C:\docs\alpha-ä.tar.gz")
Me!FsoExtension=fs.GetExtensionName(Path:="/missing/alpha-ä.tar.gz")
Me!FsoParent=fs.GetParentFolderName(Path:="C:\docs\alpha-ä.tar.gz")
Me!FsoWildcardLeaf=fs.GetFileName("C:\docs\*.csv")
Me!FsoDriveOnly=fs.GetFileName("C:")
Dim nativeFile As Scripting.File
Set nativeFile=fs.GetFile(FilePath:=Me!FsoFile)
Dim folder As Scripting.Folder, files As Scripting.Files, entry As Variant, total As Long
Set folder=fs.GetFolder(FolderPath:=Me!FsoFolder & "/")
Set files=folder.Files
Me!FsoFolderMetadata=TypeName(folder) & ":" & CStr(VarType(folder)) & ":" & CStr(TypeOf folder Is Scripting.Folder)
Me!FsoFolderPath=CStr(folder)
Me!FsoFilesCount=CStr(files.Count)
Me!FsoSubFoldersCount=CStr(folder.SubFolders.Count)
Me!FsoFileParent=nativeFile.ParentFolder.Path
For Each entry In folder.SubFolders
 Me!FsoSubFolderName=entry.Name
Next entry
For Each entry In files
 total=total+entry.Size
Next entry
Me!FsoFilesTotal=CStr(total)
Dim created As Scripting.Folder, subfolders As Scripting.Folders
Set subfolders=folder.SubFolders
Set created=fs.CreateFolder(Path:=folder.Path & "/created-ä")
Me!FsoCreated=created.Name & ":" & CStr(subfolders.Count)
On Error Resume Next
Set created=fs.CreateFolder(folder.Path & "/created-ä")
Me!FsoDuplicateFolder=CStr(Err.Number) & ":" & created.Name
On Error GoTo 0
Dim copied As Scripting.File
fs.CopyFile Source:=Me!FsoFile, Destination:=created.Path & "/copy.dat", OverWriteFiles:=False
Set copied=fs.GetFile(created.Path & "/copy.dat")
Me!FsoCopiedSize=CStr(copied.Size)
On Error Resume Next
Err.Clear
fs.CopyFile Me!FsoFile, copied.Path, False
Me!FsoCopyExisting=CStr(Err.Number)
On Error GoTo 0
copied.Copy Destination:=created.Path & "/copy2.dat", OverWriteFiles:=False
copied.Name="renamed.dat"
copied.Move Destination:=created.Path & "/moved.dat"
Me!FsoMoved=copied.Name & ":" & CStr(copied.Size)
fs.MoveFile Source:=copied.Path, Destination:=created.Path & "/final.dat"
fs.DeleteFile FileSpec:=created.Path & "/copy2.dat", Force:=False
Set copied=fs.GetFile(created.Path & "/final.dat")
copied.Delete Force:=False
Me!FsoOpsClean=CStr(created.Files.Count)
Me!FsoFileDefault=CStr(nativeFile)
Me!FsoFileTyped=CStr(TypeOf nativeFile Is Scripting.File)
Me!FsoFileObject=TypeName(nativeFile) & ":" & CStr(VarType(nativeFile))
Me!FsoFileMetadata=nativeFile.Path & "|" & nativeFile.Name & "|" & CStr(nativeFile.Size) & ":" & CStr(VarType(nativeFile.Size))

On Error Resume Next
Set fs=New Scripting.Dictionary
Me!FsoWrongClass=CStr(Err.Number) & ":" & TypeName(fs)
On Error GoTo 0
Set fs=Nothing
Set fs=CreateObject("Scripting.FileSystemObject")
Me!FsoLate=CStr(fs.FileExists(Me!FsoFile))
Set fs=Nothing
End Sub)VBA","ProbeFs",host);
        assert(fsoProbe.status==VbaExecutionStatus::Completed);
        assert(host.fields["FsoTyped"]=="True" && host.fields["FsoTypedCall"]=="True" && host.fields["FsoWrongClass"]=="13:FileSystemObject" && host.fields["FsoLate"]=="True");
        assert(host.fields["FsoIdentity"]=="FileSystemObject:9" && host.fields["FsoHasFile"]=="True" && host.fields["FsoHasFolder"]=="True");
        assert(host.fields["FsoLeaf"]=="alpha-ä.tar.gz" && host.fields["FsoBase"]=="alpha-ä.tar" && host.fields["FsoExtension"]=="gz");
        assert(host.fields["FsoParent"]=="C:\\docs" && host.fields["FsoWildcardLeaf"]=="*.csv" && host.fields["FsoDriveOnly"]=="");
        assert(host.fields["FsoFolderMetadata"]=="Folder:8:True" && host.fields["FsoFolderPath"]==directory.string());
        assert(host.fields["FsoFilesCount"]=="1" && host.fields["FsoFilesTotal"]=="3");
        assert(host.fields["FsoSubFoldersCount"]=="1" && host.fields["FsoSubFolderName"]=="folder-target" && host.fields["FsoFileParent"]==directory.string());
        assert(host.fields["FsoCreated"]=="created-ä:2" && host.fields["FsoDuplicateFolder"]=="58:created-ä" && fs::is_directory(directory/"created-ä"));
        assert(host.fields["FsoCopiedSize"]=="3" && host.fields["FsoCopyExisting"]=="58" && host.fields["FsoMoved"]=="moved.dat:3" && host.fields["FsoOpsClean"]=="0");
        assert(fs::is_empty(directory/"created-ä") && fs::file_size(file)==3);
        assert(host.fields["FsoFileObject"]=="File:8" && host.fields["FsoFileTyped"]=="True" && host.fields["FsoFileDefault"]==file.string());
        assert(host.fields["FsoFileMetadata"]==file.string()+"|"+file.filename().string()+"|3:3");
        assert(host.fields["FsoWrongKinds"]=="False:False" && host.fields["FsoMissing"]=="False" && host.fields["FsoAbsolute"]==fs::current_path().string());
        const auto textFile=directory/"read-lines.txt";
        { std::ofstream output(textFile,std::ios::binary); output<<"one\r\ntwo\nend"; }
        host.fields["FsoTextFile"]=textFile.string();
        const auto textProbe=executeVbaProcedure(R"VBA(Sub ReadNativeText()
Dim fs As Scripting.FileSystemObject, stream As Scripting.TextStream, file As Scripting.File
Set fs=New Scripting.FileSystemObject
Set stream=fs.OpenTextFile(FileName:=Me!FsoTextFile, IOMode:=1, Create:=False, Format:=0)
Me!FsoStreamType=TypeName(stream) & ":" & CStr(TypeOf stream Is Scripting.TextStream)
Me!FsoTextLine=stream.ReadLine()
Me!FsoTextPosition=CStr(stream.Line) & ":" & CStr(stream.Column)
Me!FsoTextPart=stream.Read(Characters:=3)
Me!FsoTextLineEnd=CStr(stream.AtEndOfLine)
stream.SkipLine
Me!FsoTextRest=stream.ReadAll()
Me!FsoTextEnd=CStr(stream.AtEndOfStream)
On Error Resume Next
Err.Clear
fs.DeleteFile Me!FsoTextFile
Me!FsoTextHeld=CStr(Err.Number)
On Error GoTo 0
stream.Close
fs.CopyFile Me!FsoFile, Me!FsoTextFile
Set file=fs.GetFile(Me!FsoTextFile)
Set stream=file.OpenAsTextStream(IOMode:=1, Format:=0)
Me!FsoTextBinary=stream.Read(3)
stream.Close
Set stream=file.OpenAsTextStream()
stream.Skip Characters:=1
Me!FsoTextSkipped=stream.Read(2)
stream.Close
End Sub)VBA","ReadNativeText",host);
        assert(textProbe.status==VbaExecutionStatus::Completed);
        assert(host.fields["FsoStreamType"]=="TextStream:True" && host.fields["FsoTextLine"]=="one" && host.fields["FsoTextPosition"]=="2:1");
        assert(host.fields["FsoTextPart"]=="two" && host.fields["FsoTextRest"]=="end" && host.fields["FsoTextLineEnd"]=="True" && host.fields["FsoTextEnd"]=="True" && host.fields["FsoTextHeld"]=="3251");
        assert(host.fields["FsoTextBinary"]==std::string("A\0B",3) && host.fields["FsoTextSkipped"]==std::string("\0B",2));
        const auto writeFile=directory/"write-lines.txt";
        host.fields["FsoWriteFile"]=writeFile.string();
        const auto writeProbe=executeVbaProcedure(R"VBA(Sub WriteNativeText()
Dim fs As Scripting.FileSystemObject, stream As Scripting.TextStream, other As Scripting.TextStream
Set fs=New Scripting.FileSystemObject
Set stream=fs.CreateTextFile(FileName:=Me!FsoWriteFile, Overwrite:=False, Unicode:=False)
stream.Write Text:="A"
stream.WriteLine Text:="B"
stream.WriteBlankLines Lines:=1
On Error Resume Next
Err.Clear
stream.Write "unexpected ä"
Me!FsoWriteEncoding=CStr(Err.Number)
Err.Clear
Set other=fs.OpenTextFile(Me!FsoWriteFile,1)
Me!FsoWriteHeld=CStr(Err.Number)
On Error GoTo 0
stream.Write "tail"
stream.Close
Set stream=fs.OpenTextFile(FileName:=Me!FsoWriteFile,IOMode:=8,Create:=False,Format:=0)
stream.WriteLine "+end"
stream.Close
Set stream=fs.OpenTextFile(Me!FsoWriteFile)
On Error Resume Next
Err.Clear
Set other=fs.OpenTextFile(Me!FsoWriteFile,2)
Me!FsoReadHeldWrite=CStr(Err.Number)
On Error GoTo 0
Me!FsoWritten=stream.ReadAll()
stream.Close
On Error Resume Next
Err.Clear
Set stream=fs.CreateTextFile(Me!FsoWriteFile,False)
Me!FsoTextNoOverwrite=CStr(Err.Number)
On Error GoTo 0
Dim folder As Scripting.Folder
Set folder=fs.GetFolder(Me!FsoFolder)
Set stream=folder.CreateTextFile(FileName:="via-folder.txt",Overwrite:=False,Unicode:=True)
stream.WriteLine "ä"
stream.Close
Set stream=fs.OpenTextFile(folder.Path & "/via-folder.txt",1,False,-1)
Me!FsoFolderCreatedText=stream.ReadLine()
stream.Close
End Sub)VBA","WriteNativeText",host);
        assert(writeProbe.status==VbaExecutionStatus::Completed);
        const std::string written="AB\r\n\r\ntail+end\r\n";
        assert(host.fields["FsoFolderCreatedText"]=="ä" && fs::is_regular_file(directory/"via-folder.txt"));
        assert(host.fields["FsoWritten"]==written && host.fields["FsoWriteEncoding"]=="3251" && host.fields["FsoWriteHeld"]=="3251" && host.fields["FsoReadHeldWrite"]=="3251" && host.fields["FsoTextNoOverwrite"]=="58");
        { std::ifstream input(writeFile,std::ios::binary); assert(std::string(std::istreambuf_iterator<char>(input),{})==written); }
        const auto unicodeFile=directory/"unicode-lines.txt";
        host.fields["FsoUnicodeFile"]=unicodeFile.string();
        const auto unicodeProbe=executeVbaProcedure(R"VBA(Sub UnicodeNativeText()
Dim fs As Scripting.FileSystemObject, stream As Scripting.TextStream
Set fs=New Scripting.FileSystemObject
Set stream=fs.CreateTextFile(FileName:=Me!FsoUnicodeFile,Overwrite:=False,Unicode:=True)
stream.WriteLine "ä😀X"
stream.Write "Z"
stream.Close
Set stream=fs.OpenTextFile(FileName:=Me!FsoUnicodeFile,IOMode:=8,Create:=False,Format:=-1)
stream.Write "ß"
stream.Close
Set stream=fs.OpenTextFile(FileName:=Me!FsoUnicodeFile,IOMode:=1,Create:=False,Format:=-1)
Me!FsoUnicodeFirst=stream.Read(1)
On Error Resume Next
Err.Clear
Me!FsoUnicodeHalf=stream.Read(1)
Me!FsoUnicodeSplit=CStr(Err.Number) & ":" & CStr(stream.Column)
On Error GoTo 0
Me!FsoUnicodePair=stream.Read(2)
Me!FsoUnicodeLine=stream.ReadLine()
Me!FsoUnicodeRest=stream.ReadAll()
Me!FsoUnicodeEnd=CStr(stream.AtEndOfStream)
stream.Close
End Sub)VBA","UnicodeNativeText",host);
        assert(unicodeProbe.status==VbaExecutionStatus::Completed);
        assert(host.fields["FsoUnicodeFirst"]=="ä" && host.fields["FsoUnicodeSplit"]=="3251:2" && host.fields["FsoUnicodePair"]=="😀");
        assert(host.fields["FsoUnicodeLine"]=="X" && host.fields["FsoUnicodeRest"]=="Zß" && host.fields["FsoUnicodeEnd"]=="True");
        { std::ifstream input(unicodeFile,std::ios::binary); const std::vector<unsigned char> bytes(std::istreambuf_iterator<char>(input),{});
          const std::vector<unsigned char> expected{255,254,228,0,61,216,0,222,88,0,13,0,10,0,90,0,223,0}; assert(bytes==expected); }
        VbaVariables paths{{"filepath",VbaValue::string(file.string())}};
        const auto fileEvaluate=[&](const std::string& expression) { return evaluateVbaExpression(expression,host,paths); };
        const auto folderLink=directory/"folder-link";
        fs::create_directory_symlink(directory/"created-ä",folderLink);
        paths["linkpath"]=VbaValue::string(folderLink.string()+"/");
        for(const auto& expression:{"CreateObject(\"Scripting.FileSystemObject\").GetFolder(linkPath)","CreateObject(\"Scripting.FileSystemObject\").CopyFile(filePath,linkPath)"}) {
            int number=0; try { (void)fileEvaluate(expression); } catch(const VbaRuntimeError& error) { number=error.number; }
            assert(number==3251 && fs::is_empty(directory/"created-ä"));
        }
        fs::remove(folderLink);
        const auto length=fileEvaluate("VBA.FileLen(PathName:=filePath)"); assert(length.asNumber()==3 && length.numericType==VbaValue::NumericType::Long);
        paths["filepath"]=VbaValue::string(fs::relative(file).string()); assert(fileEvaluate("FileLen(filePath)").asNumber()==3);
        paths["filepath"]=VbaValue::string(file.string());
        const auto fixed=std::chrono::sys_days(std::chrono::year(2024)/2/29)+std::chrono::hours(12)+std::chrono::minutes(34)+std::chrono::seconds(56);
        fs::last_write_time(file,fs::file_time_type::clock::from_sys(fixed));
        const char* previousZone=std::getenv("TZ"); const bool hadZone=previousZone!=nullptr; const std::string savedZone=previousZone ? previousZone : "";
        assert(::setenv("TZ","UTC0",1)==0); ::tzset();
        const auto datedFile=fileEvaluate("CreateObject(\"Scripting.FileSystemObject\").GetFile(filePath)");
        const auto modified=fileEvaluate("VBA.FileDateTime(PathName:=filePath)");
        const auto fileDate=datedFile.object->member("DateLastModified",{});
        assert(fileDate.kind==VbaValue::Kind::Date && fileDate.asText()=="2024-02-29 12:34:56");
        assert(modified.kind==VbaValue::Kind::Date && modified.asText()=="2024-02-29 12:34:56");
        assert(::setenv("TZ","EST5",1)==0); ::tzset();
        assert(fileEvaluate("FileDateTime(filePath)").asText()=="2024-02-29 07:34:56");
        assert(datedFile.object->member("DateLastModified",{}).asText()=="2024-02-29 07:34:56");
        if(hadZone) assert(::setenv("TZ",savedZone.c_str(),1)==0); else assert(::unsetenv("TZ")==0); ::tzset();
        const auto fileError=[&](const std::string& value,int expected) {
            paths["filepath"]=VbaValue::string(value); int number=0;
            try { (void)fileEvaluate("FileLen(filePath)"); } catch(const VbaRuntimeError& error) { number=error.number; }
            assert(number==expected);
        };
        fileError((directory/"missing.dat").string(),53); fileError((directory/"missing"/"file.dat").string(),76);
        fileError(directory.string(),75); fileError(std::string("bad\0path",8),52); fileError("C:\\native.dat",3251);
        paths["filepath"]=VbaValue::string(file.string());
        const auto liveFile=fileEvaluate("CreateObject(\"Scripting.FileSystemObject\").GetFile(filePath)");
        fs::resize_file(file,std::uintmax_t(1)<<31); fileError(file.string(),6);
        const auto largeSize=liveFile.object->member("Size",{});
        assert(largeSize.asNumber()==2147483648.0 && largeSize.numericType==VbaValue::NumericType::Double);
        substringError("FileLen(Null)",94); substringError("FileDateTime(Null)",94); substringError("FileLen()",450);
        fs::remove_all(directory);
    }
    {
        const auto aligned=executeVbaProcedure(R"VBA(Sub AlignStrings()
Dim text As String, value As Variant, errorNumber As Long
text=Space$(6)
LSet text="A😀Z"
If text<>"A😀Z  " Or Len(text)<>6 Then Err.Raise 5
RSet text="😀"
If text<>"    😀" Then Err.Raise 5
value="012"
RSet value="abcdef"
If value<>"abc" Then Err.Raise 5
AlignByReference text
If text<>"ok    " Then Err.Raise 5
text="_"
On Error Resume Next
LSet text="😀"
errorNumber=Err.Number
On Error GoTo 0
If errorNumber<>3251 Or text<>"_" Then Err.Raise 5
text=""
LSet text="no expansion"
If text<>"" Then Err.Raise 5
Me!AlignedStrings="checked"
End Sub
Sub AlignByReference(ByRef target As String)
LSet target="ok"
End Sub
)VBA","AlignStrings",host);
        assert(aligned.status==VbaExecutionStatus::Completed && host.fields["AlignedStrings"]=="checked");
    }
    {
        const auto mutated=executeVbaProcedure(R"VBA(Sub MutateString()
Dim text As String, errorNumber As Long
text="The dog jumps"
Mid(text,5,3)="fox"
If text<>"The fox jumps" Then Err.Raise 5
Mid(text,5)="cow jumped over"
If text<>"The cow jumpe" Then Err.Raise 5
text="A😀Z"
Mid(text,2,2)="ÖX"
If text<>"AÖXZ" Or Len(text)<>4 Then Err.Raise 5
text="A😀Z"
On Error Resume Next
Mid(text,3,1)="X"
errorNumber=Err.Number
On Error GoTo 0
If errorNumber<>3251 Or text<>"A😀Z" Then Err.Raise 5
text="abc"
MutateByReference text
If text<>"aXc" Then Err.Raise 5
Mid(text,1,0)="unchanged"
If text<>"aXc" Then Err.Raise 5
NamedMidArray
Me!MidMutation="checked"
End Sub
Sub MutateByReference(ByRef target As String)
Mid(target,2,1)="X"
End Sub
Sub NamedMidArray()
Dim Mid(0 To 1,0 To 1) As String
Mid(0,0)="array survives"
If Mid(0,0)<>"array survives" Then Err.Raise 5
End Sub
)VBA","MutateString",host);
        assert(mutated.status==VbaExecutionStatus::Completed && host.fields["MidMutation"]=="checked");
    }
    assert(evaluate("IsArray(Array(1,2))").truthy());
    assert(evaluate("VarType(Array())").asNumber()==8204);
    assert(evaluate("TypeName(Array())").asText()=="Variant()");
    assert(evaluate("LBound(Array())").asNumber()==0 && evaluate("UBound(Array())").asNumber()==-1);
    assert(evaluate("Array(10,20)(1)").asNumber()==20);
    substringError("Array(1)(2)",9); substringError("UBound(Array(1),2)",9);
    substringError("CStr(Array(1))",13);
    substringError("Array(1)(Null)",94); substringError("LBound(7)",13);
    assert(evaluate("Array(Array(5,8))(0)(1)").asNumber()==8);
    assert(!evaluate("IsObject(Array(1))").truthy());
    {
        const auto previousBracketHost=host.bracketHost; host.bracketHost=true;
        host.bracketFields["matrix"]=makeVbaArrayWithBounds({VbaValue::numeric(10),VbaValue::numeric(20),VbaValue::numeric(30),VbaValue::numeric(40)},{{1,2},{-1,0}});
        const auto matrix=executeVbaProcedure(R"VBA(Option Base 1
Sub UseMatrix()
Dim a As Variant
Dim b As Variant
Dim item As Variant
Dim output As String
a=[matrix]
b=a
a(2,-1)=99
Me!MatrixBounds=CStr(LBound(a,1)) & ":" & CStr(UBound(a,1)) & ":" & CStr(LBound(a,2)) & ":" & CStr(UBound(a,2))
Me!MatrixCopies=CStr(a(2,-1)) & ":" & CStr(b(2,-1))
For Each item In a
output=output & CStr(item) & "|"
Next
Me!MatrixWalk=output
On Error Resume Next
output=Join(a)
Me!MatrixJoinError=CStr(Err.Number)
Err.Clear
output=Filter(a,"x")
Me!MatrixFilterError=CStr(Err.Number)
Err.Clear
output=CStr(a(1))
Me!MatrixIndexError=CStr(Err.Number)
Err.Clear
ReDim Preserve a(1 To 3)
Me!MatrixPreserveError=CStr(Err.Number)
On Error GoTo 0
Me!MatrixRetained=CStr(a(2,-1))
ReDim Preserve a(1 To 2,-1 To 1)
Me!MatrixGrown=CStr(UBound(a,2)) & ":" & CStr(a(1,0)) & ":" & CStr(a(2,0)) & ":" & TypeName(a(1,1))
On Error Resume Next
ReDim Preserve a(1 To 3,-1 To 1)
Me!MatrixWrongResize=CStr(Err.Number)
On Error GoTo 0
ReDim Preserve a(1 To 2,-1 To -1)
Me!MatrixShrunk=CStr(UBound(a,2)) & ":" & CStr(a(2,-1))
ReDim a(0 To 1,1 To 2,4 To 5)
a(1,2,5)=42
Me!MatrixCube=CStr(LBound(a,3)) & ":" & CStr(UBound(a,3)) & ":" & CStr(a(1,2,5))
ReDim a(2,3)
a(2,3)=55
Me!MatrixBase=CStr(LBound(a,1)) & ":" & CStr(LBound(a,2)) & ":" & CStr(a(2,3))
End Sub
)VBA","UseMatrix",host);
        assert(matrix.status==VbaExecutionStatus::Completed);
        assert(host.fields["MatrixBounds"]=="1:2:-1:0" && host.fields["MatrixCopies"]=="99:20" && host.fields["MatrixWalk"]=="10|99|30|40|");
        assert(host.fields["MatrixJoinError"]=="13" && host.fields["MatrixFilterError"]=="13" && host.fields["MatrixIndexError"]=="9" && host.fields["MatrixPreserveError"]=="9" && host.fields["MatrixRetained"]=="99");
        assert(host.fields["MatrixGrown"]=="1:30:40:Empty" && host.fields["MatrixWrongResize"]=="9" && host.fields["MatrixShrunk"]=="-1:99");
        assert(host.fields["MatrixCube"]=="4:5:42" && host.fields["MatrixBase"]=="1:1:55");
        host.bracketHost=previousBracketHost;
    }
    const auto typedMatrix=executeVbaProcedure(R"VBA(Sub TypedMatrices()
Dim a As Variant
Dim saved As Variant
On Error Resume Next
ReDim undeclaredTyped(0) As Long
Me!TypedDeclarationError=CStr(Err.Number)
On Error GoTo 0
ReDim a(0 To 1,1 To 2) As Integer
a(1,2)=2.5
saved=a
Me!TypedMatrix=TypeName(a) & ":" & CStr(VarType(a)) & ":" & CStr(VarType(a(0,1))) & ":" & CStr(a(0,1)) & ":" & CStr(a(1,2))
On Error Resume Next
a(1,2)=40000
Me!TypedOverflow=CStr(Err.Number)
Err.Clear
ReDim Preserve a(0 To 1,1 To 3) As Long
Me!TypedPreserveError=CStr(Err.Number)
On Error GoTo 0
ReDim Preserve a(0 To 1,1 To 3) As Integer
Me!TypedPreserved=CStr(a(1,2)) & ":" & CStr(a(1,3)) & ":" & CStr(UBound(saved,2))
ReDim a(0 To 1) As String
a(0)=42
ReDim Preserve a(0 To 2) As String
Me!TypedText=TypeName(a) & ":" & Join(a,"|")
ReDim a(0 To 1)
Me!TypedReset=TypeName(a) & ":" & Join(a,"|")
ReDim a(0 To 1) As Byte
a(1)=255
Me!TypedBytes=TypeName(a) & ":" & CStr(VarType(a)) & ":" & CStr(VarType(a(1))) & ":" & CStr(a(1))
ReDim a(0) As Currency
a(0)="1.23455"
Me!TypedMoney=TypeName(a) & ":" & CStr(VarType(a(0))) & ":" & CStr(a(0))
On Error Resume Next
Me!TypedNumericJoin=Join(a)
Me!TypedNumericJoinError=CStr(Err.Number)
On Error GoTo 0
ReDim a(0) As Date
Me!TypedDate=TypeName(a) & ":" & CStr(VarType(a(0))) & ":" & CStr(CDbl(a(0)))
ReDim a(0) As Boolean
a(0)=1
Me!TypedBoolean=TypeName(a) & ":" & CStr(VarType(a(0))) & ":" & CStr(a(0))
End Sub
)VBA","TypedMatrices",host);
    assert(typedMatrix.status==VbaExecutionStatus::Completed);
    assert(host.fields["TypedDeclarationError"]=="3251");
    assert(host.fields["TypedMatrix"]=="Integer():8194:2:0:2" && host.fields["TypedOverflow"]=="6" && host.fields["TypedPreserveError"]=="13");
    assert(host.fields["TypedPreserved"]=="2:0:2" && host.fields["TypedText"]=="String():42||" && host.fields["TypedReset"]=="String():|");
    assert(host.fields["TypedBytes"]=="Byte():8209:17:255" && host.fields["TypedMoney"]=="Currency():6:1.2346" && host.fields["TypedDate"]=="Date():7:0" && host.fields["TypedBoolean"]=="Boolean():11:True");
    assert(host.fields["TypedNumericJoinError"]=="3251" && !host.fields["TypedNumericJoin"]);
    const auto declaredArrays=executeVbaProcedure(R"VBA(Option Base 1
Sub RejectTypedReference(ByRef value As Variant)
value=Empty
End Sub
Function ArrayUpper(ByVal value As Variant) As Long
ArrayUpper=UBound(value)
End Function
Sub DeclaredArrays()
Dim a() As Long, b() As Long, words() As String
Dim bytes() As Byte, objects() As Object, values()
Dim item As Variant, trail As String, inspection As Variant
Me!DeclaredInitial=CStr(IsArray(a)) & ":" & TypeName(a) & ":" & CStr(VarType(a))
On Error Resume Next
Me!NoBound=LBound(a)
Me!NoBoundError=CStr(Err.Number)
On Error GoTo 0
ReDim Preserve a(2)
a(1)=7.5
b=a
a(1)=9
inspection=a
Me!DeclaredCopy=CStr(LBound(a)) & ":" & CStr(b(1)) & ":" & CStr(a(1)) & ":" & CStr(ArrayUpper(inspection))
On Error Resume Next
a="bad"
Me!DeclaredScalarError=CStr(Err.Number)
ReDim a(2) As String
Me!DeclaredTypeError=CStr(Err.Number)
Call RejectTypedReference(a)
Me!DeclaredReferenceError=CStr(Err.Number)
On Error GoTo 0
Me!DeclaredRetained=CStr(a(1))
Erase a
Me!DeclaredErased=TypeName(a) & ":" & CStr(IsArray(a))
On Error Resume Next
Me!ErasedBound=UBound(a)
Me!ErasedBoundError=CStr(Err.Number)
On Error GoTo 0
ReDim a(0 To 1,0 To 1)
a(0,0)=1
a(1,0)=2
a(0,1)=3
a(1,1)=4
For Each item In a
trail=trail & CStr(item)
Next
Me!DeclaredTraversal=trail
ReDim words(0 To 1)
words(0)=42
Me!DeclaredStrings=Join(words,"|")
ReDim bytes(0 To 1)
bytes(1)=255
Erase bytes
Me!DeclaredByteErase=TypeName(bytes) & ":" & CStr(VarType(bytes))
ReDim Preserve bytes(0 To 1)
Me!DeclaredByteReset=CStr(bytes(1))
ReDim objects(0 To 0)
Set objects(0)=Nothing
Me!DeclaredObjects=TypeName(objects) & ":" & CStr(VarType(objects(0)))
ReDim values(0 To 0)
values(0)=Null
Me!DeclaredVariants=TypeName(values) & ":" & CStr(IsNull(values(0)))
End Sub
)VBA","DeclaredArrays",host);
    if(declaredArrays.status!=VbaExecutionStatus::Completed) std::cerr<<"Declared array error="<<declaredArrays.errorNumber<<" line="<<declaredArrays.line<<"\n";
    assert(declaredArrays.status==VbaExecutionStatus::Completed);
    assert(host.fields["DeclaredInitial"]=="True:Long():8195" && host.fields["NoBoundError"]=="9");
    assert(host.fields["DeclaredCopy"]=="1:8:9:2" && host.fields["DeclaredScalarError"]=="13" && host.fields["DeclaredTypeError"]=="13");
    assert(host.fields["DeclaredReferenceError"]=="3251" && host.fields["DeclaredRetained"]=="9");
    assert(host.fields["DeclaredErased"]=="Long():True" && host.fields["ErasedBoundError"]=="9" && host.fields["DeclaredTraversal"]=="1234");
    assert(host.fields["DeclaredStrings"]=="42|" && host.fields["DeclaredByteErase"]=="Byte():8209" && host.fields["DeclaredByteReset"]=="0");
    assert(host.fields["DeclaredObjects"]=="Object():9" && host.fields["DeclaredVariants"]=="Variant():True");
    const auto arrayParameters=executeVbaProcedure(R"VBA(Sub FillDeclared(ByRef values() As Long)
ReDim Preserve values(0 To 2)
values(0)=11
values(2)=33
End Sub
Sub ClearDeclared(values() As Long)
Erase values
End Sub
Sub NeedsIntegers(values() As Integer)
values(0)=999
End Sub
Function DeclaredTop(values() As Long) As Long
DeclaredTop=UBound(values)
End Function
Function CopiedTop(ByVal value As Variant) As Long
CopiedTop=UBound(value)
End Function
Sub ArrayParameters()
Dim a() As Long, copy As Variant
Call FillDeclared(a())
Me!ArrayParameterValues=CStr(a(0)) & ":" & CStr(a(2)) & ":" & CStr(DeclaredTop(a()))
copy=a()
Me!ArrayParameterCopy=CStr(CopiedTop(copy))
On Error Resume Next
Call NeedsIntegers(a)
Me!ArrayParameterMismatch=CStr(Err.Number)
Me!DirectArrayValue=CopiedTop(a)
Me!DirectArrayValueError=CStr(Err.Number)
On Error GoTo 0
Me!ArrayParameterRetained=CStr(a(0))
Call ClearDeclared(a)
On Error Resume Next
Me!ArrayParameterBound=UBound(a)
Me!ArrayParameterEraseError=CStr(Err.Number)
On Error GoTo 0
Me!ArrayParameterDetached=CStr(copy(2))
End Sub
)VBA","ArrayParameters",host);
    if(arrayParameters.status!=VbaExecutionStatus::Completed) std::cerr<<"Array parameter error="<<arrayParameters.errorNumber<<" line="<<arrayParameters.line<<"\n";
    assert(arrayParameters.status==VbaExecutionStatus::Completed && host.fields["ArrayParameterValues"]=="11:33:2" && host.fields["ArrayParameterCopy"]=="2");
    assert(host.fields["ArrayParameterMismatch"]=="13" && host.fields["DirectArrayValueError"]=="3251" && host.fields["ArrayParameterRetained"]=="11");
    assert(host.fields["ArrayParameterEraseError"]=="9" && host.fields["ArrayParameterDetached"]=="33");
    const std::string declaredModuleArrays=R"VBA(Private retained() As Long
Public visibleWords() As String
Sub AllocateDeclaredModule()
ReDim retained(2 To 3)
retained(2)=41
ReDim visibleWords(0 To 0)
visibleWords(0)="retained"
End Sub
Sub ReadDeclaredModule()
retained(2)=retained(2)+1
Me!DeclaredModuleValue=CStr(retained(2)) & ":" & Join(visibleWords)
End Sub
Sub EraseDeclaredModule()
Erase retained
On Error Resume Next
Me!ModuleBound=UBound(retained)
Me!ModuleBoundError=CStr(Err.Number)
End Sub
)VBA";
    assert(executeVbaProcedure(declaredModuleArrays,"AllocateDeclaredModule",host).status==VbaExecutionStatus::Completed);
    assert(executeVbaProcedure(declaredModuleArrays,"ReadDeclaredModule",host).status==VbaExecutionStatus::Completed && host.fields["DeclaredModuleValue"]=="42:retained");
    assert(executeVbaProcedure(declaredModuleArrays,"ReadDeclaredModule",host).status==VbaExecutionStatus::Completed && host.fields["DeclaredModuleValue"]=="43:retained");
    assert(executeVbaProcedure(declaredModuleArrays,"EraseDeclaredModule",host).status==VbaExecutionStatus::Completed && host.fields["ModuleBoundError"]=="9");
    const auto fixedArrays=executeVbaProcedure(R"VBA(Option Base 1
Sub ResetFixed(values() As Long)
Erase values
On Error Resume Next
ReDim values(9)
Me!FixedHelperResizeError=CStr(Err.Number)
On Error GoTo 0
values(1,-1)=22
End Sub
Sub FixedArrays()
Dim a(2,-1 To 0) As Long, words(0 To 1) As String
Dim bytes(0 To 1) As Byte, refs(0 To 0) As Object, variants(0 To 0)
Dim before As Variant
a(1,-1)=9
a(1,0)=7
before=a()
Erase a
Me!FixedBounds=CStr(LBound(a)) & ":" & CStr(UBound(a)) & ":" & CStr(LBound(a,2)) & ":" & CStr(UBound(a,2)) & ":" & CStr(a(1,-1))
Call ResetFixed(a())
Me!FixedHelperValue=CStr(a(1,-1)) & ":" & CStr(before(1,-1))
On Error Resume Next
ReDim a(3)
Me!FixedResizeError=CStr(Err.Number)
a=Array(1)
Me!FixedAssignmentError=CStr(Err.Number)
On Error GoTo 0
Me!FixedRetained=CStr(a(1,-1)) & ":" & CStr(UBound(a))
words(0)="hello"
Erase words
Me!FixedWords=Join(words,"|") & ":" & CStr(UBound(words))
bytes(1)=255
Erase bytes
Me!FixedBytes=CStr(bytes(1)) & ":" & CStr(UBound(bytes))
Set refs(0)=Nothing
Erase refs
Me!FixedObjects=CStr(VarType(refs(0))) & ":" & CStr(UBound(refs))
variants(0)=Null
Erase variants
Me!FixedVariants=CStr(IsEmpty(variants(0))) & ":" & CStr(UBound(variants))
End Sub
)VBA","FixedArrays",host);
    if(fixedArrays.status!=VbaExecutionStatus::Completed) std::cerr<<"Fixed array error="<<fixedArrays.errorNumber<<" line="<<fixedArrays.line<<"\n";
    assert(fixedArrays.status==VbaExecutionStatus::Completed);
    assert(host.fields["FixedBounds"]=="1:2:-1:0:0" && host.fields["FixedHelperValue"]=="22:9");
    assert(host.fields["FixedHelperResizeError"]=="3251" && host.fields["FixedResizeError"]=="3251" && host.fields["FixedAssignmentError"]=="3251" && host.fields["FixedRetained"]=="22:2");
    assert(host.fields["FixedWords"]=="|:1" && host.fields["FixedBytes"]=="0:1" && host.fields["FixedObjects"]=="9:0" && host.fields["FixedVariants"]=="True:0");
    const std::string fixedModuleArrays=R"VBA(Private counters(1 To 2) As Long
Sub BumpFixedModule()
counters(1)=counters(1)+1
Me!FixedModuleValue=CStr(counters(1))
End Sub
Sub ResetFixedModule()
Erase counters
Me!FixedModuleReset=CStr(counters(1)) & ":" & CStr(LBound(counters)) & ":" & CStr(UBound(counters))
End Sub
)VBA";
    assert(executeVbaProcedure(fixedModuleArrays,"BumpFixedModule",host).status==VbaExecutionStatus::Completed && host.fields["FixedModuleValue"]=="1");
    assert(executeVbaProcedure(fixedModuleArrays,"BumpFixedModule",host).status==VbaExecutionStatus::Completed && host.fields["FixedModuleValue"]=="2");
    assert(executeVbaProcedure(fixedModuleArrays,"ResetFixedModule",host).status==VbaExecutionStatus::Completed && host.fields["FixedModuleReset"]=="0:1:2");
    const std::string constantsModule=R"VBA(Option Compare Text
Private Const BASE_LO As Long = -1
Public Const BASE_HI As Long = 2+1, HEADER As String = "a,b"
Private Const MASK As Long = vbYesNo Or vbQuestion
Private fixedBuffer(BASE_LO To BASE_HI+1) As Long
Sub ChangeConstantCopy(ByRef value As Long)
value=99
End Sub
Sub UseConstants()
Const COUNT As Long = BASE_HI+1
Const TEXT As String = HEADER & ":" & "tail"
Const STAMP As Date = #2024-02-29#
Const MONEY As Currency = 1.25
Const SAME As Boolean = "a"="A"
Dim localArray(BASE_LO To (COUNT-1)) As Long
localArray(BASE_LO)=7
fixedBuffer(BASE_LO)=fixedBuffer(BASE_LO)+1
Call ChangeConstantCopy(COUNT)
Me!ConstantValue=CStr(COUNT) & ":" & TEXT & ":" & CStr(MASK) & ":" & CStr(SAME)
Me!ConstantBounds=CStr(LBound(localArray)) & ":" & CStr(UBound(localArray)) & ":" & CStr(localArray(-1))
Me!ConstantModuleBounds=CStr(LBound(fixedBuffer)) & ":" & CStr(UBound(fixedBuffer)) & ":" & CStr(fixedBuffer(-1))
Me!ConstantTypes=CStr(VarType(STAMP)) & ":" & CStr(VarType(MONEY)) & ":" & CStr(MONEY)
End Sub
Sub RejectConstantWrite()
Const IMMUTABLE As Long = 5
Me!BeforeConstantWrite="must-not-run"
IMMUTABLE=9
End Sub
)VBA";
    const auto constantState=createVbaModuleState(constantsModule);
    assert(constantState->constants.contains("base_hi") && constantState->publicVariables.contains("base_hi") && !constantState->publicVariables.contains("base_lo"));
    const auto constantsResult=executeVbaProcedure(constantsModule,"UseConstants",host);
    if(constantsResult.status!=VbaExecutionStatus::Completed) std::cerr<<"Constant error="<<constantsResult.errorNumber<<" line="<<constantsResult.line<<"\n";
    assert(constantsResult.status==VbaExecutionStatus::Completed);
    assert(host.fields["ConstantValue"]=="4:a,b:tail:36:True" && host.fields["ConstantBounds"]=="-1:3:7");
    assert(host.fields["ConstantModuleBounds"]=="-1:4:1" && host.fields["ConstantTypes"]=="7:6:1.25");
    assert(executeVbaProcedure(constantsModule,"UseConstants",host).status==VbaExecutionStatus::Completed && host.fields["ConstantModuleBounds"]=="-1:4:2");
    const auto constantWrite=executeVbaProcedure(constantsModule,"RejectConstantWrite",host);
    assert(constantWrite.status==VbaExecutionStatus::Unsupported && constantWrite.errorNumber==3251 && !host.fields["BeforeConstantWrite"]);
    const std::string enumModule=R"VBA(Private Const ENUM_BASE As Long=4
Public Enum NativeState
FirstState
NextState=ENUM_BASE
FinalState
AliasState=NativeState.NextState
Count=8
End Enum
Private Enum NativeFailure
FailureState=-2
FollowingFailure
End Enum
Private enumBuffer(FirstState To NativeState.FinalState) As Long
Function SelectState(Optional ByVal value As Long=NativeState.FinalState) As Long
SelectState=value
End Function
Sub UseEnums()
enumBuffer(FinalState)=SelectState()
Me!EnumValues=CStr(FirstState) & ":" & CStr(NextState) & ":" & CStr(FinalState) & ":" & CStr(AliasState) & ":" & CStr(FailureState) & ":" & CStr(FollowingFailure) & ":" & CStr(VarType(FinalState))
Me!EnumQualified=CStr(NativeState . NextState)
Me!EnumBounds=CStr(LBound(enumBuffer)) & ":" & CStr(UBound(enumBuffer)) & ":" & CStr(enumBuffer(FinalState))
End Sub
Sub ShadowEnumNamespace()
Dim NativeState As Collection
Set NativeState=New Collection
NativeState.Add "kept"
Me!EnumRootShadow=CStr(NativeState.Count)
End Sub
Sub RejectEnumWrite()
Me!BeforeEnumWrite="must-not-run"
NativeState . NextState=99
End Sub
)VBA";
    const auto enumResult=executeVbaProcedure(enumModule,"UseEnums",host);
    if(enumResult.status!=VbaExecutionStatus::Completed) std::cerr<<"Enum error="<<enumResult.errorNumber<<" line="<<enumResult.line<<"\n";
    assert(enumResult.status==VbaExecutionStatus::Completed && host.fields["EnumValues"]=="0:4:5:4:-2:-1:3" && host.fields["EnumBounds"]=="0:5:5" && host.fields["EnumQualified"]=="4");
    assert(executeVbaProcedure(enumModule,"ShadowEnumNamespace",host).status==VbaExecutionStatus::Completed && host.fields["EnumRootShadow"]=="1");
    const auto enumState=host.moduleState(enumModule);
    assert(enumState->constants.contains("finalstate") && enumState->publicVariables.contains("finalstate") && !enumState->publicVariables.contains("failurestate"));
    const auto enumWrite=executeVbaProcedure(enumModule,"RejectEnumWrite",host);
    assert(enumWrite.status==VbaExecutionStatus::Unsupported && enumWrite.errorNumber==3251 && !host.fields["BeforeEnumWrite"]);
    const auto enumOverflow=validateVbaProcedure("Enum Overflowing\nLastValue=2147483647\nOverflowValue\nEnd Enum\nSub CheckEnum()\nEnd Sub","CheckEnum");
    assert(enumOverflow.status==VbaExecutionStatus::Unsupported && enumOverflow.errorNumber==6);
    const auto badConstant=validateVbaProcedure(R"VBA(Sub BadConstant()
Const size As Long=FileLen("missing.dat")
End Sub
)VBA","BadConstant");
    assert(badConstant.status==VbaExecutionStatus::Unsupported && badConstant.errorNumber==3251);
    assert(evaluate("vbNewLine").asText()=="\r\n" && evaluate("vbCr & vbLf").asText()=="\r\n");
    assert(evaluate("Len(vbNullChar)").asNumber()==1 && evaluate("vbNullString").asText().empty());
    const auto foldedDefaults=executeVbaProcedure(R"VBA(Private Const DEFAULT_COUNT As Long=5
Function FoldedDefault(Optional ByVal count As Long=DEFAULT_COUNT*2+1,Optional ByVal prefix As String="a" & "b",Optional ByVal stamp As Date=#2024-02-29#) As String
FoldedDefault=CStr(count) & ":" & prefix & ":" & CStr(Year(stamp))
End Function
Function IndependentDefault(Optional ByRef value As Long=DEFAULT_COUNT) As Long
value=value+1
IndependentDefault=value
End Function
Sub FoldedDefaults()
Me!FoldedDefault=FoldedDefault()
Me!FoldedNamed=FoldedDefault(prefix:="explicit")
Me!FoldedEmpty=FoldedDefault(Empty)
Me!FoldedCopies=CStr(IndependentDefault()) & ":" & CStr(IndependentDefault())
End Sub
)VBA","FoldedDefaults",host);
    if(foldedDefaults.status!=VbaExecutionStatus::Completed) std::cerr<<"Default fold error="<<foldedDefaults.errorNumber<<" line="<<foldedDefaults.line<<"\n";
    assert(foldedDefaults.status==VbaExecutionStatus::Completed && host.fields["FoldedDefault"]=="11:ab:2024" && host.fields["FoldedNamed"]=="11:explicit:2024");
    assert(host.fields["FoldedEmpty"]=="0:ab:2024" && host.fields["FoldedCopies"]=="6:6");
    const auto badDefault=validateVbaProcedure(R"VBA(Sub BadDefault(Optional count As Long=FileLen("missing.dat"))
End Sub
)VBA","BadDefault");
    assert(badDefault.status==VbaExecutionStatus::Unsupported && badDefault.errorNumber==3251);
    {
        const std::string cachedSource=R"VBA(Private Const BASE As Long=2
Private calls As Long
Function Cached(Optional ByVal value As Long=BASE) As Long
calls=calls+1
Cached=value*10+calls
End Function
)VBA";
        FakeForm firstCacheHost,secondCacheHost;
        assert(validateVbaProcedure(cachedSource,"Cached").status==VbaExecutionStatus::Completed);
        const auto first=executeVbaProcedure(cachedSource,"Cached",firstCacheHost);
        assert(first.status==VbaExecutionStatus::Completed && first.typedReturnValue->number==21);
        const auto again=executeVbaProcedure(cachedSource,"cached",firstCacheHost);
        const auto separate=executeVbaProcedure(cachedSource,"Cached",secondCacheHost);
        assert(again.status==VbaExecutionStatus::Completed && again.typedReturnValue->number==22);
        assert(separate.status==VbaExecutionStatus::Completed && separate.typedReturnValue->number==21);
        auto edited=cachedSource; edited.replace(edited.find("BASE As Long=2"),std::string("BASE As Long=2").size(),"BASE As Long=3");
        const auto changed=executeVbaProcedure(edited,"Cached",firstCacheHost);
        assert(changed.status==VbaExecutionStatus::Completed && changed.typedReturnValue->number==31);
    }
    const auto dictionaries=executeVbaProcedure(R"VBA(Sub Dictionaries()
Dim d As Object, other As Object, held As Object
Dim c As Collection
Dim key As Variant, item As Variant, keys As Variant, copy As Variant, values As Variant
Dim total As Long
Set d=VBA.CreateObject(Class:="Scripting.Dictionary")
d.Add Key:="one",Item:=1
d("two")=2
d.Item("three")=3
d.Add CLng(1),"numeric"
d.Add "1","text"
Me!DictionaryDistinct=CStr(d.Exists(1)) & ":" & d("1") & ":" & CStr(d.Count)
d.Key("three")="renamed"
keys=d.Keys
Me!DictionaryKeys=Join(keys,"|")
Me!DictionaryMissing=CStr(IsEmpty(d("missing"))) & ":" & CStr(d.Count)
d.Remove "missing"
d.Remove 1
For Each item In d.Items
If IsNumeric(item) Then total=total+item
Next
Me!DictionaryTotal=CStr(total) & ":" & CStr(UBound(keys))
For Each key In d
On Error Resume Next
d.Add "forbidden",9
Me!DictionaryEnumerationError=CStr(Err.Number)
On Error GoTo 0
Exit For
Next
d.Add "after",4
For Each key In d.Keys
d.Remove key
Next
Me!DictionarySnapshotErase=CStr(d.Count)
values=Array(10,20)
d("array")=values
values(0)=99
copy=d("array")
copy(0)=77
Me!DictionaryArray=CStr(d("array")(0))
Set other=New Scripting.Dictionary
Set d("object")=other
other.Add "x",5
Set held=d("object")
Me!DictionaryObject=TypeName(held) & ":" & CStr(held("x"))
On Error Resume Next
d.Add "self",d
Me!DictionarySelfError=CStr(Err.Number)
On Error GoTo 0
Set c=New Collection
c.Add d
On Error Resume Next
d.Add "cycle",c
Me!DictionaryCycleError=CStr(Err.Number)
On Error GoTo 0
Set c=New Collection
d.Add "collection",c
On Error Resume Next
c.Add d
Me!CollectionDictionaryCycleError=CStr(Err.Number)
On Error GoTo 0
Set other=CreateObject("Scripting.Dictionary")
other.Add "A",1
other.Add "a",2
Me!DictionaryBinary=CStr(other.Count)
other.RemoveAll
other.CompareMode=vbTextCompare
other.Add "A",1
Me!DictionaryText=CStr(other.Exists("a")) & ":" & CStr(other.CompareMode)
On Error Resume Next
other.Add "a",2
Me!DictionaryDuplicateError=CStr(Err.Number)
other.CompareMode=0
Me!DictionaryCompareError=CStr(Err.Number)
other.Add "ä",3
Me!DictionaryRegionalError=CStr(Err.Number)
Set held=CreateObject("unsupported.class")
Me!DictionaryClassError=CStr(Err.Number)
On Error GoTo 0
End Sub
)VBA","Dictionaries",host);
    if(dictionaries.status!=VbaExecutionStatus::Completed) std::cerr<<"Dictionary error="<<dictionaries.errorNumber<<" line="<<dictionaries.line<<"\n";
    assert(dictionaries.status==VbaExecutionStatus::Completed);
    assert(host.fields["DictionaryDistinct"]=="True:text:5" && host.fields["DictionaryKeys"]=="one|two|renamed|1|1" && host.fields["DictionaryMissing"]=="True:6");
    assert(host.fields["DictionaryTotal"]=="6:4" && host.fields["DictionaryEnumerationError"]=="3251" && host.fields["DictionarySnapshotErase"]=="0");
    assert(host.fields["DictionaryArray"]=="10" && host.fields["DictionaryObject"]=="Dictionary:5");
    assert(host.fields["DictionarySelfError"]=="3251" && host.fields["DictionaryCycleError"]=="3251" && host.fields["CollectionDictionaryCycleError"]=="3251");
    assert(host.fields["DictionaryBinary"]=="2" && host.fields["DictionaryText"]=="True:1" && host.fields["DictionaryDuplicateError"]=="457" && host.fields["DictionaryCompareError"]=="5");
    assert(host.fields["DictionaryRegionalError"]=="3251" && host.fields["DictionaryClassError"]=="3251");
    const std::string retainedObjects=R"VBA(Private bag As Object
Public exposed As Object
Sub Bump()
If bag Is Nothing Then Set bag=CreateObject("Scripting.Dictionary")
If Not bag.Exists("calls") Then bag.Add "calls",0
bag("calls")=bag("calls")+1
Set exposed=bag
Me!RetainedCalls=CStr(bag("calls"))
End Sub
Sub Clear()
Set bag=Nothing
Set exposed=Nothing
End Sub
)VBA";
    FakeForm independentObjects;
    for(int call=1;call<=2;++call) {
        const auto result=executeVbaProcedure(retainedObjects,"Bump",host);
        assert(result.status==VbaExecutionStatus::Completed && host.fields["RetainedCalls"]==std::to_string(call));
    }
    const auto retainedState=host.moduleState(retainedObjects);
    assert(retainedState->publicVariables.contains("exposed") && !retainedState->publicVariables.contains("bag"));
    assert(retainedState->values.at("bag").object==retainedState->values.at("exposed").object);
    assert(executeVbaProcedure(retainedObjects,"Bump",independentObjects).status==VbaExecutionStatus::Completed);
    assert(independentObjects.fields["RetainedCalls"]=="1");
    assert(executeVbaProcedure(retainedObjects,"Clear",host).status==VbaExecutionStatus::Completed);
    assert(!retainedState->values.at("bag").object && !retainedState->values.at("exposed").object);
    assert(executeVbaProcedure(retainedObjects,"Bump",host).status==VbaExecutionStatus::Completed && host.fields["RetainedCalls"]=="1");
    bool rejectedScalarObject=false;
    try { (void)coerceVbaModuleValue(VbaValue::longInteger(1),"object"); }
    catch(const VbaRuntimeError& error) { rejectedScalarObject=error.number==13; }
    assert(rejectedScalarObject);

    const std::string objectArguments=R"VBA(Sub ByValue(ByVal item As Object)
item.Add "shared",1
Set item=CreateObject("Scripting.Dictionary")
item.Add "private",2
End Sub
Sub ByReference(ByRef item As Object)
Set item=CreateObject("Scripting.Dictionary")
item.Add "replacement",3
End Sub
Function OptionalItem(Optional ByRef item As Object=Nothing) As Long
If IsMissing(item) Then Err.Raise 5
If item Is Nothing Then Set item=CreateObject("Scripting.Dictionary")
item.Add "local",1
OptionalItem=item.Count
End Function
Sub Exercise()
Dim item As Object, objects As Collection, entry As Object
Set item=CreateObject("Scripting.Dictionary")
ByValue item
Me!ObjectByValue=CStr(item.Exists("shared")) & ":" & CStr(item.Exists("private"))
ByReference item
Me!ObjectByReference=CStr(item.Exists("replacement"))
Me!ObjectOptional=CStr(OptionalItem()) & ":" & CStr(OptionalItem())
Set objects=New Collection
objects.Add item
For Each entry In objects
Me!ObjectIteration=CStr(entry.Exists("replacement"))
Next entry
objects.Add 10
On Error Resume Next
For Each entry In objects
Next entry
Me!ObjectIterationError=CStr(Err.Number)
End Sub
)VBA";
    const auto boundObjects=executeVbaProcedure(objectArguments,"Exercise",host);
    assert(boundObjects.status==VbaExecutionStatus::Completed);
    assert(host.fields["ObjectByValue"]=="True:False" && host.fields["ObjectByReference"]=="True");
    assert(host.fields["ObjectOptional"]=="1:1" && host.fields["ObjectIteration"]=="True" && host.fields["ObjectIterationError"]=="13");
    std::vector<VbaValue> invalidObjectArgument{VbaValue::longInteger(1)};
    const auto invalidObjectBinding=executeVbaProcedureWithArguments(objectArguments,"ByValue",host,invalidObjectArgument);
    assert(invalidObjectBinding.status==VbaExecutionStatus::Error && invalidObjectBinding.errorNumber==13);
    const auto invalidObjectDefault=executeVbaProcedure("Sub Invalid(Optional item As Object=1)\nEnd Sub","Invalid",host);
    assert(invalidObjectDefault.status==VbaExecutionStatus::Unsupported);

    const std::string staticLocals=R"VBA(Function NextStatic() As Long
Static counter As Long, bag As Object
Static fixed(1 To 2) As Long, dynamic() As Long
If bag Is Nothing Then Set bag=CreateObject("Scripting.Dictionary")
counter=counter+1
fixed(1)=fixed(1)+2
If counter=1 Then ReDim dynamic(0 To 0) As Long
dynamic(0)=dynamic(0)+3
bag("calls")=counter
Me!StaticState=CStr(fixed(1)) & ":" & CStr(dynamic(0)) & ":" & CStr(bag("calls"))
NextStatic=counter
End Function
Function OtherStatic() As Long
Static counter As Long
counter=counter+10
OtherStatic=counter
End Function
Function RecursiveStatic(ByVal depth As Long) As Long
Static calls As Long
Dim ignored As Long
calls=calls+1
If depth>0 Then ignored=RecursiveStatic(depth-1)
RecursiveStatic=calls
End Function
)VBA";
    FakeForm staticHost, independentStaticHost;
    for(int call=1;call<=2;++call) {
        const auto result=executeVbaProcedure(staticLocals,"NextStatic",staticHost);
        assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==call);
    }
    assert(staticHost.fields["StaticState"]=="4:6:2");
    const auto otherStatic=executeVbaProcedure(staticLocals,"OtherStatic",staticHost);
    assert(otherStatic.status==VbaExecutionStatus::Completed && otherStatic.typedReturnValue->asNumber()==10);
    const auto separateStatic=executeVbaProcedure(staticLocals,"NextStatic",independentStaticHost);
    assert(separateStatic.status==VbaExecutionStatus::Completed && separateStatic.typedReturnValue->asNumber()==1);
    const auto editedStatic=executeVbaProcedure(staticLocals+"\n' independent edited source\n","NextStatic",staticHost);
    assert(editedStatic.status==VbaExecutionStatus::Completed && editedStatic.typedReturnValue->asNumber()==1);
    std::vector<VbaValue> staticDepth{VbaValue::longInteger(2)};
    const auto recursiveStatic=executeVbaProcedureWithArguments(staticLocals,"RecursiveStatic",staticHost,staticDepth);
    assert(recursiveStatic.status==VbaExecutionStatus::Completed && recursiveStatic.typedReturnValue->asNumber()==3);
    const auto staticProcedure=executeVbaProcedure("Static Sub UnsupportedStatic()\nMe!UnexpectedStatic=1\nEnd Sub","UnsupportedStatic",staticHost);
    assert(staticProcedure.status==VbaExecutionStatus::Unsupported && !staticHost.fields.contains("UnexpectedStatic"));

    const std::string typedDictionary=R"VBA(Private retained As Scripting.Dictionary
Function GetDictionary() As Scripting.Dictionary
If retained Is Nothing Then Set retained=New Scripting.Dictionary
Set GetDictionary=retained
End Function
Sub UpdateDictionary(ByRef value As Scripting.Dictionary)
If Not value.Exists("calls") Then value.Add "calls",0
value("calls")=value("calls")+1
End Sub
Function OptionalDictionary(Optional ByVal value As Scripting.Dictionary=Nothing) As Boolean
OptionalDictionary=value Is Nothing
End Function
Sub TypedDictionary()
Dim value As Scripting.Dictionary
Set value=GetDictionary()
UpdateDictionary value
Me!TypedDictionaryCalls=CStr(value("calls"))
Me!TypedDictionaryOptional=CStr(OptionalDictionary()) & ":" & CStr(OptionalDictionary(value))
On Error Resume Next
Set value=New Collection
Me!TypedDictionaryError=CStr(Err.Number)
Me!TypedDictionaryIdentity=CStr(value Is retained)
End Sub
)VBA";
    FakeForm typedDictionaryHost;
    for(int call=1;call<=2;++call) {
        const auto result=executeVbaProcedure(typedDictionary,"TypedDictionary",typedDictionaryHost);
        assert(result.status==VbaExecutionStatus::Completed && typedDictionaryHost.fields["TypedDictionaryCalls"]==std::to_string(call));
        assert(typedDictionaryHost.fields["TypedDictionaryOptional"]=="True:False" && typedDictionaryHost.fields["TypedDictionaryError"]=="13" && typedDictionaryHost.fields["TypedDictionaryIdentity"]=="True");
    }

    const std::string nextLists=R"VBA(Sub NextLists()
Dim i As Long,j As Long,total As Long
Dim item As Variant, values As Collection
Dim trace As String
For i=1 To 2
For j=1 To 2
trace=trace & CStr(i) & CStr(j) & ","
Next j,i
Me!NextNumeric=trace & ":" & CStr(i) & ":" & CStr(j)
total=0
For i=1 To 3
total=total+1
For j=2 To 1
Err.Raise 5
Next j,i
Me!NextEmptyInner=CStr(total) & ":" & CStr(i) & ":" & CStr(j)
total=0
For i=1 To 3
For j=1 To 5
total=total+i
Exit For
Next j,i
Me!NextExit=CStr(total) & ":" & CStr(i) & ":" & CStr(j)
j=0
For i=2 To 1
For j=1 To 2
Err.Raise 5
Next j,i
Me!NextEmptyOuter=CStr(i) & ":" & CStr(j)
Set values=New Collection
For i=1 To 2
For Each item In values
Err.Raise 5
Next item,i
values.Add 1
values.Add 2
total=0
For i=1 To 2
For Each item In values
total=total+i*item
Next item,i
values.Add 3
Me!NextCollection=CStr(total) & ":" & CStr(values.Count)
item=7
For i=1 To 2
For Each item In Array()
Err.Raise 5
Next item,i
Me!NextEmptyArray=CStr(i) & ":" & CStr(item)
total=0
Do
For i=1 To 1
For j=1 To 1
total=total+1
Next j,i
Exit Do
Loop
Me!NextDo=CStr(total)
total=0
If True Then For i=1 To 2:For j=1 To 2:total=total+1:Next j,i
Me!NextInline=CStr(total)
End Sub
Sub NextFault()
Dim i As Long,j As Integer
On Error Resume Next
For i=1 To 2
For j=32767 To 32767
Next j,i
Me!NextFault=CStr(Err.Number) & ":" & CStr(i) & ":" & CStr(j)
End Sub
Sub NextInlineFault()
Dim i As Long,j As Integer
On Error Resume Next
If True Then For i=1 To 2:For j=32767 To 32767:Next j,i
Me!NextInlineFault=CStr(Err.Number) & ":" & CStr(i) & ":" & CStr(j)
End Sub
Sub NextLeaseFault()
Dim i As Long,item As Object,values As Collection
Set values=New Collection
values.Add New Scripting.Dictionary
values.Add 3
On Error GoTo caught
For i=1 To 2
For Each item In values
Next item,i
values.Add 4
Me!NextLeaseCount=CStr(values.Count)
Exit Sub
caught:
Me!NextLeaseFault=CStr(Err.Number)
Resume Next
End Sub
)VBA";
    const std::string nextLeaseLabel=R"VBA(Sub NextLeaseLabel()
Dim i As Long,item As Object,values As Collection
Set values=New Collection
values.Add New Scripting.Dictionary
values.Add 3
On Error GoTo caught
For i=1 To 2
For Each item In values
Next item,i
Exit Sub
aftergroup:
values.Add 4
Me!NextLabelCount=CStr(values.Count)
Exit Sub
caught:
Me!NextLabelFault=CStr(Err.Number)
Resume aftergroup
End Sub
)VBA";
    const auto nextListResult=executeVbaProcedure(nextLists,"NextLists",host);
    if(nextListResult.status!=VbaExecutionStatus::Completed) std::cerr<<"Next-list error="<<nextListResult.errorNumber<<" line="<<nextListResult.line<<"\n";
    assert(nextListResult.status==VbaExecutionStatus::Completed);
    assert(host.fields["NextNumeric"]=="11,12,21,22,:3:3" && host.fields["NextEmptyInner"]=="3:4:2" && host.fields["NextExit"]=="6:4:1");
    assert(host.fields["NextEmptyOuter"]=="2:0" && host.fields["NextCollection"]=="9:3" && host.fields["NextEmptyArray"]=="3:7");
    assert(host.fields["NextDo"]=="1" && host.fields["NextInline"]=="4");
    assert(executeVbaProcedure(nextLists,"NextFault",host).status==VbaExecutionStatus::Completed && host.fields["NextFault"]=="6:1:32767");
    assert(executeVbaProcedure(nextLists,"NextInlineFault",host).status==VbaExecutionStatus::Completed && host.fields["NextInlineFault"]=="6:1:32767");
    assert(executeVbaProcedure(nextLists,"NextLeaseFault",host).status==VbaExecutionStatus::Completed && host.fields["NextLeaseFault"]=="13" && host.fields["NextLeaseCount"]=="3");
    assert(executeVbaProcedure(nextLeaseLabel,"NextLeaseLabel",host).status==VbaExecutionStatus::Completed && host.fields["NextLabelFault"]=="13" && host.fields["NextLabelCount"]=="3");
    const auto invalidNext=executeVbaProcedure("Sub InvalidNext()\nDim i As Long,j As Long\nMe!ForbiddenNext=1\nFor i=1 To 2\nFor j=1 To 2\nNext i,j\nEnd Sub","InvalidNext",host);
    assert(invalidNext.status==VbaExecutionStatus::Error && !host.fields.contains("ForbiddenNext"));
    const auto scopedJumps=executeVbaProcedure(R"VBA(Sub ScopedJumps()
Dim items As Collection,item As Variant,total As Long,i As Long
Set items=New Collection
items.Add 1
items.Add 2
For Each item In items
With items
If item=1 Then GoTo ContinueItem
total=total+.Count
End With
ContinueItem:
Next
items.Add 3
For i=1 To 3
If i=2 Then GoTo Tail
total=total+i
Tail:
Next
If True Then GoTo InlineLabel:total=999:InlineLabel:total=total+10
For Each item In items
GoTo ReleasedEnumeration
Next
ReleasedEnumeration:
items.Add 4
Me!ScopedJumps=CStr(total) & ":" & CStr(items.Count) & ":" & CStr(i)
End Sub
)VBA","ScopedJumps",host);
    if(scopedJumps.status!=VbaExecutionStatus::Completed) std::cerr<<"Scoped jump error="<<scopedJumps.errorNumber<<" line="<<scopedJumps.line<<" detail="<<scopedJumps.explanation<<"\n";
    assert(scopedJumps.status==VbaExecutionStatus::Completed && host.fields["ScopedJumps"]=="16:4:4");
    const auto uninitializedNext=executeVbaProcedure("Sub UninitializedNext()\nDim i As Long,j As Long\nMe!ForbiddenUninitializedNext=1\nGoTo finish\nFor i=1 To 2\nFor j=1 To 2\nfinish:\nNext j,i\nEnd Sub","UninitializedNext",host);
    assert(uninitializedNext.status==VbaExecutionStatus::Unsupported && !host.fields.contains("ForbiddenUninitializedNext"));

    const std::string resumeScopes=R"VBA(Sub ResumeBody()
Dim values As Collection,item As Variant,text As String,total As Long,handled As Long
Set values=New Collection
values.Add 1
values.Add 2
text="bad"
On Error GoTo caught
For Each item In values
total=total+CLng(text)
On Error Resume Next
values.Add 9
Me!ResumeLiveMutation=CStr(Err.Number)
On Error GoTo caught
Next
values.Add 3
Me!ResumeBodyResult=CStr(total) & ":" & CStr(values.Count) & ":" & CStr(handled)
Exit Sub
caught:
handled=handled+1
If handled>1 Then Err.Raise 777
text="1"
Resume
End Sub
Sub ResumeLabelBody()
Dim values As Collection,item As Variant,handled As Long
Set values=New Collection
values.Add 1
values.Add 2
On Error GoTo caught
For Each item In values
Err.Raise 13
Next
Exit Sub
afterloop:
values.Add 3
Me!ResumeLabelResult=CStr(values.Count) & ":" & CStr(handled)
Exit Sub
caught:
handled=handled+1
If handled>1 Then Err.Raise 777
Resume afterloop
End Sub
Sub SingleNextFault()
Dim values As Collection,item As Object
Set values=New Collection
values.Add New Scripting.Dictionary
values.Add 3
On Error Resume Next
For Each item In values
Next
Me!SingleNextFault=CStr(Err.Number)
values.Add 4
Me!SingleNextCount=CStr(values.Count)
End Sub
)VBA";
    assert(executeVbaProcedure(resumeScopes,"ResumeBody",host).status==VbaExecutionStatus::Completed && host.fields["ResumeBodyResult"]=="2:3:1" && host.fields["ResumeLiveMutation"]=="3251");
    assert(executeVbaProcedure(resumeScopes,"ResumeLabelBody",host).status==VbaExecutionStatus::Completed && host.fields["ResumeLabelResult"]=="3:1");
    assert(executeVbaProcedure(resumeScopes,"SingleNextFault",host).status==VbaExecutionStatus::Completed && host.fields["SingleNextFault"]=="13" && host.fields["SingleNextCount"]=="3");

    const std::string arrayCellModule=R"VBA(Private cells() As Long
Private cellIndexes As Long
Function CellIndex() As Long
cellIndexes=cellIndexes+1
CellIndex=0
End Function
Function ReadCells() As Long
ReadCells=cells(0,-1)
End Function
Sub IncrementCell(ByRef value As Long)
value=value+1
Me!CellImmediate=CStr(ReadCells())
End Sub
Sub ForwardCell(ByRef value As Long)
Call IncrementCell(value)
End Sub
Sub AliasCells(ByRef first As Long,ByRef second As Long)
first=first+1
second=second+first
End Sub
Function LaterCell() As Long
cells(0,-1)=10
LaterCell=0
End Function
Function RefCell(ByRef value As Long,ByVal ignored As Long) As Long
value=value+1
RefCell=value
End Function
Function ValueCell(ByVal value As Long,ByVal ignored As Long) As Long
ValueCell=value
End Function
Sub ReplaceActiveArray(ByRef value As Long)
ReDim cells(0 To 4)
End Sub
Sub IncrementByte(ByRef value As Byte)
value=value+1
End Sub
Sub ArrayCellCalls()
Dim copy As Variant,bytes() As Byte
ReDim cells(0 To 1,-1 To -1)
cells(0,-1)=1
ForwardCell cells(CellIndex(),-1)
copy=cells
Me!CellAfterLater=CStr(RefCell(cells(CellIndex(),-1),LaterCell()))
cells(0,-1)=1
Me!CellValueAfterLater=CStr(ValueCell(cells(CellIndex(),-1),LaterCell()))
AliasCells cells(0,-1),cells(0,-1)
Me!CellNamed=CStr(RefCell(ignored:=0,value:=cells(CellIndex(),-1)))
On Error Resume Next
ReplaceActiveArray cells(0,-1)
Me!CellReplacementGuard=CStr(Err.Number)
On Error GoTo 0
ReDim bytes(0)
bytes(0)=1
IncrementByte bytes(0)
Me!CellSummary=CStr(cells(0,-1)) & ":" & CStr(copy(0,-1)) & ":" & CStr(cellIndexes) & ":" & CStr(UBound(cells)) & ":" & CStr(bytes(0))
End Sub
)VBA";
    const auto arrayCellResult=executeVbaProcedure(arrayCellModule,"ArrayCellCalls",host);
    if(arrayCellResult.status!=VbaExecutionStatus::Completed) std::cerr<<"Array cell error="<<arrayCellResult.errorNumber<<" line="<<arrayCellResult.line<<" detail="<<arrayCellResult.explanation<<"\n";
    assert(arrayCellResult.status==VbaExecutionStatus::Completed);
    assert(host.fields["CellImmediate"]=="2" && host.fields["CellAfterLater"]=="11" && host.fields["CellValueAfterLater"]=="1");
    assert(host.fields["CellNamed"]=="23" && host.fields["CellReplacementGuard"]=="3251" && host.fields["CellSummary"]=="23:2:4:1:2");
    const auto indexedReference=executeVbaProcedure(R"VBA(Sub Change(ByRef value As Variant)
value=99
End Sub
Sub ChangeElement()
Dim a As Variant
a=Array(1)
On Error Resume Next
Call Change(a(0))
Me!IndexedReferenceError=CStr(Err.Number)
Me!ChangedElement=a(0)
a=Array(Array(1))
Err.Clear
Call Change(a(0)(0))
Me!CompoundReferenceError=CStr(Err.Number)
Me!CompoundUnchanged=a(0)(0)
End Sub
)VBA","ChangeElement",host);
    assert(indexedReference.status==VbaExecutionStatus::Completed);
    assert(host.fields["IndexedReferenceError"]=="0" && host.fields["ChangedElement"]=="99");
    assert(host.fields["CompoundReferenceError"]=="3251" && host.fields["CompoundUnchanged"]=="1");
    const auto arrays=executeVbaProcedure(R"VBA(Option Base 1
Sub Modify(ByRef items As Variant)
items(1)=42
End Sub
Sub PrivateCopy(ByVal items As Variant)
items(1)=99
End Sub
Sub Arrays()
Dim a As Variant, b As Variant
Dim i As Long, total As Long
a=Array(10,20)
b=a
a(1)=30
Call Modify(a)
Call PrivateCopy(a)
For i=LBound(a) To UBound(a)
total=total+a(i)
Next
Me!ArrayResult=CStr(a(1)) & ":" & CStr(b(1)) & ":" & CStr(total)
On Error Resume Next
a(0)=7
Me!ArrayError=CStr(Err.Number)
End Sub
)VBA","Arrays",host);
    if(arrays.status!=VbaExecutionStatus::Completed) std::cerr<<"Array error="<<arrays.errorNumber<<" line="<<arrays.line<<"\n";
    assert(arrays.status==VbaExecutionStatus::Completed);
    assert(host.fields["ArrayResult"]=="42:10:62" && host.fields["ArrayError"]=="9");
    const auto resized=executeVbaProcedure(R"VBA(Option Base 1
Sub ResizeVariants()
Dim a As Variant, copy As Variant
ReDim a(-2 To 0)
a(-2)=7
a(0)=9
copy=a
ReDim Preserve a(-2 To 2)
Me!Grown=CStr(a(-2)) & ":" & CStr(a(0)) & ":" & TypeName(a(2))
ReDim Preserve a(-2 To -1)
Me!Shrunk=CStr(UBound(a)) & ":" & CStr(copy(0))
On Error Resume Next
ReDim Preserve a(-1 To 3)
Me!ResizeError=CStr(Err.Number)
On Error GoTo 0
ReDim a(2)
Me!Reset=CStr(LBound(a)) & ":" & TypeName(a(1))
Erase a
Me!Erased=TypeName(a)
ReDim Preserve a(3)
Me!Reallocated=CStr(UBound(a))
End Sub
)VBA","ResizeVariants",host);
    if(resized.status!=VbaExecutionStatus::Completed) std::cerr<<"Resize error="<<resized.errorNumber<<" line="<<resized.line<<"\n";
    assert(resized.status==VbaExecutionStatus::Completed);
    assert(host.fields["Grown"]=="7:9:Empty" && host.fields["Shrunk"]=="-1:9");
    assert(host.fields["ResizeError"]=="9" && host.fields["Reset"]=="1:Empty");
    assert(host.fields["Erased"]=="Empty" && host.fields["Reallocated"]=="3");
    assert(evaluate("Join(Split(\"a,,b,\",\",\"),\"|\")").asText()=="a||b|");
    assert(evaluate("Join(Split(\"a,b,c\",\",\",2),\"|\")").asText()=="a|b,c");
    assert(evaluate("Join(Split(\"a b\"))").asText()=="a b");
    assert(evaluate("Join(Split(\"oneXtwoxtail\",\"x\",,vbTextCompare),\"|\")").asText()=="one|two|tail");
    assert(evaluate("Join(Split(\"ö😀x😀\",\"😀\"),\"|\")").asText()=="ö|x|");
    assert(evaluate("VarType(Split(\"a b\"))").asNumber()==8200);
    assert(evaluate("TypeName(Split(\"a b\"))").asText()=="String()");
    assert(evaluate("UBound(Split(\"\"))").asNumber()==-1);
    assert(evaluate("UBound(Split(\"a b\",,0))").asNumber()==-1);
    assert(evaluate("Join(Split(\"a b\",\"\"),\"|\")").asText()=="a b");
    substringError("Split(Null)",94); substringError("Split(\"a\",Null)",94);
    substringError("Split(\"a\",,-2)",5); substringError("Join(Array(Null))",94);
    substringError("Join(7)",13);
    assert(evaluate("Join(Filter(Split(\"alpha,beta,alps\",\",\"),\"al\"),\"|\")").asText()=="alpha|alps");
    assert(evaluate("Join(Filter(Array(\"Abc\",\"xyz\"),\"a\",False,vbTextCompare),\"|\")").asText()=="xyz");
    assert(evaluate("UBound(Filter(Array(\"a\"),\"z\"))").asNumber()==-1);
    substringError("Filter(Array(\"a\"),Null)",94);
    const auto stringArrays=executeVbaProcedure(R"VBA(Option Base 1
Sub TextArrays()
Dim words As Variant, original As Variant
words=Split("one,two,three",",")
original=words
words(1)=42
Me!TextArray=CStr(LBound(words)) & ":" & Join(words,"|") & ":" & Join(original,"/")
Me!StringType=TypeName(words) & ":" & CStr(VarType(words))
End Sub
)VBA","TextArrays",host);
    assert(stringArrays.status==VbaExecutionStatus::Completed);
    assert(host.fields["TextArray"]=="0:one|42|three:one/two/three");
    assert(host.fields["StringType"]=="String():8200");
    const auto arrayLoops=executeVbaProcedure(R"VBA(Sub ChangeLater(ByRef source As Variant)
source(1)="changed"
End Sub
Sub ArrayLoops()
Dim words As Variant, original As Variant, item As Variant
Dim result As String
words=Split("a,b,c",",")
original=words
For Each item In words
result=result & item
If item="a" Then
Call ChangeLater(words)
End If
Next
Me!ArrayLoop=result & ":" & Join(original,",")
result=""
For Each item In Filter(Array("alpha","beta","alps"),"al")
result=result & item & ";"
Next
Me!TemporaryLoop=result
For Each item In Array()
Me!EmptyLoop="bad"
Next
On Error Resume Next
For Each item In words
ReDim words(3)
Next
Me!LoopResizeError=CStr(Err.Number)
End Sub
)VBA","ArrayLoops",host);
    assert(arrayLoops.status==VbaExecutionStatus::Completed);
    assert(host.fields["ArrayLoop"]=="achangedc:a,b,c");
    assert(host.fields["TemporaryLoop"]=="alpha;alps;" && !host.fields["EmptyLoop"]);
    assert(host.fields["LoopResizeError"]=="3251");
    assert(evaluate("InStrRev(\"abcabc\",\"bc\")").asNumber()==5);
    assert(evaluate("InStrRev(\"abcabc\",\"bc\",5)").asNumber()==2);
    assert(evaluate("InStrRev(\"AbcaBC\",\"bc\",,vbTextCompare)").asNumber()==5);
    assert(evaluate("InStrRev(\"a😀x😀\",\"😀\")").asNumber()==5);
    assert(evaluate("InStrRev(\"abc\",\"\")").asNumber()==3);
    assert(evaluate("InStrRev(\"abc\",\"a\",4)").asNumber()==0);
    assert(evaluate("VarType(InStrRev(\"a\",\"a\"))").asNumber()==3);
    assert(evaluate("StrReverse(\"abö\")").asText()=="öba");
    assert(evaluate("InStrRev(Null,\"a\")").kind==VbaValue::Kind::Null && evaluate("InStrRev(\"a\",Null)").kind==VbaValue::Kind::Null);
    substringError("InStrRev(\"a\",\"a\",Null)",94); substringError("InStrRev(\"a\",\"a\",0)",5);
    substringError("StrReverse(Null)",94); substringError("StrReverse(\"😀\")",3251);
    const auto optionalArguments=executeVbaProcedure(R"VBA(Function Describe(ByVal quantity As Long, Optional ByVal factor As Long=2, Optional ByVal prefix As String="A,B") As String
Describe=prefix & ":" & CStr(quantity*factor)
End Function
Sub OptionalCounter(Optional ByRef counter As Long=7)
counter=counter+1
Me!DefaultCounter=CStr(counter)
End Sub
Sub OptionalArguments()
Dim counter As Long
Me!DefaultArgs=Describe(3)
Me!ExplicitArgs=Describe(3,4,"custom")
Call OptionalCounter()
counter=10
Call OptionalCounter(counter)
Me!ExplicitCounter=CStr(counter)
Call OptionalCounter()
On Error Resume Next
Me!RequiredMissing=Describe()
Me!MissingArgumentError=CStr(Err.Number)
End Sub
)VBA","OptionalArguments",host);
    if(optionalArguments.status!=VbaExecutionStatus::Completed) std::cerr<<"Optional error="<<optionalArguments.errorNumber<<" line="<<optionalArguments.line<<"\n";
    assert(optionalArguments.status==VbaExecutionStatus::Completed);
    assert(host.fields["DefaultArgs"]=="A,B:6" && host.fields["ExplicitArgs"]=="custom:12");
    assert(host.fields["DefaultCounter"]=="8" && host.fields["ExplicitCounter"]=="11");
    assert(host.fields["MissingArgumentError"]=="450");
    assert(evaluate("\"alpha7\" Like \"a*#\"").truthy());
    assert(evaluate("\"a7\" Like \"[a-z][!a-z]\"").truthy());
    assert(matchVbaLikePattern(u"A7",u"[A-Z]#",false,{}));
    assert(!matchVbaLikePattern(u"A7",u"[!A-Z]#",false,{}));
    assert(!evaluate("\"Abc\" Like \"a*\"").truthy());
    assert(evaluate("\"[x]\" Like \"[[]?]\"").truthy());
    assert(evaluate("\"ab\" Like \"a[]b\"").truthy());
    assert(evaluate("\"😀\" Like \"??\"").truthy());
    assert(evaluate("Null Like \"*\"").kind==VbaValue::Kind::Null);
    assert(evaluate("Not \"abc\" Like \"z*\" And True").truthy());
    substringError("\"x\" Like \"[z-a]\"",93);
    substringError("\"x\" Like \"[\"",93);
    const auto likeText=executeVbaProcedure(R"VBA(Option Compare Text
Sub LikeText()
If "AB7" Like "[a-z][a-z]#" Then
Me!LikeText="matched"
End If
End Sub
)VBA","LikeText",host);
    assert(likeText.status==VbaExecutionStatus::Completed && host.fields["LikeText"]=="matched");
    const auto variadicArguments=executeVbaProcedure(R"VBA(Option Base 1
Function ReadCount(ByVal arguments As Variant) As Long
ReadCount=UBound(arguments)-LBound(arguments)+1
End Function
Function Total(ByVal initial As Long, ParamArray values() As Variant) As Long
Dim item As Variant
Total=initial
Me!ForwardCount=CStr(ReadCount(values))
For Each item In values
Total=Total+item
Next
Me!ParamArrayBounds=CStr(LBound(values)) & ":" & CStr(UBound(values))
End Function
Function Pack(ParamArray values()) As String
On Error Resume Next
values(0)="bad"
Me!ParamWriteError=CStr(Err.Number)
On Error GoTo 0
Pack=Join(values,"|")
End Function
Sub VariadicArguments()
Me!VariadicTotal=CStr(Total(5,1,2,3))
Me!VariadicPack=Pack("a","b","c")
Me!VariadicEmpty=CStr(Total(7))
End Sub
)VBA","VariadicArguments",host);
    if(variadicArguments.status!=VbaExecutionStatus::Completed) std::cerr<<"ParamArray error="<<variadicArguments.errorNumber<<" line="<<variadicArguments.line<<"\n";
    assert(variadicArguments.status==VbaExecutionStatus::Completed);
    assert(host.fields["VariadicTotal"]=="11" && host.fields["VariadicPack"]=="a|b|c");
    assert(host.fields["VariadicEmpty"]=="7" && host.fields["ParamArrayBounds"]=="0:-1");
    assert(host.fields["ForwardCount"]=="0" && host.fields["ParamWriteError"]=="3251");
    const auto missingArguments=executeVbaProcedure(R"VBA(Function Inspect(Optional value As Variant, Optional ByVal count As Long) As String
Inspect=CStr(IsMissing(value)) & ":" & CStr(count) & ":" & TypeName(value)
End Function
Function Forward(Optional value) As String
Forward=Inspect(value)
Me!ForwardMissing=CStr(IsMissing(value))
End Function
Function Pass(Optional value) As Variant
Pass=value
End Function
Sub Capture(Optional value, Optional ByVal count As Long=9)
Me!HoleCall=Inspect(value,count)
End Sub
Sub MissingArguments()
Me!MissingDefault=Inspect()
Me!ExplicitEmpty=Inspect(Empty,3)
Me!ExplicitNull=Inspect(Null)
Me!ErrorNotMissing=Inspect(CVErr(448))
Me!FunctionHole=Inspect(,3)
Call Capture(,4)
Me!ForwardResult=Forward()
Me!ReturnMissing=CStr(IsMissing(Pass()))
End Sub
)VBA","MissingArguments",host);
    if(missingArguments.status!=VbaExecutionStatus::Completed) std::cerr<<"Missing error="<<missingArguments.errorNumber<<" line="<<missingArguments.line<<"\n";
    assert(missingArguments.status==VbaExecutionStatus::Completed);
    assert(host.fields["MissingDefault"]=="True:0:Error" && host.fields["ExplicitEmpty"]=="False:3:Empty");
    assert(host.fields["ExplicitNull"]=="False:0:Null" && host.fields["ErrorNotMissing"]=="False:0:Error");
    assert(host.fields["FunctionHole"]=="True:3:Error" && host.fields["HoleCall"]=="True:4:Error");
    assert(host.fields["ForwardMissing"]=="True" && host.fields["ForwardResult"]=="True:0:Error");
    assert(host.fields["ReturnMissing"]=="True");
    const auto namedArguments=executeVbaProcedure(R"VBA(Function Multiply(ByVal amount As Long, Optional ByVal factor As Long=2) As Long
Multiply=amount*factor
End Function
Sub Change(ByRef first As Long, ByRef second As Long)
first=first+1
second=second+10
End Sub
Sub NamedArguments()
Dim a As Long, b As Long
a=1
b=2
Me!NamedResult=CStr(Multiply(factor:=3,amount:=4))
Me!NamedDefault=CStr(Multiply(amount:=5))
Call Change(second:=b,first:=a)
Me!NamedRefs=CStr(a) & ":" & CStr(b)
Change second:=b,first:=a
Me!ImplicitNamedRefs=CStr(a) & ":" & CStr(b)
On Error Resume Next
Me!UnknownName=Multiply(unknown:=3)
Me!NamedError=CStr(Err.Number)
End Sub
)VBA","NamedArguments",host);
    if(namedArguments.status!=VbaExecutionStatus::Completed) std::cerr<<"Named error="<<namedArguments.errorNumber<<" line="<<namedArguments.line<<"\n";
    assert(namedArguments.status==VbaExecutionStatus::Completed);
    assert(host.fields["NamedResult"]=="12" && host.fields["NamedDefault"]=="10");
    assert(host.fields["NamedRefs"]=="2:12" && host.fields["ImplicitNamedRefs"]=="3:22");
    assert(host.fields["NamedError"]=="448");
    assert(evaluate("Replace(Replace:=\"X\",Expression:=\"a-a\",Find:=\"a\")").asText()=="X-X");
    assert(evaluate("Join(Delimiter:=\"|\",SourceArray:=Split(Delimiter:=\",\",Expression:=\"a,b\"))").asText()=="a|b");
    assert(evaluate("InStr(String2:=\"b\",String1:=\"abc\")").asNumber()==2);
    assert(evaluate("DateDiff(Date2:=\"2024-03-01\",Interval:=\"d\",Date1:=\"2024-02-28\")").asNumber()==2);
    assert(evaluate("RGB(Blue:=3,Red:=1,Green:=2)").asNumber()==197121);
    assert(evaluate("Mid(Length:=2,String:=\"abcd\",Start:=2)").asText()=="bc");
    substringError("Replace(\"a\",Expression:=\"b\",Find:=\"a\",Replace:=\"x\")",450);
    substringError("Replace(Unknown:=1)",448);
    assert(evaluate("Nz(Null)").kind==VbaValue::Kind::Empty);
    assert(evaluate("Nz(Null,)").kind==VbaValue::Kind::Empty);
    substringError("Nz(,7)",450);
    assert(evaluate("Nz(Null)+2").asNumber()==2 && evaluate("Nz(Null)&2").asText()=="2");
    assert(evaluate("Nz(Null)+\"2\"").asText()=="2");
    assert(evaluate("Nz(ValueIfNull:=7,Value:=Null)").asNumber()==7);
    assert(evaluateAccessExpression("Nz(Null)",host).kind==VbaValue::Kind::Text);
    FakeForm stopHost;
    const auto stopped=executeVbaProcedure(R"VBA(Function Stopper() As Long
DoCmd.Quit
Stopper=1
End Function
Function Forbidden() As Long
Me!Forbidden="bad"
Forbidden=1
End Function
Sub StopFrames()
On Error Resume Next
Me!StopResult=Stopper()+Forbidden()
Me!AfterStop="bad"
End Sub
)VBA","StopFrames",stopHost);
    assert(stopped.status==VbaExecutionStatus::Completed && stopped.exitSub);
    assert(!stopHost.fields["Forbidden"] && !stopHost.fields["AfterStop"] && !stopHost.fields["StopResult"]);
    assert(executeVbaProcedure("Sub Fresh()\nMe!Fresh=1\nEnd Sub","Fresh",stopHost).status==VbaExecutionStatus::Completed && stopHost.fields["Fresh"]=="1");
    stopHost.rejectQuit=true;
    assert(executeVbaProcedure("Sub CancelQuit()\nOn Error Resume Next\nDoCmd.Quit\nMe!CancelError=CStr(Err.Number)\nEnd Sub","CancelQuit",stopHost).status==VbaExecutionStatus::Completed && stopHost.fields["CancelError"]=="2501");
    const auto arrayReturn=executeVbaProcedure(R"VBA(Function ReturnedArray() As Variant
ReturnedArray=Array(4,5)
End Function
Sub ArrayReturn()
Dim a As Variant
a=ReturnedArray()
Me!ArrayReturn=CStr(a(0)) & ":" & CStr(UBound(a))
End Sub
)VBA","ArrayReturn",host);
    assert(arrayReturn.status==VbaExecutionStatus::Completed && host.fields["ArrayReturn"]=="4:1");
    std::vector<VbaValue> byteArguments{VbaValue::byteArray({0,128,255})};
    const auto byteReturn=executeVbaProcedureWithArguments("Function Echo(ByVal value As Variant) As Variant\nEcho=value\nEnd Function","Echo",host,byteArguments);
    assert(byteReturn.status==VbaExecutionStatus::Completed && byteReturn.typedReturnValue->kind==VbaValue::Kind::ByteArray);
    assert(byteReturn.typedReturnValue->bytes==byteArguments[0].bytes && !byteReturn.returnValue);
    const std::string byteEnumeration=R"VBA(Sub ChangeByte(ByRef source As Variant)
source(1)=200
End Sub
Sub PackedBytes(ByVal source As Variant)
Dim item As Variant, count As Long, total As Long
For Each item In source
total=total+item
count=count+1
If count=1 Then ChangeByte source
Me!ByteControlType=TypeName(item)
Next item
Me!ByteEnumeration=CStr(total) & ":" & CStr(count) & ":" & CStr(IsObject(source))
End Sub
Sub TypedBytes()
Dim source(-1 To 0,2 To 3) As Byte
Dim item As Variant, values As String
source(-1,2)=0
source(0,2)=1
source(-1,3)=128
source(0,3)=255
For Each item In source
values=values & CStr(item) & ":"
Next item
Me!TypedByteEnumeration=values & TypeName(item)
End Sub
Sub ReplaceBytes(ByVal source As Variant)
Dim item As Variant
On Error Resume Next
For Each item In source
ReDim source(0 To 2) As Byte
Me!ByteReplacement=CStr(Err.Number)
Exit For
Next item
End Sub
)VBA";
    const auto packedEnumeration=executeVbaProcedureWithArguments(byteEnumeration,"PackedBytes",host,byteArguments);
    assert(packedEnumeration.status==VbaExecutionStatus::Completed && host.fields["ByteEnumeration"]=="455:3:False" && host.fields["ByteControlType"]=="Byte");
    assert(byteArguments[0].bytes==std::vector<unsigned char>({0,128,255}));
    assert(executeVbaProcedure(byteEnumeration,"TypedBytes",host).status==VbaExecutionStatus::Completed && host.fields["TypedByteEnumeration"]=="0:1:128:255:Byte");
    std::vector<VbaValue> emptyEnumeration{VbaValue::byteArray({})};
    assert(executeVbaProcedureWithArguments(byteEnumeration,"PackedBytes",host,emptyEnumeration).status==VbaExecutionStatus::Completed && host.fields["ByteEnumeration"]=="0:0:False");
    assert(executeVbaProcedureWithArguments(byteEnumeration,"ReplaceBytes",host,byteArguments).status==VbaExecutionStatus::Completed && host.fields["ByteReplacement"]=="3251");
    const auto typedPreserve=executeVbaProcedure(R"VBA(Function PreserveStrings() As String
Dim words As Variant
Dim before As Variant
words=Split("a b")
before=words
ReDim Preserve words(0 To 3)
words(3)=23
PreserveStrings=TypeName(words) & ":" & CStr(VarType(words)) & ":" & TypeName(words(2)) & ":" & Join(words,"|") & ":" & Join(before,"|")
ReDim Preserve words(0 To 0)
PreserveStrings=PreserveStrings & ":" & Join(words,"|")
End Function
)VBA","PreserveStrings",host);
    assert(typedPreserve.status==VbaExecutionStatus::Completed && typedPreserve.returnValue=="String():8200:String:a|b||23:a|b:a");
    const auto grownBytes=executeVbaProcedureWithArguments(R"VBA(Function GrowBytes(ByVal value As Variant) As Variant
ReDim Preserve value(0 To 4)
value(4)=254
GrowBytes=value
End Function
)VBA","GrowBytes",host,byteArguments);
    assert(grownBytes.status==VbaExecutionStatus::Completed && grownBytes.typedReturnValue->bytes==std::vector<unsigned char>({0,128,255,0,254}));
    assert(byteArguments[0].bytes==std::vector<unsigned char>({0,128,255}));
    std::vector<VbaValue> erasedBytes{VbaValue::byteArray({0,128,255})};
    const auto erasedByteArray=executeVbaProcedureWithArguments(R"VBA(Sub EraseBytes(ByRef value As Variant)
Dim before As Variant
before=value
Erase value
Me!ByteErase=TypeName(value) & ":" & CStr(UBound(before)) & ":" & CStr(before(1))
End Sub
)VBA","EraseBytes",host,erasedBytes);
    assert(erasedByteArray.status==VbaExecutionStatus::Completed && erasedBytes[0].kind==VbaValue::Kind::Empty && host.fields["ByteErase"]=="Empty:2:128");

    const auto erasedSeveral=executeVbaProcedure(R"VBA(Sub EraseSeveral()
Dim bytes() As Byte, fixed(1 To 2) As Long
Dim values As Variant, before As Variant
ReDim bytes(0 To 1)
bytes(0)=255
fixed(1)=7
values=Array(1,2)
before=values
Erase bytes,fixed,values
Me!MultiErase=TypeName(bytes) & ":" & CStr(fixed(1)) & ":" & CStr(UBound(fixed)) & ":" & TypeName(values) & ":" & CStr(before(1))
On Error Resume Next
Me!InvalidErasedBounds=UBound(bytes)
Me!MultiEraseBoundsError=CStr(Err.Number)
End Sub
)VBA","EraseSeveral",host);
    assert(erasedSeveral.status==VbaExecutionStatus::Completed && host.fields["MultiErase"]=="Byte():0:2:Empty:2" && host.fields["MultiEraseBoundsError"]=="9");
    const auto partialErase=executeVbaProcedure(R"VBA(Sub PartialErase()
Dim first As Variant, invalid As Variant, last As Variant
first=Array(1)
invalid=3
last=Array(9)
On Error Resume Next
Erase first,invalid,last
Me!PartialErase=CStr(Err.Number) & ":" & TypeName(first) & ":" & CStr(last(0))
End Sub
)VBA","PartialErase",host);
    assert(partialErase.status==VbaExecutionStatus::Completed && host.fields["PartialErase"]=="13:Empty:9");
    const auto constantErase=executeVbaProcedure("Sub ConstantErase()\nConst blocked=1\nDim value As Variant\nMe!ForbiddenConstantErase=1\nErase value,blocked\nEnd Sub","ConstantErase",host);
    assert(constantErase.status==VbaExecutionStatus::Unsupported && constantErase.errorNumber==3251 && !host.fields.contains("ForbiddenConstantErase"));
    const std::string resizeSeveral=R"VBA(Private retained() As Long
Function RetainedUpper() As Long
RetainedUpper=UBound(retained)
End Function
Sub ResizeSeveral()
Dim text As Variant, before As Variant
ReDim retained(0 To 2),text(0 To RetainedUpper()) As String
retained(2)=7
text(2)="x"
before=text
ReDim Preserve retained(0 To 3),text(0 To 3)
Me!MultiResize=CStr(UBound(retained)) & ":" & CStr(retained(2)) & ":" & CStr(retained(3)) & ":" & TypeName(text) & ":" & Join(text,"|") & ":" & CStr(UBound(before))
End Sub
Sub PartialResize()
Dim first As Variant, second As Variant, last As Variant
first=Array(9)
second=Array(8)
last=Array(7)
On Error Resume Next
ReDim Preserve first(0 To 2),second(1 To 3),last(0 To 5)
Me!PartialResize=CStr(Err.Number) & ":" & CStr(UBound(first)) & ":" & CStr(UBound(second)) & ":" & CStr(UBound(last))
End Sub
)VBA";
    const auto resizedSeveral=executeVbaProcedure(resizeSeveral,"ResizeSeveral",host);
    assert(resizedSeveral.status==VbaExecutionStatus::Completed && host.fields["MultiResize"]=="3:7:0:String():||x|:2");
    const auto partialResize=executeVbaProcedure(resizeSeveral,"PartialResize",host);
    assert(partialResize.status==VbaExecutionStatus::Completed && host.fields["PartialResize"]=="9:2:0:0");
    const auto constantResize=executeVbaProcedure("Sub ConstantResize()\nConst blocked=1\nDim value As Variant\nMe!ForbiddenConstantResize=1\nReDim value(1),blocked(1)\nEnd Sub","ConstantResize",host);
    assert(constantResize.status==VbaExecutionStatus::Unsupported && constantResize.errorNumber==3251 && !host.fields.contains("ForbiddenConstantResize"));
    const auto shortenedBytes=executeVbaProcedureWithArguments(R"VBA(Sub ShortenBytes(ByRef value As Variant)
ReDim Preserve value(0 To 1)
On Error Resume Next
ReDim Preserve value(1 To 3)
Me!ByteLowerError=Err.Number
End Sub
)VBA","ShortenBytes",host,byteArguments);
    assert(shortenedBytes.status==VbaExecutionStatus::Completed && byteArguments[0].bytes==std::vector<unsigned char>({0,128}) && host.fields["ByteLowerError"]=="9");
    assert(evaluate("VBA.UCase$(String:=\"ab\")").asText()=="AB");
    assert(evaluate("VBA.VarType(VBA.DateSerial(2024,3,0))=VBA.vbDate").truthy());
    assert(evaluate("VBA.TypeName(VBA.Date)").asText()=="Date" && evaluate("VBA.Len(\"ab\")").asNumber()==2);
    int qualifiedError=0; try { (void)evaluate("VBA.CurrentDb()"); } catch(const VbaRuntimeError& error) { qualifiedError=error.number; } assert(qualifiedError==3251);
    qualifiedError=0; try { (void)evaluate("VBA.Len(1,2)"); } catch(const VbaRuntimeError& error) { qualifiedError=error.number; } assert(qualifiedError==450);
    const auto qualifiedArrays=executeVbaProcedure(R"VBA(Option Base 1
Sub QualifiedArrays()
Dim ordinary As Variant
Dim qualified As Variant
Dim item As Variant
Dim joined As String
ordinary=Array(5)
qualified=VBA.Array(6,7)
Me!QualifiedBounds=CStr(LBound(ordinary)) & ":" & CStr(LBound(qualified)) & ":" & CStr(UBound(qualified)) & ":" & CStr(qualified(0))
Me!QualifiedEmpty=CStr(UBound(VBA.Array()))
For Each item In VBA.Array("x","y")
joined=joined & item
Next
Me!QualifiedLoop=joined
End Sub
)VBA","QualifiedArrays",host);
    assert(qualifiedArrays.status==VbaExecutionStatus::Completed && host.fields["QualifiedBounds"]=="1:0:1:6" && host.fields["QualifiedEmpty"]=="-1" && host.fields["QualifiedLoop"]=="xy");
    const auto localNamespace=executeVbaProcedure(R"VBA(Sub LocalNamespace()
Dim VBA As Variant
Dim result As Variant
VBA=123
On Error Resume Next
result=VBA.Array(9)
Me!LocalNamespaceError=Err.Number
End Sub
)VBA","LocalNamespace",host);
    assert(localNamespace.status==VbaExecutionStatus::Completed && host.fields["LocalNamespaceError"]=="91");
    const auto letValues=executeVbaProcedure(R"VBA(Function LetValues() As Long
Dim amount As Long
Dim items As Variant
Let amount=2
items=Array(1)
Let items(0)=amount+3
Let Me!LetAssignment=items(0)
Let LetValues=amount+items(0)
End Function
)VBA","LetValues",host);
    assert(letValues.status==VbaExecutionStatus::Completed && letValues.returnValue=="7" && host.fields["LetAssignment"]=="5");
    const auto invalidLet=executeVbaProcedure("Sub InvalidLet()\nLet MsgBox \"invalid assignment\"\nEnd Sub","InvalidLet",host);
    assert(invalidLet.status==VbaExecutionStatus::Error);
    assert(evaluate("VarType(CInt(3) And CInt(1))=vbInteger").truthy());
    assert(evaluate("VarType(CLng(3) Or CLng(4))=vbLong").truthy());
    assert(evaluate("1.5 And 2").asNumber()==2);
    substringError("2147483648 And 0",6); substringError("-2147483649 Or 0",6);
    assert(evaluate("&H8000").asNumber()==-32768 && evaluate("&HFFFF&").asNumber()==65535);
    assert(evaluate("VarType(&HFFFF)=vbInteger And VarType(&HFFFF&)=vbLong").truthy());
    assert(evaluate("Not &H80000000&").asNumber()==2147483647);
    assert(evaluate("&HFF00& And &H0F0F&").asNumber()==3840);
    assert(evaluate("&O177777").asNumber()==-1 && evaluate("&O37777777777&").asNumber()==-1);
    assert(evaluate("\"mask=\" & &HFF&").asText()=="mask=255");
    substringError("&H100000000",6); substringError("&H10000%",6);
    const auto basedDefault=executeVbaProcedure("Function BasedDefault(Optional value As Long=&HFFFF&) As Long\nBasedDefault=value\nEnd Function","BasedDefault",host);
    assert(basedDefault.status==VbaExecutionStatus::Completed && basedDefault.returnValue=="65535");
    assert(evaluate("AscW(String:=ChrW$(CharCode:=937))").asNumber()==937);
    assert(evaluate("Chr$(CharCode:=34)").asText()=="\"");
    std::cout << "Core tests passed\n";
}
