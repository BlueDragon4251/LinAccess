#include "csv.hpp"
#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include "vba_expression.hpp"

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

int main() {
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
        int actual=0; try { createVbaModuleState("Public rows() As Long\n"); } catch(const VbaRuntimeError& failure) { actual=failure.number; } assert(actual==3251);
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
        FakeForm fractional; result=executeVbaProcedure(module,"FractionalCounter",fractional); assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251 && !fractional.fields.contains("Body"));
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
        assert(result.status == VbaExecutionStatus::Error && result.errorNumber == 424);
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
    assert(host.messages.back().size() == 19);
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
    assert(host.messages.back().size()==19 && host.messages.back()[10]==' ');
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
    assert(evaluate("CInt(2.5)").asNumber()==2 && evaluate("CInt(3.5)").asNumber()==4);
    assert(evaluate("CInt(-2.5)").asNumber()==-2 && evaluate("CInt(True)").asNumber()==-1);
    assert(evaluate("CInt(Empty)").asNumber()==0 && evaluate("CInt(CCur(\"1.5\"))").asNumber()==2);
    assert(evaluate("VarType(CInt(1))=vbInteger").truthy());
    assert(evaluate("VarType(CLng(1))=vbLong").truthy());
    assert(evaluate("VarType(CDbl(CInt(1)))=vbDouble").truthy());
    assert(evaluate("VarType(CInt(1))").asNumber()==2 && evaluate("TypeName(CInt(1))").asText()=="Integer");
    assert(evaluate("VarType(CLng(1))").asNumber()==3 && evaluate("TypeName(CLng(1))").asText()=="Long");
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
    auto scalarProbe=std::make_shared<ScalarProbe>(); scalarProbe->value=VbaValue::string("ö😀");
    const VbaVariables scalarValues{{"probe",VbaValue::reference(scalarProbe)}};
    assert(evaluateVbaExpression("Len(probe)",host,scalarValues).asNumber()==3 && scalarProbe->reads==1);
    scalarProbe->value=VbaValue{};
    assert(evaluateVbaExpression("Len(probe)",host,scalarValues).kind==VbaValue::Kind::Null && scalarProbe->reads==2);
    const VbaVariables unicodeValues{{"embedded",VbaValue::string(std::string("A\0B",3))},{"invalid",VbaValue::string(std::string(1,static_cast<char>(0xff)))}};
    assert(evaluateVbaExpression("Len(embedded)",host,unicodeValues).asNumber()==3);
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
    std::cout << "Core tests passed\n";
}
