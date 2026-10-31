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
#include <optional>

namespace {
class TestSequence final : public linaccess::VbaObject {
public:
    std::vector<linaccess::VbaValue> values;
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override {
        throw linaccess::VbaRuntimeError(438,"No members");
    }
    std::vector<linaccess::VbaValue> enumerate() override { return values; }
};
class FakeForm final : public linaccess::VbaHost {
public:
    std::map<std::string, std::optional<std::string>> fields;
    std::vector<std::string> messages;
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
    void saveRecord() override { ++saves; }
    void maximizeForm() override { ++maximized; }
    void gotoLastRecord() override { ++navigatedLast; }
    void gotoNextRecord() override { ++navigatedNext; }
    void gotoNewRecord() override { ++navigatedNew; }
    void requery() override { ++requeries; }
    void executeSql(std::string_view statement) override { sqlStatements.emplace_back(statement); }
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
    assert(evaluate("CLng(2.5)").asNumber() == 2);
    assert(evaluate("CLng(3.5)").asNumber() == 4);
    assert(evaluate("CLng(-2.5)").asNumber() == -2);
    assert(evaluate("Null & Null").kind == VbaValue::Kind::Null);
    assert(evaluate("Nothing Is Nothing").truthy());
    assert(evaluate("Len(Null)").kind == VbaValue::Kind::Null);
    bool invalidNull = false;
    try { (void)evaluate("CStr(Null)"); } catch (...) { invalidNull = true; }
    assert(invalidNull);
    std::cout << "Core tests passed\n";
}
