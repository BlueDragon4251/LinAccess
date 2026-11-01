#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"
#include <QApplication>
#include <QCryptographicHash>
#include <QDialog>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QLineEdit>
#include <QPointer>
#include <QPushButton>
#include <QRegularExpression>
#include <QSet>
#include <QSqlQuery>
#include <QTest>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <memory>

void testEval() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-eval"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(query.exec("CREATE TABLE markers(value INTEGER)"));
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,answer INTEGER)"));
    assert(query.exec("INSERT INTO records VALUES(1,0)"));
    const QByteArray module=R"VBA(Public Function Mark(value As Long) As Long
CurrentDb.Execute "INSERT INTO markers VALUES(" & value & ")"
Mark=value+1
End Function
Private Function Hidden() As Long
CurrentDb.Execute "INSERT INTO markers VALUES(98)"
Hidden=1
End Function
Public Sub NotAFunction()
CurrentDb.Execute "INSERT INTO markers VALUES(99)"
End Sub
Public Function CloseEvalOwner() As Long
DoCmd.Close acForm,"eval closer"
CloseEvalOwner=1
End Function
)VBA";
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','EvalFunctions',?)"); query.addBindValue(module); assert(query.exec());
    const QByteArray definition="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin TextBox\nName =\"answer\"\nControlSource =\"answer\"\nEnd\nBegin TextBox\nName =\"probe\"\nTop =400\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Sub Arithmetic()
Me.answer=Eval("Forms![Eval host]![probe] * 2")
Me.Dirty=False
End Sub
Sub ProjectFunction()
Me.answer=Application.Eval("Mark(7)")
Me.Dirty=False
End Sub
Sub Statements()
Eval("Mark(10)")
Call Eval("Mark(11)")
Eval "Mark(12)"
End Sub
Sub OpenTarget()
DoCmd.OpenForm "eval closer"
End Sub
)VBA";
    AccessFormPreview form("Eval host",definition,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto model=makePreviewFormObject(form); auto app=makePreviewApplicationObject(form);
    model->assignItem("probe",VbaValue::numeric(21));
    assert(app->member("Eval",{VbaValue::string("Forms![Eval host]![probe]")}).asNumber()==21);
    assert(!app->member("Eval",{VbaValue::string("Forms![Eval host]![probe] Is Null")}).truthy());
    model->assignItem("probe",VbaValue{});
    assert(app->member("Eval",{VbaValue::string("Forms![Eval host]![probe] Is Null")}).truthy());
    assert(!app->member("Eval",{VbaValue::string("Forms![Eval host]![probe] Is Not Null")}).truthy());
    model->assignItem("probe",VbaValue::numeric(21));
    std::vector<VbaValue> arguments;
    assert(form.executeProcedure("Arithmetic",arguments).status==VbaExecutionStatus::Completed);
    assert(query.exec("SELECT answer FROM records WHERE id=1") && query.next() && query.value(0).toInt()==42); query.finish();
    assert(form.executeProcedure("ProjectFunction",arguments).status==VbaExecutionStatus::Completed);
    assert(query.exec("SELECT answer FROM records WHERE id=1") && query.next() && query.value(0).toInt()==8); query.finish();
    assert(form.executeProcedure("Statements",arguments).status==VbaExecutionStatus::Completed);
    assert(app->member("Eval",{VbaValue::string("Mark(20)")}).asNumber()==21);
    for (const auto& expression:{"Hidden()","NotAFunction()","Unknown()"}) {
        int number=0; try { app->member("Eval",{VbaValue::string(expression)}); }
        catch (const VbaRuntimeError& failure) { number=failure.number; }
        assert(number==2425);
    }
    assert(query.exec("SELECT count(*),sum(value) FROM markers") && query.next() && query.value(0).toInt()==5 && query.value(1).toInt()==60); query.finish();
    int number=0; try { app->member("Eval",{VbaValue::string("Forms![missing form]![probe]")}); }
    catch (const VbaRuntimeError& failure) { number=failure.number; } assert(number==2450);
    number=0; try { app->member("Eval",{VbaValue::string("Forms![Eval host]")}); }
    catch (const VbaRuntimeError& failure) { number=failure.number; } assert(number==13);
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','eval closer',?)");
    query.addBindValue(QByteArray("Begin Form\nBegin Section\nBegin TextBox\nName =\"probe\"\nEnd\nEnd\nEnd\n")); assert(query.exec());
    assert(form.executeProcedure("OpenTarget",arguments).status==VbaExecutionStatus::Completed); QApplication::processEvents();
    QPointer<AccessFormPreview> closer;
    for (auto* widget:QApplication::topLevelWidgets())
        if (auto* document=dynamic_cast<AccessFormPreview*>(widget); document && document!=&form
            && makePreviewFormObject(*document)->member("Name",{}).asText()=="eval closer") closer=document;
    assert(closer);
    auto closingApp=makePreviewApplicationObject(*closer);
    bool interacted=false;
    QTimer::singleShot(0,[&] {
        QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
        assert(closer); // Closing its owner must not delete the active modal dialog.
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="access-input-box");
        auto* editor=dialog->findChild<QLineEdit*>("access-input-box-editor"); assert(editor); editor->setText("safe");
        interacted=true; dialog->findChild<QPushButton*>("access-input-box-ok")->click();
    });
    const auto closed=closingApp->member("Eval",{VbaValue::string("IIf(CloseEvalOwner(),InputBox('prompt'),'unused')")});
    assert(interacted && closed.asText()=="safe");
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!closer);
    assert(form.close());
}

namespace {
class CountingEvalHost final : public linaccess::VbaHost {
public:
    explicit CountingEvalHost(AccessFormPreview& anchor):anchor_(anchor) {}
    int formReads{};
    std::optional<std::string> field(std::string_view) override { throw linaccess::VbaRuntimeError(438,"No implicit form context"); }
    void setField(std::string_view,std::optional<std::string>) override { throw linaccess::VbaRuntimeError(438,"No implicit form context"); }
    void message(std::string_view) override { throw linaccess::VbaRuntimeError(3251,"Unexpected original dialog"); }
    void saveRecord() override { throw linaccess::VbaRuntimeError(3251,"Unexpected original save"); }
    std::shared_ptr<linaccess::VbaObject> formsCollection() override { ++formReads; return makePreviewFormsObject(anchor_); }
private:
    AccessFormPreview& anchor_;
};
}
void testOriginalEval(const QDir& directory,const QJsonArray& objects) {
    using namespace linaccess;
    QByteArray raw;
    for (const auto& value:objects) {
        const auto object=value.toObject();
        if (object["kind"]!="vba" || object["sha256"].toString().toLower()!="aaea2a6871d7e2249e0e213b6acd74ac5a22ec4e12608be82a053368d16355b1") continue;
        QFile source(directory.filePath(object["file"].toString())); assert(source.open(QIODevice::ReadOnly)); raw=source.readAll();
        assert(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
    }
    assert(!raw.isEmpty()); const auto code=decodeVbaSource(raw); const auto analysis=analyzeVba(code.toStdString()); assert(analysis.procedures.size()==1);
    const QRegularExpression calls("\\bEval\\s*\\(\\s*\"((?:\"\"|[^\"])*)\"\\s*\\)",QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression reference("^\\s*(?:Forms|\\[Forms\\])\\s*!\\s*(?:\\[([^\\]]+)\\]|([\\p{L}_][\\p{L}\\p{N}_]*))\\s*!\\s*(?:\\[([^\\]]+)\\]|([\\p{L}_][\\p{L}\\p{N}_]*))\\s+Is\\s+Null\\s*$",QRegularExpression::CaseInsensitiveOption);
    QMap<QString,QSet<QString>> controls; int references=0;
    auto matches=calls.globalMatch(code);
    while (matches.hasNext()) {
        const auto expression=matches.next().captured(1); const auto match=reference.match(expression); assert(match.hasMatch());
        const auto form=match.captured(1).isEmpty() ? match.captured(2) : match.captured(1);
        const auto control=match.captured(3).isEmpty() ? match.captured(4) : match.captured(3);
        controls[form].insert(control); ++references;
    }
    assert(references==15);
    auto db=QSqlDatabase::addDatabase("QSQLITE","original-eval"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','OriginalEval',?)"); query.addBindValue(raw); assert(query.exec());
    std::vector<std::unique_ptr<AccessFormPreview>> forms;
    const auto escaped=[](QString value) { return value.replace('\\',"\\\\").replace('"',"\\\""); };
    for (auto it=controls.cbegin();it!=controls.cend();++it) {
        QString layout="Begin Form\nBegin Section\n"; int top=0;
        for (const auto& name:it.value()) {
            layout+="Begin TextBox\nName =\""+escaped(name)+"\"\nTop ="+QString::number(top)+"\nWidth =1800\nHeight =300\nEnd\n"; top+=350;
        }
        layout+="End\nEnd\n";
        forms.push_back(std::make_unique<AccessFormPreview>(it.key(),layout.toUtf8(),QByteArray{},db)); forms.back()->show();
    }
    assert(!forms.empty()); QApplication::processEvents();
    CountingEvalHost counted(*forms.front()); auto app=makePreviewApplicationObject(*forms.front());
    int executed=0;
    // Three non-Null states, all Null, each reference individually Null, and a
    // mixed state. Keep original names/expressions out of diagnostic output.
    for (int state=0;state<20;++state) {
        int ordinal=0, nullCount=0;
        for (const auto& form:forms) {
            const auto model=makePreviewFormObject(*form); const auto name=QString::fromStdString(model->member("Name",{}).asText());
            for (const auto& control:controls.value(name)) {
                const bool null=state==3 || (state>=4 && state<19 && ordinal==state-4) || (state==19 && ordinal%2==0);
                if (null) ++nullCount;
                model->assignItem(control.toStdString(),null ? VbaValue{} : VbaValue::numeric(state==0 ? 0 : state==2 ? -1 : 1));
                ++ordinal;
            }
        }
        assert(ordinal==15);
        counted.formReads=0; std::vector<VbaValue> arguments;
        const auto result=executeVbaProcedureWithArguments(code.toStdString(),analysis.procedures.front().name,counted,arguments);
        if (result.status!=VbaExecutionStatus::Completed) std::cerr<<"Original Eval error="<<result.errorNumber<<", line="<<result.line<<'\n';
        // Each Null branch performs one additional original Forms lookup.
        const int expectedReads=15+nullCount;
        if (counted.formReads!=expectedReads) std::cerr<<"Original Eval state="<<state<<", form reads="<<counted.formReads<<'\n';
        assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue && counted.formReads==expectedReads);
        const auto native=app->member("Eval",{VbaValue::string(analysis.procedures.front().name+"()")});
        assert(native.kind==result.typedReturnValue->kind && native.asText()==result.typedReturnValue->asText()); ++executed;
    }
    for (const auto& form:forms) assert(form->close());
    assert(executed==20);
    std::cout<<"Original Eval module: "<<executed<<" native public-function evaluations; 15 expression references and one extra lookup per Null branch; synthetic states, no original rows\n";
}
