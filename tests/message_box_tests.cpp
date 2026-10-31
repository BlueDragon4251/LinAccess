#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QMessageBox>
#include <QPushButton>
#include <QTimer>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

void testMessageBoxes() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","message-box-tests");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE choices(value INTEGER)"));
    const QByteArray definition="Begin Form\nBegin Section\nBegin TextBox\nName =\"answer\"\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Sub Choose(style As Long)
Me!answer=MsgBox("<b>literal</b>",style,"choice")
If Me!answer=vbYes Then
CurrentDb.Execute "INSERT INTO choices VALUES(6)"
End If
End Sub
Sub Missing()
Me!answer=MsgBox("optional", , "")
End Sub
Sub Statement()
MsgBox "statement", vbInformation
Call MsgBox("called", vbOKOnly)
End Sub
Sub Invalid(style As Long)
Me!answer=MsgBox("invalid",style)
End Sub
Sub Help()
Me!answer=MsgBox("help", , , "file")
End Sub
Sub NullPrompt()
Me!answer=MsgBox(Null)
End Sub
Sub Bare()
answer=42
Me.Tag=answer
End Sub
Sub Local()
Dim answer As Long
answer=7
Me.Tag=answer
End Sub
Sub Parameter(answer As Long)
answer=99
Me.Tag=answer
End Sub
Sub Reference()
Dim selected As Object
Set selected=answer
selected.Value=9
answer.SetFocus
End Sub
Sub Typo()
Me.Tag=unavailableControl
End Sub
)VBA";
    AccessFormPreview form("message test",definition,code,db);
    auto object=makePreviewFormObject(form);
    struct Choice { int style; QMessageBox::StandardButton button; int result; };
    const Choice cases[]={{0,QMessageBox::Ok,1},{1,QMessageBox::Cancel,2},{2,QMessageBox::Abort,3},
        {2,QMessageBox::Retry,4},{2,QMessageBox::Ignore,5},{3,QMessageBox::Yes,6},
        {4+32+256,QMessageBox::No,7},{5,QMessageBox::Retry,4}};
    for (const auto& item:cases) {
        bool visited=false;
        QTimer::singleShot(0,&form,[&] {
            auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box);
            assert(box->text()=="<b>literal</b>" && box->textFormat()==Qt::PlainText && box->windowTitle()=="choice");
            if (item.style==4+32+256) {
                assert(box->icon()==QMessageBox::Question && box->defaultButton()==box->button(QMessageBox::No));
                assert(!box->close() && box->isVisible());
                QTest::keyClick(box,Qt::Key_Escape); assert(box->isVisible());
            }
            auto* button=box->button(item.button); assert(button); visited=true; button->click();
        });
        std::vector<linaccess::VbaValue> args{linaccess::VbaValue::numeric(item.style)};
        const auto result=form.executeProcedure("Choose",args);
        assert(visited && result.status==linaccess::VbaExecutionStatus::Completed);
        assert(object->item("answer").asNumber()==item.result);
    }
    assert(query.exec("SELECT COUNT(*) FROM choices WHERE value=6") && query.next() && query.value(0).toInt()==1);
    query.finish();
    std::vector<linaccess::VbaValue> none;
    bool missing=false;
    QTimer::singleShot(0,&form,[&] {
        auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box);
        assert(box->windowTitle().isEmpty() && box->standardButtons()==QMessageBox::Ok);
        missing=true; box->button(QMessageBox::Ok)->click();
    });
    assert(form.executeProcedure("Missing",none).status==linaccess::VbaExecutionStatus::Completed && missing);
    int statements=0;
    QTimer clicker;
    QObject::connect(&clicker,&QTimer::timeout,[&] {
        if (auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) {
            ++statements; box->button(QMessageBox::Ok)->click();
        }
    });
    clicker.start(1);
    assert(form.executeProcedure("Statement",none).status==linaccess::VbaExecutionStatus::Completed);
    clicker.stop(); assert(statements==2);
    for (const int style:{4096,16384,6,768}) {
        std::vector<linaccess::VbaValue> args{linaccess::VbaValue::numeric(style)};
        const auto result=form.executeProcedure("Invalid",args);
        assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==(style==4096 || style==16384 ? 3251 : 5));
    }
    assert(form.executeProcedure("Help",none).errorNumber==3251);
    assert(form.executeProcedure("NullPrompt",none).errorNumber==94);
    assert(form.executeProcedure("Bare",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(object->item("answer").asNumber()==42 && object->member("Tag",{}).asText()=="42");
    assert(form.executeProcedure("Local",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(object->item("answer").asNumber()==42 && object->member("Tag",{}).asText()=="7");
    std::vector<linaccess::VbaValue> parameter{linaccess::VbaValue::numeric(1)};
    assert(form.executeProcedure("Parameter",parameter).status==linaccess::VbaExecutionStatus::Completed);
    assert(object->item("answer").asNumber()==42 && parameter.front().asNumber()==99 && object->member("Tag",{}).asText()=="99");
    form.show(); QApplication::processEvents();
    assert(form.executeProcedure("Reference",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(object->item("answer").asNumber()==9);
    assert(form.executeProcedure("Typo",none).errorNumber==5);
    assert(form.close());
}

#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QRegularExpression>
#include <QLineEdit>
#include <QComboBox>
#include <QAbstractItemView>

void testOriginalMessageBox(const QDir& directory,const QJsonArray& objects) {
    int executed=0;
    for (const auto& entry:objects) {
        const auto artifact=entry.toObject();
        if (artifact["kind"]!="vba" || artifact["sha256"]!="fa72dad21f4875398da7de601750d963c1d5d1292a36f3db6fbc240ed66959ba") continue;
        QFile file(directory.filePath(artifact["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto raw=file.readAll();
        assert(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
        const auto source=linaccess::decodeVbaSource(raw);
        const auto analysis=linaccess::analyzeVba(source.toStdString());
        const auto lines=source.split('\n');
        QString handler,firstInput,secondInput;
        for (std::size_t i=0;i<analysis.procedures.size();++i) {
            const auto& procedure=analysis.procedures[i];
            const int end=i+1<analysis.procedures.size() ? analysis.procedures[i+1].line-1 : lines.size();
            const auto body=lines.mid(procedure.line-1,end-procedure.line+1).join('\n');
            if (body.count(QRegularExpression("\\bMsgBox\\b",QRegularExpression::CaseInsensitiveOption))!=3) continue;
            auto inputs=QRegularExpression("If\\s+IsNull\\((\\w+)\\)",QRegularExpression::CaseInsensitiveOption).globalMatch(body);
            assert(inputs.hasNext()); firstInput=inputs.next().captured(1);
            assert(inputs.hasNext()); secondInput=inputs.next().captured(1);
            handler=QString::fromStdString(procedure.name);
        }
        assert(!handler.isEmpty());
        // Original module, actual native controls, bounded missing-input branch.
        // No original credentials or authentication SQL are used or recorded.
        const auto definition=QString("Begin Form\nBegin Section\nBegin ComboBox\nName =\"%1\"\nRowSourceType =\"Table/Query\"\nRowSource =\"SELECT id,label FROM choices ORDER BY id\"\nColumnCount =2\nEnd\nBegin TextBox\nName =\"%2\"\nEnd\nEnd\nEnd\n").arg(firstInput,secondInput).toUtf8();
        auto db=QSqlDatabase::addDatabase("QSQLITE","original-message-tests"); db.setDatabaseName(":memory:"); assert(db.open());
        QSqlQuery choices(db); assert(choices.exec("CREATE TABLE choices(id INTEGER PRIMARY KEY,label TEXT)"));
        assert(choices.exec("INSERT INTO choices VALUES(1,'first'),(2,'second')"));
        AccessFormPreview form("original message handler",definition,raw,db);
        auto model=makePreviewFormObject(form);
        model->member("Controls",{linaccess::VbaValue::string(firstInput.toStdString())}).object->assign("Value",linaccess::VbaValue::string("selected"));
        model->member("Controls",{linaccess::VbaValue::string(secondInput.toStdString())}).object->assign("Value",linaccess::VbaValue{});
        form.show(); QApplication::processEvents();
        bool visited=false;
        QTimer::singleShot(0,&form,[&] {
            auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box);
            assert(box->textFormat()==Qt::PlainText && !box->text().isEmpty());
            visited=true; box->button(QMessageBox::Ok)->click();
        });
        std::vector<linaccess::VbaValue> args;
        const auto result=form.executeProcedure(handler,args);
        if (!visited || result.status!=linaccess::VbaExecutionStatus::Completed)
            std::cerr << "Original MsgBox path: visited=" << visited << ", status=" << static_cast<int>(result.status) << ", line=" << result.line << ", error=" << result.errorNumber << "\n";
        assert(visited && result.status==linaccess::VbaExecutionStatus::Completed);
        assert(form.focusWidget()==form.findChild<QLineEdit*>("access-control-"+secondInput));
        model->member("Controls",{linaccess::VbaValue::string(firstInput.toStdString())}).object->assign("Value",linaccess::VbaValue{});
        bool firstInputMessage=false;
        QTimer::singleShot(0,&form,[&] {
            auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box);
            firstInputMessage=true; box->button(QMessageBox::Ok)->click();
        });
        const auto missingFirst=form.executeProcedure(handler,args);
        assert(firstInputMessage && missingFirst.status==linaccess::VbaExecutionStatus::Completed);
        auto* combo=form.findChild<QComboBox*>("access-control-"+firstInput);
        assert(combo && combo->count()==2 && combo->view()->isVisible()); combo->hidePopup();
        assert(form.close()); ++executed;
    }
    assert(executed==1);
}
