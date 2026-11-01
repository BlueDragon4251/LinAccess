#include "access_form_preview.hpp"
#include "vba_expression.hpp"
#include <QApplication>
#include <QClipboard>
#include <QDialog>
#include <QInputMethodEvent>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QSqlQuery>
#include <QTest>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <functional>

void testInputBoxes() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","input-box-tests"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'initial')"));
    const QByteArray definition="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin TextBox\nName =\"answer\"\nControlSource =\"value\"\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Sub Ask()
Me.answer = InputBox("<b>literal</b>","input title","seed",1440,720)
Me.Dirty = False
End Sub
Sub AskDefault()
Me.answer = InputBox$("default prompt")
Me.Dirty = False
End Sub
)VBA";
    AccessFormPreview form("input boxes",definition,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    const auto stored=[&] {
        assert(query.exec("SELECT value FROM records WHERE id=1") && query.next());
        const auto value=query.value(0); query.finish(); assert(!value.isNull()); return value.toString();
    };
    const auto ask=[&](const QString& procedure,const std::function<void(QDialog*,QLineEdit*)>& interact) {
        bool delivered=false;
        QTimer::singleShot(0,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());
            assert(dialog && dialog->objectName()=="access-input-box");
            auto* editor=dialog->findChild<QLineEdit*>("access-input-box-editor"); assert(editor && editor->maxLength()==255);
            if (procedure=="Ask") {
                assert(dialog->windowTitle()=="input title" && editor->text()=="seed" && editor->selectedText()=="seed");
                auto* label=dialog->findChild<QLabel*>(); assert(label && label->text()=="<b>literal</b>" && label->textFormat()==Qt::PlainText);
            } else assert(dialog->windowTitle()=="LinAccess" && editor->text().isEmpty());
            delivered=true; interact(dialog,editor);
        });
        std::vector<linaccess::VbaValue> arguments;
        const auto result=form.executeProcedure(procedure,arguments);
        assert(delivered && result.status==linaccess::VbaExecutionStatus::Completed && !form.hasUnsavedChanges() && form.isVisible());
    };
    ask("Ask",[](QDialog*,QLineEdit* editor) { QTest::keyClicks(editor,"answer's text"); QTest::keyClick(editor,Qt::Key_Return); });
    assert(stored()=="answer's text");
    ask("Ask",[](QDialog*,QLineEdit* editor) { QTest::keyClicks(editor,"discard"); QTest::keyClick(editor,Qt::Key_Escape); });
    assert(stored().isEmpty());
    ask("Ask",[](QDialog* dialog,QLineEdit*) { dialog->findChild<QPushButton*>("access-input-box-cancel")->click(); });
    assert(stored().isEmpty());
    ask("Ask",[](QDialog* dialog,QLineEdit* editor) {
        editor->setText(QString(260,'x')); assert(editor->text().size()==255);
        dialog->findChild<QPushButton*>("access-input-box-ok")->click();
    });
    assert(stored()==QString(254,'x'));
    ask("AskDefault",[](QDialog*,QLineEdit* editor) {
        QApplication::clipboard()->setText("first\r\nsecond"); QTest::keyClick(editor,Qt::Key_V,Qt::ControlModifier);
        assert(editor->text()=="first"); QTest::keyClick(editor,Qt::Key_Return);
    });
    assert(stored()=="first");
    ask("AskDefault",[](QDialog*,QLineEdit* editor) {
        QInputMethodEvent text; text.setCommitString(QString::fromUtf8("Grüße")); QApplication::sendEvent(editor,&text);
        QTest::keyClick(editor,Qt::Key_Return);
    });
    assert(stored()==QString::fromUtf8("Grüße") && form.close());
}
