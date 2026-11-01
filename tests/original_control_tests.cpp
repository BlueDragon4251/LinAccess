#include "access_form_objects.hpp"
#include "access_form_preview.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QRegularExpression>
#include <QSet>
#include <QToolBar>
#include <QAction>
#include <QApplication>
#include <QPushButton>
#include <QTimer>
#include <QMessageBox>
#include <QTemporaryDir>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <algorithm>
#include <iostream>

namespace {
std::vector<bool> visible(const std::vector<linaccess::VbaValue>& controls) {
    std::vector<bool> values;
    for (const auto& control : controls) values.push_back(control.object->member("Visible",{}).boolean);
    return values;
}
}
void testOriginalControls(const QDir& directory,const QJsonArray& objects) {
    const QSet<QByteArray> hashes{
        "57cac4e2bfa6762b8b5b9772ab65b45fbd1c003bbb32fc889e9bb54da019c789",
        "c7abb2cc7ce0a829df92d518cd463cc5cdb5bbcb7000fb88bbe72fda285006c3",
        "2c4d460b8f707760d720093274e4cdcf1f5403fcd33d47fe302e8742922ee57d"};
    int executed = 0;
    for (const auto& item : objects) {
        const auto module = item.toObject();
        if (module["kind"]!="vba") continue;
        QFile file(directory.filePath(module["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto raw = file.readAll();
        if (!hashes.contains(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex())) continue;
        const auto code = linaccess::decodeVbaSource(raw);
        const auto text = code.toUtf8().toStdString();
        const auto analysis = linaccess::analyzeVba(text);
        const auto lines = code.split('\n');
        QByteArray formBytes;
        const auto formName = module["name"].toString().mid(7);
        for (const auto& candidate : objects) {
            const auto form = candidate.toObject();
            if (form["kind"]=="report" && form["name"]==formName) {
                QFile source(directory.filePath(form["file"].toString())); assert(source.open(QIODevice::ReadOnly));
                formBytes = source.readAll(); break;
            }
        }
        assert(!formBytes.isEmpty());
        auto db = QSqlDatabase::addDatabase("QSQLITE","original-controls-" + QString::number(executed));
        db.setDatabaseName(":memory:"); assert(db.open());
        AccessFormPreview preview(formName,formBytes,raw,db);
        assert(preview.eventResult("OnOpen") && preview.eventResult("OnOpen")->status==linaccess::VbaExecutionStatus::Completed);
        auto controls = makePreviewApplicationObject(preview)->member("CommandBars",{linaccess::VbaValue::string("Print Preview")}).object->member("Controls",{}).object->enumerate();
        assert(!controls.empty());
        for (std::size_t i=0;i<analysis.procedures.size();++i) {
            const auto& procedure = analysis.procedures[i];
            const int start = procedure.line-1;
            const int end = i+1<analysis.procedures.size() ? analysis.procedures[i+1].line-1 : lines.size();
            const auto body = lines.mid(start,end-start).join('\n');
            if (!body.contains(QRegularExpression("For\\s+Each[^\\r\\n]*\\bControls\\b",QRegularExpression::CaseInsensitiveOption))) continue;
            std::vector<linaccess::VbaValue> args{linaccess::VbaValue::numeric(0)};
            const auto first = preview.executeProcedure(QString::fromStdString(procedure.name),args);
            if (first.status!=linaccess::VbaExecutionStatus::Completed)
                std::cerr << "Original report handler failed: status=" << static_cast<int>(first.status) << ", line=" << first.line << ", error=" << first.errorNumber << '\n';
            assert(first.status==linaccess::VbaExecutionStatus::Completed);
            const auto zero = visible(controls);
            // Cancel is an unused Report_Open argument in these originals.
            // Restore initial visibility first, then verify the handler restores its
            // own deterministic outcome on the same actual controls.
            for (const auto& control : controls) control.object->assign("Visible",linaccess::VbaValue::logical(true));
            args[0] = linaccess::VbaValue::numeric(0);
            const auto second = preview.executeProcedure(QString::fromStdString(procedure.name),args);
            assert(second.status==linaccess::VbaExecutionStatus::Completed);
            const auto one = visible(controls);
            assert(zero==one);
            assert(std::any_of(one.begin(),one.end(),[](bool value) { return value; }));
            auto* toolbar = preview.findChild<QToolBar*>("access-print-preview-toolbar"); assert(toolbar);
            const auto actions = toolbar->actions(); assert(actions.size()==controls.size());
            for (int index=0;index<actions.size();++index) {
                const bool expected = actions[index]->text()=="&Schließen";
                assert(actions[index]->isVisible()==expected && one[index]==expected);
            }
            QTemporaryDir output; assert(output.isValid());
            const auto pdf = output.filePath("layout.pdf"); assert(preview.exportLayoutPdf(pdf));
            QFile document(pdf); assert(document.open(QIODevice::ReadOnly));
            const auto content = document.readAll();
            assert(content.startsWith("%PDF-") && content.size()>500);
            auto pages = QRegularExpression("/Type\\s*/Page(?!s)").globalMatch(QString::fromLatin1(content));
            int pageCount = 0; while (pages.hasNext()) { pages.next(); ++pageCount; }
            assert(pageCount==1);
            const auto protectedFile = output.filePath("original.mdb");
            QFile sentinel(protectedFile); assert(sentinel.open(QIODevice::WriteOnly)); sentinel.write("unchanged"); sentinel.close();
            assert(!preview.exportLayoutPdf(protectedFile));
            assert(sentinel.open(QIODevice::ReadOnly) && sentinel.readAll()=="unchanged");

            assert(preview.close());
            assert(preview.eventResult("OnClose") && preview.eventResult("OnClose")->status==linaccess::VbaExecutionStatus::Completed);
            ++executed;
        }
    }
    assert(executed==3);
    std::cout << "Original report control handlers executed after resetting visibility: " << executed << '\n';
    int initialized=0,assignments=0;
    for (const auto& item : objects) {
        const auto module=item.toObject(); if (module["kind"]!="vba") continue;
        if (module["sha256"].toString().toLatin1().toLower()!="dbd3623b9404e490a814b70de6318d2c0877969b7feef081696d3e235ae3b427") continue;
        QFile codeFile(directory.filePath(module["file"].toString())); assert(codeFile.open(QIODevice::ReadOnly));
        const auto raw=codeFile.readAll();
        assert(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex()==module["sha256"].toString().toLatin1().toLower());
        const auto code=linaccess::decodeVbaSource(raw);
        const auto handler=QRegularExpression("Sub\\s+Form_Load\\s*\\(\\s*\\)(.*?)End Sub",QRegularExpression::CaseInsensitiveOption|QRegularExpression::DotMatchesEverythingOption).match(code);
        assert(handler.hasMatch());
        const auto name=module["name"].toString().mid(5);
        QByteArray definition;
        for (const auto& candidate : objects) {
            const auto form=candidate.toObject();
            if (form["kind"]!="form" || form["name"]!=name) continue;
            QFile input(directory.filePath(form["file"].toString())); assert(input.open(QIODevice::ReadOnly)); definition=input.readAll();
            assert(QCryptographicHash::hash(definition,QCryptographicHash::Sha256).toHex()==form["sha256"].toString().toLatin1().toLower());
        }
        assert(!definition.isEmpty());
        auto db=QSqlDatabase::addDatabase("QSQLITE","original-load-controls"); db.setDatabaseName(":memory:"); assert(db.open());
        AccessFormPreview preview(name,definition,raw,db);
        const auto loaded=preview.eventResult("OnLoad");
        if (!loaded || loaded->status!=linaccess::VbaExecutionStatus::Completed)
            std::cerr<<"Original initializer: bound="<<static_cast<bool>(loaded)<<", status="<<(loaded ? static_cast<int>(loaded->status) : -1)<<", error="<<(loaded ? loaded->errorNumber : 0)<<", line="<<(loaded ? loaded->line : 0)<<'\n';
        assert(loaded && loaded->status==linaccess::VbaExecutionStatus::Completed);
        const auto controls=makePreviewFormObject(preview)->member("Controls",{}).object;
        const QRegularExpression assignment("Me[.!](?:\\[([^\\]]+)\\]|([\\p{L}\\p{N}_]+))\\.(Enabled|Visible)\\s*=\\s*(True|False)",QRegularExpression::CaseInsensitiveOption);
        const auto verify=[&](bool reset) {
            int count=0;
            auto matches=assignment.globalMatch(handler.captured(1));
            while (matches.hasNext()) {
                const auto match=matches.next();
                const auto controlName=match.captured(1).isEmpty() ? match.captured(2) : match.captured(1);
                const auto control=controls->member("Item",{linaccess::VbaValue::string(controlName.toUtf8().toStdString())}).object;
                const auto property=match.captured(3).toUtf8().toStdString();
                const bool expected=match.captured(4).compare("True",Qt::CaseInsensitive)==0;
                if (reset) control->assign(property,linaccess::VbaValue::logical(!expected));
                else {
                    const auto actual=control->member(property,{}).boolean;
                    if (actual!=expected) std::cerr<<"Original initializer state mismatch: property="<<property<<", type="<<control->member("ControlType",{}).asNumber()<<", expected="<<expected<<", actual="<<actual<<'\n';
                    assert(actual==expected);
                }
                ++count;
            }
            return count;
        };
        assert(verify(false)==6); assert(verify(true)==6);
        std::vector<linaccess::VbaValue> arguments;
        assert(preview.executeProcedure("Form_Load",arguments).status==linaccess::VbaExecutionStatus::Completed);
        assignments+=verify(false); ++initialized;
        const auto procedures=linaccess::analyzeVba(code.toStdString()).procedures;
        const auto sourceLines=code.split('\n'); QString closeHandler;
        for (std::size_t i=0;i<procedures.size();++i) {
            const auto& procedure=procedures[i];
            const int end=i+1<procedures.size() ? procedures[i+1].line-1 : sourceLines.size();
            const auto body=sourceLines.mid(procedure.line-1,end-procedure.line+1).join('\n');
            if (body.contains(QRegularExpression("^\\s*DoCmd\\.Close\\s*$",QRegularExpression::MultilineOption|QRegularExpression::CaseInsensitiveOption)))
                closeHandler=QString::fromStdString(procedure.name);
        }
        assert(!closeHandler.isEmpty());
        QPushButton* closeButton=nullptr;
        const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(definition));
        for (const auto& control:layout.controls)
            if (control.name+"_Click"==closeHandler && control.onClick=="[Event Procedure]")
                closeButton=preview.findChild<QPushButton*>("access-control-"+control.name);
        assert(closeButton);
        preview.show(); QApplication::processEvents();
        int unexpectedDialogs=0; QTimer guard;
        QObject::connect(&guard,&QTimer::timeout,[&] {
            if (auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) {
                ++unexpectedDialogs; box->accept();
            }
        });
        guard.start(10); closeButton->click(); guard.stop();
        assert(unexpectedDialogs==0 && !preview.isVisible());
        bool closedReference=false;
        try { (void)controls->member("Count",{}); }
        catch(const linaccess::VbaRuntimeError& error) { closedReference=error.number==2450; }
        assert(closedReference);
    }
    assert(initialized==1 && assignments==6);
    std::cout<<"Original Form_Load control initializer executed: "<<initialized<<", verified property assignments: "<<assignments<<'\n';
}
