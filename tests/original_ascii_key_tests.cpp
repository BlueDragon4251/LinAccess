#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_source_codec.hpp"
#include <QApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QLineEdit>
#include <QRegularExpression>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <cmath>

void testOriginalAsciiKeyPress(const QDir& directory,const QJsonArray& objects) {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","original-ascii-keypress"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const auto options=QRegularExpression::CaseInsensitiveOption|QRegularExpression::MultilineOption|QRegularExpression::DotMatchesEverythingOption;
    const QRegularExpression procedures(R"rx(^\s*(?:(?:Private|Public|Friend)\s+)?Sub\s+([^\s(]+)\s*\(([^)]*)\)\s*\r?\n(.*?)^\s*End\s+Sub\b)rx",options);
    const QRegularExpression ascCall(R"(\bAsc\s*\()",QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression fields(R"(\bMe[.!](?:\[([^]]+)\]|([\p{L}\p{N}_]+)))",QRegularExpression::CaseInsensitiveOption);
    const auto quoted=[](QString value) { value.replace('"',"\"\""); return '"'+value+'"'; };
    int tested=0,directCalls=0,keyEvents=0;
    for(const auto& value:objects) {
        const auto module=value.toObject(); if(module.value("kind")!="vba") continue;
        QFile source(directory.filePath(module.value("file").toString())); assert(source.open(QIODevice::ReadOnly));
        const auto bytes=source.readAll(); const auto text=linaccess::decodeVbaSource(bytes);
        auto matches=procedures.globalMatch(text);
        while(matches.hasNext()) {
            const auto match=matches.next(); const auto handler=match.captured(1),body=match.captured(3);
            if(!handler.endsWith("_KeyPress",Qt::CaseInsensitive) || !ascCall.match(body).hasMatch()) continue;
            const auto expected=module.value("sha256").toString();
            assert(!expected.isEmpty() && QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==expected.toLatin1());
            QStringList names{"ascii_probe"}; auto references=fields.globalMatch(body);
            while(references.hasNext()) { const auto field=references.next(); const auto name=field.captured(1).isEmpty() ? field.captured(2) : field.captured(1); if(!names.contains(name,Qt::CaseInsensitive)) names<<name; }
            const auto prefix=handler.left(handler.size()-9);
            QString layout="Begin Form\nWidth =7000\nBegin Section\n";
            for(int i=0;i<names.size();++i) {
                layout+="Begin TextBox\nName ="+quoted(names[i])+"\nTop ="+QString::number(i*500)+"\nWidth =5000\nHeight =350\n";
                if(i==0) layout+="EventProcPrefix ="+quoted(prefix)+"\nOnKeyPress =\"[Event Procedure]\"\n";
                layout+="End\n";
            }
            layout+="End\nEnd\n";
            AccessFormPreview form("original ASCII handler",layout.toUtf8(),bytes,db); form.show(); form.activateWindow(); QApplication::processEvents();
            auto model=makePreviewFormObject(form); auto* editor=form.findChild<QLineEdit*>("access-control-ascii_probe"); assert(editor);
            const auto reset=[&] { for(const auto& name:names) model->assignItem(name.toStdString(),VbaValue::string("")); };
            std::vector<std::vector<VbaValue>> expectedValues(128); std::vector<int> keyCodes(128);
            for(int ascii=0;ascii<128;++ascii) {
                reset(); std::vector<VbaValue> arguments{VbaValue::integer(static_cast<std::int16_t>(ascii))};
                const auto result=form.executeProcedure(handler,arguments); assert(result.status==linaccess::VbaExecutionStatus::Completed);
                assert(arguments.size()==1 && arguments[0].kind==VbaValue::Kind::Number);
                const auto code=arguments[0].number; assert(std::isfinite(code) && code==std::floor(code) && code>=0 && code<128);
                keyCodes[ascii]=static_cast<int>(code);
                for(int i=1;i<names.size();++i) expectedValues[ascii].push_back(model->item(names[i].toStdString()));
                ++directCalls;
            }
            for(int ascii=32;ascii<127;++ascii) {
                reset(); editor->setFocus(); QTest::keyClicks(editor,QString(QChar(ascii)));
                const auto event=form.eventResult("control:ascii_probe:keypress"); assert(event && event->status==linaccess::VbaExecutionStatus::Completed); ++keyEvents;
                for(int i=1;i<names.size();++i) {
                    const auto actual=model->item(names[i].toStdString()); const auto& expected=expectedValues[ascii][i-1];
                    assert(actual.kind==expected.kind && actual.asText()==expected.asText());
                }
                const auto code=keyCodes[ascii];
                if(code==0 || (code>=32 && code<127)) assert(editor->text()==(code==0 ? QString{} : QString(QChar(code))));
            }
            assert(form.close()); ++tested;
        }
    }
    assert(tested==3 && directCalls==384 && keyEvents==285);
    assert(sql.exec("SELECT COUNT(*) FROM _la_access_artifacts") && sql.next() && sql.value(0).toInt()==0); sql.finish();
}
