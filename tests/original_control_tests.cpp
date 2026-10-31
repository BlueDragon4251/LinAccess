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

            ++executed;
        }
    }
    assert(executed==3);
    std::cout << "Original report control handlers executed after resetting visibility: " << executed << '\n';
}
