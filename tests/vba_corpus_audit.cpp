#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include "vba_source_codec.hpp"
#include <QCoreApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QMap>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
int main(int argc, char** argv) {
    QCoreApplication application(argc, argv);
    assert(argc == 3);
    const QDir root(QString::fromLocal8Bit(argv[1]));
    QFile manifest(root.filePath("manifest.json")); assert(manifest.open(QIODevice::ReadOnly));
    const auto objects = QJsonDocument::fromJson(manifest.readAll()).object()["objects"].toArray();
    assert(objects.size() == 192);
    QJsonArray records;
    QMap<QString,int> categories;
    int modules = 0;
    for (const auto& item : objects) {
        const auto object = item.toObject();
        QFile artifact(root.filePath(object["file"].toString())); assert(artifact.open(QIODevice::ReadOnly));
        const auto bytes = artifact.readAll();
        assert(QCryptographicHash::hash(bytes, QCryptographicHash::Sha256).toHex() == object["sha256"].toString().toLatin1().toLower());
        if (object["kind"] != "vba") continue;
        ++modules;
        const auto source = linaccess::decodeVbaSource(bytes).toUtf8().toStdString();
        const auto analysis = linaccess::analyzeVba(source);
        for (const auto& procedure : analysis.procedures) {
            const auto result = linaccess::validateVbaProcedure(source, procedure.name);
            const auto reason = QString::fromStdString(result.explanation);
            QString category = "syntax_accepted_unexecuted";
            if (result.status != linaccess::VbaExecutionStatus::Completed) {
                if (reason.contains("Parameter") || reason.contains("signatur")) category = "parameters_or_signature";
                else if (reason.contains("If-Bedingung")) category = "conditional_expression";
                else if (reason.contains("Zuweisung")) category = "assignment_expression";
                else if (reason.contains("VBA-Anweisung")) category = "statement";
                else if (reason=="Object method arguments") category="object_method_arguments";
                else if (reason=="MsgBox argument syntax unsupported") category="message_box_arguments";
                else if (reason=="Set-Ausdruck") category="object_set_expression";
                else if (reason=="Variablendeklaration") category="variable_declaration";
                else if (reason=="Objektzuweisung") category="object_assignment_expression";
                else if (reason=="Aufrufargument") category="call_argument_expression";
                else category = "other_syntax";
            }
            ++categories[category];
            // Never serialize explanations: they can contain original literals.
            records.append(QJsonObject{{"artifact_sha256",object["sha256"]},
                {"procedure",QString::fromStdString(procedure.name)}, {"category",category},
                {"line",result.line}, {"executed",false}});
        }
    }
    assert(modules == 32 && records.size() == 187);
    QFile output(QString::fromLocal8Bit(argv[2])); assert(output.open(QIODevice::WriteOnly));
    assert(output.write(QJsonDocument(records).toJson()) > 0);
    for (auto it = categories.cbegin(); it != categories.cend(); ++it)
        std::cout << it.key().toStdString() << ": " << it.value() << '\n';
    std::cout << "32 modules, 187 syntax inspections; no procedures executed by this audit\n";
}
