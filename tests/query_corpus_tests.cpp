#include "access_query.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include <QCoreApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlQuery>
#include <QTemporaryDir>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc,char** argv) {
    QCoreApplication application(argc,argv);
    assert(argc == 4);
    QTemporaryDir temporary; assert(temporary.isValid());
    auto db = QSqlDatabase::addDatabase("QSQLITE","query-corpus");
    db.setDatabaseName(temporary.filePath("copy.linaccess")); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(linaccess::importAccessDatabase(db,QString::fromLocal8Bit(argv[2])) == 29);
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QDir root(QString::fromLocal8Bit(argv[1]));
    QFile manifest(root.filePath("manifest.json")); assert(manifest.open(QIODevice::ReadOnly));
    const auto objects = QJsonDocument::fromJson(manifest.readAll()).object()["objects"].toArray();
    QVector<QJsonObject> definitions;
    for (const auto& value : objects) {
        const auto object = value.toObject();
        if (object["kind"] != "query") continue;
        QFile artifact(root.filePath(object["file"].toString())); assert(artifact.open(QIODevice::ReadOnly));
        const auto bytes = artifact.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex() == object["sha256"].toString().toLatin1().toLower());
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES ('query',?,?)");
        query.addBindValue(object["name"].toString()); query.addBindValue(bytes); assert(query.exec());
        definitions.append(object);
    }
    assert(definitions.size() == 100);
    const int views = linaccess::installAccessQueryViews(db);
    assert(views > 0);
    auto database = linaccess::makeDaoDatabase(db);
    QMap<QString,int> counts;
    QJsonArray report;
    int valueDeclarations=0;
    for (const auto& object : definitions) {
        QString status;
        bool executed = false;
        assert(db.transaction());
        try {
            const auto source = linaccess::accessQuerySource(db,object["name"].toString());
            if (source.trimmed().startsWith("PARAMETERS",Qt::CaseInsensitive)) {
                const auto headerEnd=source.indexOf(';'); assert(headerEnd>=0);
                const auto declarations=linaccess::compileAccessQuery(source.left(headerEnd+1)+" SELECT 1");
                bool hasValue=false;
                for (const auto& parameter : declarations.parameters)
                    if (parameter.type.compare("value",Qt::CaseInsensitive)==0) hasValue=true;
                if (hasValue) ++valueDeclarations;
            }
            const auto plan = linaccess::compileAccessQuery(source);
            if (!plan.parameters.empty()) status = "parameters_required_unexecuted";
            else if (!plan.sql.startsWith("SELECT ",Qt::CaseInsensitive)) status = "action_query_unexecuted";
            else {
                auto definition = database->member("QueryDefs",{linaccess::VbaValue::string(object["name"].toString().toUtf8().toStdString())});
                auto rows = definition.object->member("OpenRecordset",{linaccess::VbaValue::numeric(4)});
                assert(rows.object->member("RecordCount",{}).asNumber() >= 0);
                rows.object->member("Close",{});
                executed = true; status = "executed_no_access_reference_oracle";
            }
        } catch (const linaccess::VbaRuntimeError& error) {
            status = error.number == 3085 ? "unsupported_function_or_compilation"
                : error.number == 3078 ? "missing_query_dependency" : error.number == 3075 ? "unsupported_sql_grammar"
                : error.number == 3035 ? "execution_budget_exceeded" : "unsupported_jet_or_binding";
        }
          catch (const std::exception&) { status = "sqlite_execution_failed"; }
        assert(db.rollback());
        ++counts[status];
        report.append(QJsonObject{{"artifact_sha256",object["sha256"]},{"status",status},{"executed",executed}});
    }
    assert(report.size() == 100);
    assert(valueDeclarations==6);
    assert(counts["executed_no_access_reference_oracle"] > 0);
    QFile output(QString::fromLocal8Bit(argv[3])); assert(output.open(QIODevice::WriteOnly));
    assert(output.write(QJsonDocument(report).toJson()) > 0);
    for (auto it = counts.cbegin(); it != counts.cend(); ++it) std::cout << it.key().toStdString() << ": " << it.value() << '\n';
    std::cout << "100 original query artifacts classified on isolated copy; no complete Jet compatibility claim\n";
}
