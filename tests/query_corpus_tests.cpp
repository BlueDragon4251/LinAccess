#include "access_query.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include <QCoreApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlRecord>
#include <QSet>
#include <QRegularExpression>
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
    const QSet<QString> derivedQueryHashes{
        "4a8db1b0b4f74446d4d95098591dbb4f36c871de05533f3c182b23bfa6127e6e",
        "dd53bc871702a225285fd9c1c7bf66b4e95414f457aae401e3e6a442de5e1caf",
        "0c05875192b34a6ae31d0929227e8e86eab3bc0bb0d1785e7bf79c4ccbbe3044",
        "8c50591d9badb66b4842c5df4a84384fb7d9f7cb9fa724ab788ad61021100255",
        "223132c09574d3477409de4a557540e3f2b85dddf7d9f7fcf7c419b1fe39632f",
        "2c2c6d97a6bd3168396880060c905c2548364de2408dfa82aa23428585f631e5"};
    int valueDeclarations=0,executedDerivedQueries=0,sourceColumnMismatches=0;
    for (const auto& object : definitions) {
        QString status;
        bool executed = false; int errorNumber=0;
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
            if (!plan.parameters.empty()) {
                auto definition=database->member("QueryDefs",{linaccess::VbaValue::string(object["name"].toString().toStdString())}).object;
                for (const auto& parameter : plan.parameters) definition->member("Parameters",{linaccess::VbaValue::string(parameter.name.toStdString())}).object->assign("Value",linaccess::VbaValue::numeric(2004));
                int expectedCount=-1;
                if (derivedQueryHashes.contains(object["sha256"].toString()) && object["sha256"].toString()!="4a8db1b0b4f74446d4d95098591dbb4f36c871de05533f3c182b23bfa6127e6e") {
                    const auto where=plan.sql.lastIndexOf("WHERE",-1,Qt::CaseInsensitive); assert(where>=0);
                    auto predicate=plan.sql.mid(where); predicate.remove('('); predicate.remove(')');
                    const auto left=QRegularExpression(R"filter(:la_parameter_\d+\s*=\s*(?:"((?:[^"]|"")+)"|([\p{L}_][\p{L}\p{N}_]*)))filter").match(predicate);
                    const auto right=QRegularExpression(R"filter((?:"((?:[^"]|"")+)"|([\p{L}_][\p{L}\p{N}_]*))\s*=\s*:la_parameter_\d+)filter").match(predicate);
                    assert((left.hasMatch() || right.hasMatch()) && plan.parameters.size()==1);
                    const auto filter=left.hasMatch() ? left : right;
                    auto field=filter.captured(1).isEmpty() ? filter.captured(2) : filter.captured(1); field.replace("\"\"","\"");
                    linaccess::AccessQueryPlan unfiltered; unfiltered.sql=plan.sql.left(where);
                    const auto records=linaccess::executeAccessQueryPlan(db,unfiltered,{});
                    int column=-1;
                    for (int i=0;i<records.columns.size();++i) if (records.columns[i].compare(field,Qt::CaseInsensitive)==0) { assert(column<0); column=i; }
                    assert(column>=0);
                    QVariant input;
                    for (const auto& row : records.rows) if (!row[column].isNull()) { input=row[column]; break; }
                    assert(!input.isNull());
                    expectedCount=0;
                    for (const auto& row : records.rows) if (row[column]==input) ++expectedCount;
                    assert(expectedCount>0);
                    definition->member("Parameters",{linaccess::VbaValue::string(plan.parameters.front().name.toStdString())}).object->assign("Value",linaccess::qtFieldValue(input));
                }
                auto rows=definition->member("OpenRecordset",{linaccess::VbaValue::numeric(4)}).object;
                const auto actualCount=rows->member("RecordCount",{}).asNumber();
                assert(actualCount>=0 && (expectedCount<0 || actualCount==expectedCount)); rows->member("Close",{});
                executed=true; status="executed_with_test_parameters_no_access_reference_oracle";
            }
            else if (!plan.sql.startsWith("SELECT ",Qt::CaseInsensitive)) status = "action_query_unexecuted";
            else {
                auto definition = database->member("QueryDefs",{linaccess::VbaValue::string(object["name"].toString().toUtf8().toStdString())});
                auto rows = definition.object->member("OpenRecordset",{linaccess::VbaValue::numeric(4)});
                assert(rows.object->member("RecordCount",{}).asNumber() >= 0);
                rows.object->member("Close",{});
                executed = true; status = "executed_no_access_reference_oracle";
            }
        } catch (const linaccess::VbaRuntimeError& error) {
            errorNumber=error.number;
            status = error.number == 3085 ? "unsupported_function_or_compilation"
                : error.number == 3078 ? "missing_query_dependency" : error.number == 3075 ? "unsupported_sql_grammar"
                : error.number == 3035 ? "execution_budget_exceeded" : "unsupported_jet_or_binding";
            if (object["sha256"].toString()=="4a8db1b0b4f74446d4d95098591dbb4f36c871de05533f3c182b23bfa6127e6e") {
                assert(error.number==3146);
                const auto plan=linaccess::compileAccessQuery(linaccess::accessQuerySource(db,object["name"].toString()));
                QSqlQuery probe(db); assert(!probe.prepare(plan.sql));
                const auto message=probe.lastError().text();
                const auto offset=message.indexOf("no such column:",0,Qt::CaseInsensitive); assert(offset>=0);
                const auto missing=message.mid(offset+15).section(" Unable",0,0).trimmed();
                assert(missing.count('.')==1);
                const auto record=db.record(missing.section('.',0,0));
                assert(record.count()==6 && record.indexOf(missing.section('.',-1))<0);
                ++sourceColumnMismatches; status="unresolved_original_source_column";
            }
        }
          catch (const std::exception&) {
            status = "sqlite_execution_failed";
        }
        assert(db.rollback());
        if (executed && derivedQueryHashes.contains(object["sha256"].toString())) ++executedDerivedQueries;
        ++counts[status];
        report.append(QJsonObject{{"artifact_sha256",object["sha256"]},{"status",status},{"executed",executed},{"error_number",errorNumber}});
    }
    assert(report.size() == 100);
    assert(valueDeclarations==6);
    assert(executedDerivedQueries==5 && sourceColumnMismatches==1);
    assert(counts["executed_no_access_reference_oracle"] > 0);
    QFile output(QString::fromLocal8Bit(argv[3])); assert(output.open(QIODevice::WriteOnly));
    assert(output.write(QJsonDocument(report).toJson()) > 0);
    for (auto it = counts.cbegin(); it != counts.cend(); ++it) std::cout << it.key().toStdString() << ": " << it.value() << '\n';
    std::cout << "100 original query artifacts classified on isolated copy; no complete Jet compatibility claim\n";
}
