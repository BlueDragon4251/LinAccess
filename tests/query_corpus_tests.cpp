#include "access_query.hpp"
#include "access_catalog.hpp"
#include "vba_source_codec.hpp"
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

namespace {
void testPreparedOriginalJoinedUpdate(const QDir& root,const QVector<QJsonObject>& definitions,const QString& sourceFile) {
    using namespace linaccess;
    QJsonObject artifact;
    for (const auto& object:definitions) if (object["sha256"].toString()=="5dcb8fb9cfb7e5eb47d6e82989adc3ce77b8574f0481d309764a9829894dc1a9") artifact=object;
    assert(!artifact.isEmpty()); QFile file(root.filePath(artifact["file"].toString())); assert(file.open(QIODevice::ReadOnly)); const auto raw=file.readAll();
    assert(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
    const auto source=decodeVbaSource(raw);
    const QString name="(?:\\[[^\\]]+\\]|[\\p{L}_][\\p{L}\\p{N}_]*)";
    const QRegularExpression shape("^\\s*UPDATE\\s+("+name+")\\s+INNER\\s+JOIN\\s+("+name+")\\s+ON\\s+(.*?)\\s+SET\\s+("+name+")\\s*\\.\\s*("+name+")\\s*=\\s*("+name+")\\s*\\.\\s*("+name+")\\s*;?\\s*$",QRegularExpression::CaseInsensitiveOption|QRegularExpression::DotMatchesEverythingOption);
    const auto match=shape.match(source); assert(match.hasMatch());
    const auto unquote=[](QString value) { return value.startsWith('[') ? value.mid(1,value.size()-2).replace("]]","]") : value; };
    const auto target=unquote(match.captured(1)),other=unquote(match.captured(2));
    assert(unquote(match.captured(4)).compare(target,Qt::CaseInsensitive)==0 && unquote(match.captured(6)).compare(other,Qt::CaseInsensitive)==0);
    const QRegularExpression reference("("+name+")\\s*\\.\\s*("+name+")");
    auto refs=reference.globalMatch(match.captured(3)); QMap<QString,QString> keys;
    while (refs.hasNext()) { const auto item=refs.next(); keys.insert(unquote(item.captured(1)).toCaseFolded(),unquote(item.captured(2))); }
    assert(keys.size()==2 && keys.contains(target.toCaseFolded()) && keys.contains(other.toCaseFolded()));
    const auto destination=unquote(match.captured(5)),input=unquote(match.captured(7));
    auto db=QSqlDatabase::addDatabase("QSQLITE","prepared-original-joined-update"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    AccessCatalog catalog(sourceFile);
    for (const auto& table:QStringList{target,other}) {
        const auto schema=parseMdbSchema(catalog.tableSchema(table)); assert(!schema.columns.isEmpty());
        const bool left=table==target; std::vector<std::vector<CsvField>> rows(2);
        for (const auto& column:schema.columns) {
            rows[0].push_back({column.name.toStdString(),true});
            const bool key=column.name.compare(keys.value(table.toCaseFolded()),Qt::CaseInsensitive)==0;
            QString value;
            if (column.sqliteType=="BLOB") { assert(!column.required); rows[1].push_back({"\x1e",false}); continue; }
            if (column.sourceType=="BOOLEAN" || column.sourceType=="BIT") value=key || left ? "0" : "-1";
            else if (column.sourceType=="DATETIME" || column.sourceType=="DATE") value=key || left ? "2026-01-01 00:00:00" : "2026-01-02 00:00:00";
            else if (column.sqliteType=="INTEGER") value=key ? "7" : left ? "11" : "22";
            else if (column.sqliteType=="REAL") value=key ? "7.5" : left ? "1.25" : "2.75";
            else value=key ? "shared" : left ? "synthetic target" : "synthetic source";
            if (left && column.name.compare(destination,Qt::CaseInsensitive)==0) assert(!column.autoNumber);
            rows[1].push_back({value.toStdString(),true});
        }
        importAccessTable(db,"synthetic",table,schema,rows);
    }
    const auto quote=[](QString value) { return '"'+value.replace('"',"\"\"")+'"'; };
    const auto cell=[&](const QString& table,const QString& field) { QSqlQuery read(db); assert(read.exec("SELECT "+quote(field)+" FROM "+quote(table)) && read.next()); const auto value=read.value(0); assert(!read.next()); return value; };
    const auto expected=cell(other,input),before=cell(target,destination); assert(before!=expected);
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes)VALUES('query',?,?)"); sql.addBindValue(artifact["name"].toString()); sql.addBindValue(raw); assert(sql.exec());
    auto definition=makeDaoDatabase(db)->member("QueryDefs",{VbaValue::string(artifact["name"].toString().toStdString())}).object;
    definition->member("Execute",{VbaValue::numeric(128)});
    assert(definition->member("RecordsAffected",{}).asNumber()==1 && cell(target,destination)==expected);
    std::cout<<"Original INNER JOIN UPDATE: unchanged source, actual two table schemas, one prepared synthetic target changed to independently read source value\n";
}
}

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
    const QSet<QString> conditionalAliasHashes{
        "aee3c2eb8334d24f8268bda68943e5c59c5a993da65d96ea824174da55ed16bf",
        "c603f4a1eb63c4c0ad7a7427a87ae0978485d1bbea492b1caebedb387ad52c00",
        "e074e74467212cd73520402ffb4239aef220557883d048a2be672de10c3d48de"};
    const QSet<QString> sourceColumnMismatchHashes{
        "4a8db1b0b4f74446d4d95098591dbb4f36c871de05533f3c182b23bfa6127e6e",
        "53bc42a528dfbb8ac1867eb721d30e8312c9d7fe5dab0f0ba8c292c3749bbf1e",
        "bf9698243a44a3efc802a079b1525807da18f94b1d2108467858e4aa3c14b325"};
    linaccess::AccessCatalog originalCatalog(QString::fromLocal8Bit(argv[2]));
    int valueDeclarations=0,executedDerivedQueries=0,sourceColumnMismatches=0,positiveConditionalQueries=0;
    for (const auto& object : definitions) {
        QString status,errorReason;
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
                auto parameters=definition->member("Parameters",{}).object;
                assert(parameters->member("Count",{}).asNumber()==plan.parameters.size());
                const auto members=parameters->enumerate(); assert(members.size()==static_cast<std::size_t>(plan.parameters.size()));
                for (int i=0;i<plan.parameters.size();++i) {
                    auto byName=definition->member("Parameters",{linaccess::VbaValue::string(plan.parameters[i].name.toStdString())}).object;
                    assert(members[i].object==byName);
                    assert(parameters->member("Item",{linaccess::VbaValue::numeric(i)}).object==byName);
                    assert(byName->member("Name",{}).asText()==plan.parameters[i].name.toStdString());
                    byName->assign("Value",linaccess::VbaValue::numeric(2004));
                }
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
            else if (!plan.sql.startsWith("SELECT ",Qt::CaseInsensitive)) {
                if (plan.updateTarget.isEmpty()) status="action_query_unexecuted";
                else {
                    auto target=plan.updateTarget; target.replace('"',"\"\"");
                    const auto snapshot=[&] {
                        QSqlQuery read(db); assert(read.exec("SELECT * FROM \""+target+"\" ORDER BY rowid"));
                        QVector<QVariantList> rows;
                        while (read.next()) { QVariantList row; for (int i=0;i<read.record().count();++i) row.append(read.value(i)); rows.append(std::move(row)); }
                        assert(!read.lastError().isValid()); return rows;
                    };
                    const auto before=snapshot();
                    auto definition=database->member("QueryDefs",{linaccess::VbaValue::string(object["name"].toString().toStdString())}).object;
                    definition->member("Execute",{linaccess::VbaValue::numeric(128)});
                    const auto affected=definition->member("RecordsAffected",{}).asNumber(); const auto after=snapshot();
                    assert(after.size()==before.size() && affected>=0);
                    if (!affected) assert(after==before);
                    executed=true; status="executed_joined_action_no_access_reference_oracle";
                }
            }
            else {
                auto definition = database->member("QueryDefs",{linaccess::VbaValue::string(object["name"].toString().toUtf8().toStdString())});
                auto rows = definition.object->member("OpenRecordset",{linaccess::VbaValue::numeric(4)});
                const auto count=rows.object->member("RecordCount",{}).asNumber(); assert(count>=0);
                if (conditionalAliasHashes.contains(object["sha256"].toString())) {
                    assert(count>0);
                    bool nonNull=false;
                    for (const auto& field:rows.object->member("Fields",{}).object->enumerate())
                        if (field.object->member("Value",{}).kind!=linaccess::VbaValue::Kind::Null) nonNull=true;
                    assert(nonNull); ++positiveConditionalQueries;
                }
                rows.object->member("Close",{});
                executed = true; status = "executed_no_access_reference_oracle";
            }
        } catch (const linaccess::VbaRuntimeError& error) {
            errorNumber=error.number;
            const std::string explanation=error.what();
            if(explanation=="Joined UPDATE has conflicting source values for one target") errorReason="joined_update_source_conflict";
            else if(explanation=="Joined UPDATE shape unsupported") errorReason="joined_update_shape";
            else if(explanation=="Joined UPDATE nullable-side target is not implemented") errorReason="joined_update_nullable_target";
            else if(explanation=="Joined UPDATE requires a writable table target") errorReason="joined_update_target";
            else if(explanation=="Joined UPDATE row identity unavailable") errorReason="joined_update_identity";
            if (object["sha256"].toString()=="5dcb8fb9cfb7e5eb47d6e82989adc3ce77b8574f0481d309764a9829894dc1a9") {
                const std::string detail=error.what();
                const auto category=detail=="Joined UPDATE shape unsupported" ? "compiler_shape"
                    : detail=="Joined UPDATE requires a writable table target" ? "target_kind"
                    : detail=="Joined UPDATE has conflicting source values for one target" ? "source_conflict"
                    : detail=="Jet SQL function argument invalid" ? "runtime_function"
                    : detail=="Joined UPDATE row identity unavailable" ? "row_identity" : "other";
                std::cout<<"Original joined UPDATE boundary: error="<<error.number<<", stage="<<category<<'\n';
            }
            status = error.number == 3085 ? "unsupported_function_or_compilation"
                : error.number == 3078 ? "missing_query_dependency" : error.number == 3075 ? "unsupported_sql_grammar"
                : error.number == 3035 ? "execution_budget_exceeded" : "unsupported_jet_or_binding";
            if (sourceColumnMismatchHashes.contains(object["sha256"].toString())) {
                assert(error.number==3146);
                const auto plan=linaccess::compileAccessQuery(linaccess::accessQuerySource(db,object["name"].toString()));
                QSqlQuery probe(db); assert(!probe.prepare(plan.sql));
                const auto message=probe.lastError().text();
                const auto offset=message.indexOf("no such column:",0,Qt::CaseInsensitive); assert(offset>=0);
                const auto missing=message.mid(offset+15).section(" Unable",0,0).trimmed();
                assert(missing.count('.')==1);
                const auto owner=missing.section('.',0,0),field=missing.section('.',-1);
                const auto record=db.record(owner); assert(record.count()>0 && record.indexOf(field)<0);
                QSqlQuery catalog(db); assert(catalog.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='query' AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1"));
                catalog.addBindValue(owner); assert(catalog.exec());
                if(catalog.next()) {
                    const auto originalSource=linaccess::decodeVbaSource(catalog.value(0).toByteArray());
                    const QRegularExpression reference("(?<![\\p{L}\\p{N}_])"+QRegularExpression::escape(field)+"(?![\\p{L}\\p{N}_])",QRegularExpression::CaseInsensitiveOption);
                    assert(reference.isValid() && !reference.match(originalSource).hasMatch());
                    errorReason="original_saved_query_omits_column";
                } else {
                    const auto schema=linaccess::parseMdbSchema(originalCatalog.tableSchema(owner));
                    assert(!schema.columns.isEmpty() && schema.columns.size()==record.count());
                    for(const auto& column:schema.columns) assert(column.name.compare(field,Qt::CaseInsensitive)!=0);
                    errorReason="original_table_omits_column";
                }
                ++sourceColumnMismatches; status="unresolved_original_source_column";
            }
        }
          catch (const std::exception&) {
            status = "sqlite_execution_failed";
        }
        assert(db.rollback());
        if(object["sha256"].toString()=="32b25d06328e208b7d2e94ec33a17422fca977ef7b5988e6b040123562c890ec") assert(executed && errorNumber==0);
        if (executed && derivedQueryHashes.contains(object["sha256"].toString())) ++executedDerivedQueries;
        ++counts[status];
        report.append(QJsonObject{{"artifact_sha256",object["sha256"]},{"status",status},{"executed",executed},{"error_number",errorNumber},{"error_reason",errorReason}});
    }
    assert(report.size() == 100);
    testPreparedOriginalJoinedUpdate(root,definitions,QString::fromLocal8Bit(argv[2]));
    assert(valueDeclarations==6 && positiveConditionalQueries==3);
    assert(executedDerivedQueries==5 && sourceColumnMismatches==3);
    assert(counts["executed_no_access_reference_oracle"]>=72
        && counts["executed_with_test_parameters_no_access_reference_oracle"]>=5);
    QFile output(QString::fromLocal8Bit(argv[3])); assert(output.open(QIODevice::WriteOnly));
    assert(output.write(QJsonDocument(report).toJson()) > 0);
    for (auto it = counts.cbegin(); it != counts.cend(); ++it) std::cout << it.key().toStdString() << ": " << it.value() << '\n';
    std::cout << "100 original query artifacts classified on isolated copy; no complete Jet compatibility claim\n";
}
