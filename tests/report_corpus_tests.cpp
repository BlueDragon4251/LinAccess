#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_form_import.hpp"
#include "access_table_import.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include "vba_source_codec.hpp"
#include "vba_qt_value.hpp"
#include <QCoreApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QLineEdit>
#include <QComboBox>
#include <QPushButton>
#include <QMainWindow>
#include <QMenuBar>
#include <QMenu>
#include <QToolBar>
#include <QTimer>
#include <QApplication>
#include <QMessageBox>
#include <QRegularExpression>
#include <QSqlQuery>
#include <QSqlDriver>
#include <QSqlRecord>
#include <QSqlError>
#include <QTemporaryDir>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <memory>
#include <cmath>
namespace {
QString adapterCategory(const linaccess::VbaRuntimeError& error) {
    const auto reason=QString::fromUtf8(error.what()).toCaseFolded();
    if (reason.contains("concatenation")) return "unsupported_concat";
    if (reason.contains("distinctrow")) return "unsupported_distinctrow";
    if (reason.contains("like character")) return "unsupported_like_class";
    return "adapter_error_"+QString::number(error.number);
}
QString failureCategory(const QString& error,const QJsonArray& definitions,const QDir& directory) {
    const auto text=error.toCaseFolded();
    if (const auto number=linaccess::jetSqlFunctionErrorNumber(text.toUtf8().toStdString()))
        return "adapter_error_"+QString::number(*number);
    if (text.contains("no such column")) return "unresolved_column";
    if (text.contains("no such table")) {
        auto name=error.mid(error.toCaseFolded().indexOf("no such table:")+14).trimmed();
        if (name.startsWith("temp.") || name.startsWith("main.")) name=name.mid(5);
        for (const auto& value : definitions) {
            const auto object=value.toObject();
            if (object["kind"]=="query" && object["name"].toString().compare(name,Qt::CaseInsensitive)==0) {
                QFile source(directory.filePath(object["file"].toString())); assert(source.open(QIODevice::ReadOnly));
                try {
                    const auto plan=linaccess::compileAccessQuery(linaccess::decodeVbaSource(source.readAll()));
                    if (!plan.parameters.empty()) return "parameterized_query_dependency";
                    if (plan.hiddenRank) return "ranked_top_query_dependency";
                } catch (const linaccess::VbaRuntimeError& dependency) { return "query_dependency_"+adapterCategory(dependency); }
                return "unsupported_saved_query_dependency";
            }
        }
        return "absent_dependency";
    }
    if (text.contains("no such function")) return "unsupported_function";
    if (text.contains("ambiguous")) return "ambiguous_name";
    if (text.contains("aggregate")) return "aggregate_usage";
    if (text.contains("interrupted")) return "execution_budget";
    if (text.contains("unrecognized token")) {
        const auto match=QRegularExpression("unrecognized token: \"([^\"]*)\"").match(error);
        const auto token=match.captured(1);
        if (token.size()==1 && !token[0].isLetterOrNumber())
            return "unrecognized_token_u"+QString::number(token[0].unicode(),16).rightJustified(4,'0');
        return "unrecognized_token_length_"+QString::number(token.size());
    }
    if (text.contains("syntax")) return "unsupported_grammar";
    return "other_sql_failure";
}
int verifyOriginalFilter(const QDir& directory,const QJsonArray& objects,QSqlDatabase db) {
    // This regression measures original layout/data binding. CodeBehind is
    // deliberately absent; native VBA property dispatch is tested separately.
    for (const auto& value : objects) {
        const auto object=value.toObject(); if (object["kind"]!="form") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(bytes));
        if (layout.filterOnLoad || !db.tables().contains(layout.recordSource,Qt::CaseInsensitive)) continue;
        QSqlQuery source(db); assert(source.exec("SELECT * FROM "+db.driver()->escapeIdentifier(layout.recordSource,QSqlDriver::TableName)));
        const auto record=source.record(); if (record.isEmpty()) continue;
        QVector<QVariantList> rows;
        while (source.next()) {
            QVariantList row;
            for (int i=0;i<record.count();++i) row.append(source.value(i));
            rows.append(row);
        }
        source.finish();
        if (rows.size()<2 || rows.size()>10000 || rows.front().front().isNull()) continue;
        const auto type=rows.front().front().metaType().id();
        if (type!=QMetaType::Int && type!=QMetaType::LongLong) continue;
        const auto input=rows.front().front().toLongLong();
        int expected=0;
        for (const auto& row : rows) if (row.front().toLongLong()==input) ++expected;
        if (expected==rows.size()) continue;
        AccessFormPreview form(object["name"].toString(),bytes,{},db);
        assert(form.hasDataSource());
        auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(selector && selector->count()==rows.size()+1);
        auto field=record.fieldName(0); field.replace(']',"]]");
        auto native=makePreviewFormObject(form);
        native->assign("Filter",linaccess::VbaValue::string(("["+field+"]="+QString::number(input)).toStdString()));
        native->assign("FilterOn",linaccess::VbaValue::logical(true));
        assert(selector->count()==expected+1 && selector->currentIndex()>0);
        int checked=0;
        for (const auto& control : layout.controls) {
            auto name=control.controlSource.trimmed();
            if (name.startsWith('[') && name.endsWith(']')) name=name.mid(1,name.size()-2);
            const int column=record.indexOf(name);
            if (column<0 || control.type!="TextBox") continue;
            auto* widget=form.findChild<QLineEdit*>("access-control-"+control.name); if (!widget) continue;
            assert(widget->property("la_null").toBool()==rows.front()[column].isNull());
            assert(widget->text()==rows.front()[column].toString()); ++checked;
        }
        if (!checked) { assert(form.close()); continue; }
        native->assign("FilterOn",linaccess::VbaValue::logical(false));
        assert(selector->count()==rows.size()+1 && !makePreviewFormObject(form)->member("FilterOn",{}).truthy());
        assert(form.close()); return checked;
    }
    return 0;
}
int verifyOriginalTags(AccessFormPreview& preview,const QString& code,QSqlDatabase db) {
    const auto options=QRegularExpression::CaseInsensitiveOption|QRegularExpression::DotMatchesEverythingOption;
    const auto body=QRegularExpression("Sub\\s+Form_Load\\s*\\(\\s*\\)(.*?)End Sub",options).match(code).captured(1);
    const auto loop=QRegularExpression(R"tag(Me\s*\(\s*"((?:[^"]|"")*)"\s*&\s*(\w+)\s*\)\s*\.Tag\s*=\s*(\w+)\s*!\s*(\w+))tag",QRegularExpression::CaseInsensitiveOption).match(body);
    assert(loop.hasMatch());
    const auto key=QRegularExpression("\\b"+QRegularExpression::escape(loop.captured(2))+"\\s*=\\s*"+QRegularExpression::escape(loop.captured(3))+"!\\s*(\\w+)",QRegularExpression::CaseInsensitiveOption).match(body);
    const auto source=QRegularExpression("Set\\s+"+QRegularExpression::escape(loop.captured(3))+R"sql(\s*=\s*CurrentDb\.OpenRecordset\s*\(\s*"((?:[^"]|"")*)")sql",QRegularExpression::CaseInsensitiveOption).match(body);
    assert(key.hasMatch() && source.hasMatch());
    auto name=source.captured(1); name.replace("\"\"","\"");
    auto prefix=loop.captured(1); prefix.replace("\"\"","\"");
    linaccess::AccessQueryPlan plan;
    QMap<QString,QString> types;
    if (db.tables().contains(name,Qt::CaseInsensitive)) {
        plan.sql="SELECT * FROM "+db.driver()->escapeIdentifier(name,QSqlDriver::TableName);
        types=linaccess::sqliteColumnTypes(db,name);
    } else plan=linaccess::compileAccessQuery(name.startsWith("SELECT",Qt::CaseInsensitive) ? name : linaccess::accessQuerySource(db,name));
    const auto bindings=linaccess::bindAccessQueryParameters(plan,{});
    linaccess::SqliteExecutionBudget budget(db);
    QSqlQuery reference(db); assert(reference.prepare(plan.sql));
    for (auto it=bindings.cbegin();it!=bindings.cend();++it) reference.bindValue(it.key(),it.value());
    assert(reference.exec());
    const auto idColumn=reference.record().indexOf(key.captured(1));
    const auto tagColumn=reference.record().indexOf(loop.captured(4));
    assert(idColumn>=0 && tagColumn>=0);
    auto controls=makePreviewFormObject(preview)->member("Controls",{}).object;
    int count=0;
    while (reference.next()) {
        if (plan.hiddenRank && reference.value(reference.record().count()-1).toLongLong()>plan.topLimit) break;
        const auto id=linaccess::qtFieldValue(reference.value(idColumn),types.value(key.captured(1).toCaseFolded())).asText();
        const auto expected=linaccess::qtFieldValue(reference.value(tagColumn),types.value(loop.captured(4).toCaseFolded())).asText();
        auto control=controls->member("Item",{linaccess::VbaValue::string(prefix.toUtf8().toStdString()+id)}).object;
        assert(control->member("Tag",{}).asText()==expected); ++count;
    }
    assert(!reference.lastError().isValid() && count>0);
    return count;
}
}
void testReportCorpus(const QDir& directory,const QJsonArray& objects,const QString& mdb) {
    QTimer acknowledge;
    QObject::connect(&acknowledge,&QTimer::timeout,[] {
        for (auto* widget : QApplication::topLevelWidgets())
            if (auto* message=qobject_cast<QMessageBox*>(widget)) message->accept();
    });
    acknowledge.start(10);
    QTemporaryDir temporary; assert(temporary.isValid());
    auto db=QSqlDatabase::addDatabase("QSQLITE","report-corpus");
    db.setDatabaseName(temporary.filePath("copy.linaccess")); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(linaccess::importAccessDatabase(db,mdb)==29);
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    for (const auto& value : objects) {
        const auto object=value.toObject();
        if (object["kind"]!="query" && object["kind"]!="form" && object["kind"]!="vba") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES(?,?,?)");
        query.addBindValue(object["kind"].toString()); query.addBindValue(object["name"].toString()); query.addBindValue(bytes); assert(query.exec()); query.finish();
    }
    assert(linaccess::installAccessQueryViews(db)>0);
    assert(db.transaction());
    const auto originalFilterCells=verifyOriginalFilter(directory,objects,db);
    assert(originalFilterCells>0 && db.rollback());
    std::cout<<"Original native filter layout/data cells verified: "<<originalFilterCells<<'\n';
    assert(db.transaction());
    // Resolve references taken from the original SQL against original control
    // definitions. Synthetic inputs exercise lookup, not Access result parity.
    const QRegularExpression refs(R"((?:\[Formulare\]|Formulare|\[Forms\]|Forms)\s*!\s*(?:\[([^\]]+)\]|([\p{L}\p{N}_]+))\s*!\s*(?:\[([^\]]+)\]|([\p{L}\p{N}_]+)))",QRegularExpression::CaseInsensitiveOption);
    int referenceCount=0,referencingQueries=0,mutableReferences=0;
    std::vector<std::unique_ptr<AccessFormPreview>> referencedForms;
    QMap<QString,AccessFormPreview*> openForms;
    QMap<QString,linaccess::AccessFormLayout> referencedLayouts;
    for (const auto& value : objects) {
        const auto object=value.toObject(); if (object["kind"]!="query") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto source=linaccess::decodeVbaSource(file.readAll());
        auto matches=refs.globalMatch(source);
        bool referenced=false;
        while (matches.hasNext()) {
            const auto match=matches.next();
            const auto formName=match.captured(1).isEmpty() ? match.captured(2) : match.captured(1);
            const auto field=match.captured(3).isEmpty() ? match.captured(4) : match.captured(3);
            const auto key=formName.toCaseFolded();
            if (!openForms.contains(key)) {
                QByteArray definition;
                for (const auto& candidate : objects) {
                    const auto form=candidate.toObject();
                    if (form["kind"]!="form" || form["name"].toString().compare(formName,Qt::CaseInsensitive)!=0) continue;
                    QFile input(directory.filePath(form["file"].toString())); assert(input.open(QIODevice::ReadOnly));
                    definition=input.readAll();
                    assert(QCryptographicHash::hash(definition,QCryptographicHash::Sha256).toHex()==form["sha256"].toString().toLatin1().toLower());
                }
                assert(!definition.isEmpty());
                referencedLayouts.insert(key,linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(definition)));
                referencedForms.push_back(std::make_unique<AccessFormPreview>(formName,definition,QByteArray{},db));
                openForms.insert(key,referencedForms.back().get());
            }
            auto control=makePreviewFormObject(*openForms.value(key))->member("Controls",{}).object
                ->member("Item",{linaccess::VbaValue::string(field.toUtf8().toStdString())}).object;
            linaccess::AccessQueryModel lookup;
            const auto initial=control->member("Value",{});
            lookup.execute(db,"SELECT "+match.captured());
            assert(lookup.rowCount()==1);
            const auto expected=linaccess::sqlFieldValue(initial,"");
            const auto actual=lookup.data(lookup.index(0,0));
            assert(actual.isNull()==expected.isNull());
            if (!expected.isNull()) assert(actual.toString()==expected.toString());
            // Query-bound/computed controls correctly reject assignment. Only
            // unbound original controls receive synthetic test inputs.
            const auto& originalLayout=referencedLayouts.value(key);
            for (const auto& item : originalLayout.controls) if (item.name.compare(field,Qt::CaseInsensitive)==0 && item.controlSource.isEmpty()) {
                control->assign("Value",linaccess::VbaValue::numeric(2004));
                lookup.execute(db,"SELECT "+match.captured());
                assert(lookup.data(lookup.index(0,0)).toInt()==2004);
                ++mutableReferences;
            }
            ++referenceCount; referenced=true;
        }
        if (referenced) {
            const auto plan=linaccess::compileAccessQuery(source);
            assert(plan.sql.contains("linaccess_jet_object_value"));
            ++referencingQueries;
        }
    }
    assert(referenceCount==9 && referencingQueries==5);
    // Keep the report inventory independent of these synthetic filter inputs.
    for (auto& form : referencedForms) assert(form->close());
    referencedForms.clear();
    assert(db.rollback());
    std::cout<<"Original SQL object references: "<<referenceCount<<", query definitions: "<<referencingQueries<<", mutable input checks: "<<mutableReferences<<'\n';
    int completedInitializers=0;
    for (const auto& value : objects) {
        const auto module=value.toObject();
        if (module["kind"]!="vba" || module["sha256"].toString().toLatin1().toLower()!="6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8") continue;
        QFile source(directory.filePath(module["file"].toString())); assert(source.open(QIODevice::ReadOnly));
        const auto code=source.readAll();
        assert(QCryptographicHash::hash(code,QCryptographicHash::Sha256).toHex()==module["sha256"].toString().toLatin1().toLower());
        const auto formName=module["name"].toString().mid(5);
        QByteArray definition;
        for (const auto& candidate : objects) {
            const auto form=candidate.toObject(); if (form["kind"]!="form" || form["name"]!=formName) continue;
            QFile file(directory.filePath(form["file"].toString())); assert(file.open(QIODevice::ReadOnly)); definition=file.readAll();
            assert(QCryptographicHash::hash(definition,QCryptographicHash::Sha256).toHex()==form["sha256"].toString().toLatin1().toLower());
        }
        assert(!definition.isEmpty() && db.transaction());
        {
            QMainWindow window;
            auto* toolbar=window.addToolBar("Project"); toolbar->addAction("Test command"); toolbar->setEnabled(false);
            auto* menuBar=window.menuBar(); menuBar->addMenu("Project")->addAction(toolbar->actions().first()); menuBar->setEnabled(false);
            AccessFormPreview initialized(formName,definition,code,db,&window);
            initialized.show(); QApplication::setActiveWindow(&initialized); QApplication::processEvents();
            const auto activated=initialized.eventResult("OnActivate");
            assert(activated && activated->status==linaccess::VbaExecutionStatus::Completed);
            const auto opened=initialized.eventResult("OnOpen");
            std::cerr<<"Original OnOpen menu adapter: status="<<(opened ? static_cast<int>(opened->status) : -1)<<", error="<<(opened ? opened->errorNumber : 0)<<", line="<<(opened ? opened->line : 0)<<", menu enabled="<<menuBar->isEnabled()<<'\n';
            assert(opened && opened->status==linaccess::VbaExecutionStatus::Completed && menuBar->isEnabled());
            const auto result=initialized.eventResult("OnLoad");
            if (!result || result->status!=linaccess::VbaExecutionStatus::Completed)
                std::cerr<<"Complex original initializer: bound="<<static_cast<bool>(result)<<", status="<<(result ? static_cast<int>(result->status) : -1)<<", error="<<(result ? result->errorNumber : 0)<<", line="<<(result ? result->line : 0)<<'\n';
            assert(result && result->status==linaccess::VbaExecutionStatus::Completed);
            assert(toolbar->isEnabled());
            const auto tags=verifyOriginalTags(initialized,linaccess::decodeVbaSource(code),db);
            int subforms=0,calculated=0,calculationErrors=0;
            QJsonArray calculations;
            const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(definition));
            for (qsizetype index=0;index<layout.controls.size();++index) {
                const auto& control=layout.controls[index];
                if (control.type=="Subform") {
                    auto* child=dynamic_cast<AccessFormPreview*>(initialized.findChild<QWidget*>("access-control-"+control.name));
                    if (child && child->hasDataSource()) ++subforms;
                }
                if (control.controlSource.trimmed().startsWith('=')) {
                    const auto value=makePreviewFormObject(initialized)->item(control.name.toUtf8().toStdString());
                    ++calculated; calculationErrors+=value.kind==linaccess::VbaValue::Kind::Error;
                    if (value.kind==linaccess::VbaValue::Kind::Error) std::cerr<<"Calculated control position="<<index<<", error="<<value.errorNumber<<'\n';
                    calculations.append(QJsonObject{{"module_sha256",module["sha256"]},{"control_index",index},
                        {"value_kind",static_cast<int>(value.kind)},{"error_number",value.errorNumber}});
                }
            }
            int originalDomainSums=0;
            const QRegularExpression domainSum(R"domain(^=DSum\(\s*"([^"]+)"\s*[,;]\s*"([^"]+)"(?:\s*[,;]\s*"([^"]*)")?\s*\)$)domain",QRegularExpression::CaseInsensitiveOption);
            for (const auto& control : layout.controls) {
                const auto match=domainSum.match(control.controlSource.trimmed());
                if (!match.hasMatch()) continue;
                // Independently sum the actual projected records in C++; do
                // not reuse the domain adapter's SQLite aggregate result.
                const auto domain=db.driver()->escapeIdentifier(match.captured(2),QSqlDriver::TableName);
                const auto source="SELECT "+match.captured(1)+" FROM "+domain+(match.captured(3).isEmpty() ? QString{} : " WHERE "+match.captured(3));
                const auto plan=linaccess::compileAccessQuery(source);
                const auto rows=linaccess::executeAccessQueryPlan(db,plan,linaccess::bindAccessQueryParameters(plan,{}));
                double expected=0; int nonnull=0;
                for (const auto& row : rows.rows) if (!row.front().isNull()) { expected+=row.front().toDouble(); ++nonnull; }
                const auto actual=makePreviewFormObject(initialized)->item(control.name.toUtf8().toStdString());
                if (nonnull) assert(std::abs(actual.asNumber()-expected)<=1e-8*std::max(1.0,std::abs(expected)));
                else assert(actual.kind==linaccess::VbaValue::Kind::Null);
                ++originalDomainSums;
            }
            assert(originalDomainSums==2);
            const auto current=initialized.eventResult("OnCurrent");
            assert(current && current->status==linaccess::VbaExecutionStatus::Completed);
            assert(calculated==34 && calculationErrors==1);
            int originalCurrencyChecks=0;
            const QRegularExpression currencyInput(R"currency(CCur\(\s*\[([^\]]+)\]\s*\))currency",QRegularExpression::CaseInsensitiveOption);
            for (const auto& control : layout.controls) {
                const auto input=currencyInput.match(control.controlSource);
                if (!input.hasMatch()) continue;
                auto originalForm=makePreviewFormObject(initialized);
                const auto initial=originalForm->item(control.name.toStdString());
                assert(initial.kind==linaccess::VbaValue::Kind::Error && initial.errorNumber==94);
                auto inputControl=originalForm->member("Controls",{linaccess::VbaValue::string(input.captured(1).toStdString())}).object;
                const auto before=inputControl->member("Value",{});
                assert(before.kind==linaccess::VbaValue::Kind::Null);
                inputControl->assign("Value",linaccess::VbaValue::numeric(1.25));
                const auto actual=originalForm->item(control.name.toStdString());
                assert(actual.kind==linaccess::VbaValue::Kind::Currency && actual.currencyUnits==12500);
                // Pass the original expression's actual typed result through
                // a separate native VBA helper; original code bytes stay intact.
                const QByteArray helperCode="Sub Adjust(ByRef amount As Variant)\nCall Increment(amount)\nEnd Sub\nSub Increment(ByRef amount As Variant)\namount=amount+CCur(\"0.0001\")\nEnd Sub\n";
                AccessFormPreview helper("currency value interop","Begin Form\nEnd\n",helperCode,db);
                std::vector<linaccess::VbaValue> currencyArgs{actual};
                assert(helper.executeProcedure("Adjust",currencyArgs).status==linaccess::VbaExecutionStatus::Completed);
                assert(currencyArgs.front().kind==linaccess::VbaValue::Kind::Currency && currencyArgs.front().currencyUnits==12501);
                assert(helper.close());
                inputControl->assign("Value",before);
                ++originalCurrencyChecks;
            }
            assert(originalCurrencyChecks==1);
            std::cout<<"Original domain sums independently verified: "<<originalDomainSums<<"; original Form_Current completed\n";
            assert(subforms>=2 && calculated>0);
            assert(initialized.close());
            QJsonArray events;
            for (auto event=layout.events.cbegin();event!=layout.events.cend();++event) {
                if (event.value().compare("[Event Procedure]",Qt::CaseInsensitive)!=0) continue;
                const auto outcome=initialized.eventResult(event.key());
                events.append(QJsonObject{{"module_sha256",module["sha256"]},{"event",event.key()},
                    {"dispatched",outcome.has_value()},{"status",outcome ? static_cast<int>(outcome->status) : -1},{"error_number",outcome ? outcome->errorNumber : 0},{"line",outcome ? outcome->line : 0}});
            }
            QFile eventInventory(QCoreApplication::applicationDirPath()+"/private-form-event-inventory.json"); assert(eventInventory.open(QIODevice::WriteOnly));
            assert(eventInventory.write(QJsonDocument(events).toJson())>0);
            QFile calculationInventory(QCoreApplication::applicationDirPath()+"/private-form-calculation-inventory.json"); assert(calculationInventory.open(QIODevice::WriteOnly));
            assert(calculationInventory.write(QJsonDocument(calculations).toJson())>0);
            std::cout<<"Complex original Form_Load completed; independently verified control Tags: "<<tags<<", available subforms: "<<subforms<<", calculated controls: "<<calculated<<", calculation error values: "<<calculationErrors<<'\n';
            ++completedInitializers;
        }
        assert(db.rollback());
    }
    assert(completedInitializers==1);
    int scanned=0,available=0,cells=0;
    QJsonArray inventory;
    QMap<QString,int> failures;
    for (const auto& value : objects) {
        const auto object=value.toObject(); if (object["kind"]!="report") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(bytes));
        assert(layout.kind==linaccess::AccessDocumentKind::Report);
        // Source modules are deliberately absent: measure actual record binding,
        // without confusing this structural/data test with VBA execution.
        AccessFormPreview report(object["name"].toString(),bytes,{},db);
        QString status;
        int reportCells=0;
        if (report.hasDataSource()) {
            ++available;
            const auto source=layout.recordSource.trimmed();
            linaccess::AccessQueryPlan plan;
            if (db.tables().contains(source,Qt::CaseInsensitive))
                plan.sql="SELECT * FROM "+db.driver()->escapeIdentifier(source,QSqlDriver::TableName);
            else plan=linaccess::compileAccessQuery(source.startsWith("SELECT",Qt::CaseInsensitive)
                ? source : linaccess::accessQuerySource(db,source));
            const auto bindings=linaccess::bindAccessQueryParameters(plan,{});
            linaccess::SqliteExecutionBudget budget(db);
            QSqlQuery reference(db); assert(reference.prepare(plan.sql));
            for (auto it=bindings.cbegin();it!=bindings.cend();++it) reference.bindValue(it.key(),it.value());
            assert(reference.exec()); const bool hasRow=reference.next();
            assert((report.findChild<QComboBox*>("access-record-selector")->count()>1)==hasRow);
            for (const auto& control : layout.controls) {
                if (control.controlSource.isEmpty() || control.controlSource.startsWith('=')) continue;
                if (auto* line=report.findChild<QLineEdit*>("access-control-"+control.name)) {
                    assert(line->isReadOnly());
                    auto field=control.controlSource.trimmed();
                    if (field.startsWith('[') && field.endsWith(']')) field=field.mid(1,field.size()-2);
                    const int column=reference.record().indexOf(field);
                    if (hasRow && column>=0) {
                        assert(line->text()==reference.value(column).toString());
                        assert(line->property("la_null").toBool()==reference.value(column).isNull());
                    }
                    if (!line->property("la_null").toBool()) ++reportCells;
                }
            }
            cells+=reportCells;
            status=reportCells ? "bound_cells_no_access_reference_oracle" : "source_executed_no_bound_cells";
        } else {
            status="unbound_source";
            try {
                const auto source=layout.recordSource.trimmed();
                const auto sql=source.startsWith("SELECT",Qt::CaseInsensitive) || source.startsWith("PARAMETERS",Qt::CaseInsensitive)
                    ? source : linaccess::accessQuerySource(db,source);
                const auto plan=linaccess::compileAccessQuery(sql);
                if (!plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive)) throw linaccess::VbaRuntimeError(3065,"Report source must be SELECT");
                const auto bindings=linaccess::bindAccessQueryParameters(plan,{});
                linaccess::SqliteExecutionBudget budget(db);
                QSqlQuery diagnostic(db);
                bool ok=diagnostic.prepare(plan.sql);
                if (ok) {
                    for (auto it=bindings.cbegin();it!=bindings.cend();++it) diagnostic.bindValue(it.key(),it.value());
                    ok=diagnostic.exec();
                }
                if (!ok) status=failureCategory(diagnostic.lastError().databaseText(),objects,directory);
                else {
                    while (diagnostic.next()) {}
                    if (diagnostic.lastError().isValid()) status=failureCategory(diagnostic.lastError().databaseText(),objects,directory);
                }
            } catch (const linaccess::VbaRuntimeError& error) { status=adapterCategory(error); }
            ++failures[status];
        }
        auto* save=report.findChild<QPushButton*>("access-save-record"); assert(save && !save->isEnabled());
        inventory.append(QJsonObject{{"artifact_sha256",object["sha256"]},{"status",status},{"bound_cells",reportCells},{"vba_executed",false}});
        assert(report.close()); ++scanned;
    }
    assert(scanned==9);
    QFile output(QCoreApplication::applicationDirPath()+"/private-report-inventory.json"); assert(output.open(QIODevice::WriteOnly));
    assert(output.write(QJsonDocument(inventory).toJson())>0);
    std::cout<<"Original reports: "<<scanned<<", executed sources: "<<available<<", non-NULL bound controls: "<<cells<<'\n';
    for (auto it=failures.cbegin();it!=failures.cend();++it) std::cout<<"Report source category "<<it.key().toStdString()<<": "<<it.value()<<'\n';
    assert(available>0 && cells>0);
}
