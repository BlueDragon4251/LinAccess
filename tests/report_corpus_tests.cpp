#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_form_import.hpp"
#include "access_table_import.hpp"
#include "access_query.hpp"
#include "dao_database.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include "vba_source_codec.hpp"
#include "vba_analyzer.hpp"
#include "vba_qt_value.hpp"
#include <QCoreApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QLineEdit>
#include <QKeyEvent>
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
#include <set>
#include <cmath>
namespace {
class OriginalKeyDelivery final : public QObject {
public:
    explicit OriginalKeyDelivery(QWidget& scope):scope_(scope) { QApplication::instance()->installEventFilter(this); }
    ~OriginalKeyDelivery() override { QApplication::instance()->removeEventFilter(this); }
    int presses{},dialogs{};
    QString deliveredText,beforeReturn;
    bool sawReturn{};
    QPointer<QWidget> firstReceiver;
    bool initialWholeSelection{};
protected:
    bool eventFilter(QObject* object,QEvent* event) override {
        if (event->type()==QEvent::Show && qobject_cast<QMessageBox*>(object)) ++dialogs;
        if (event->type()==QEvent::KeyPress)
            if (auto* widget=qobject_cast<QWidget*>(object);widget && (widget==&scope_ || scope_.isAncestorOf(widget))) {
                ++presses;
                if (!firstReceiver) {
                    firstReceiver=widget;
                    if (auto* editor=qobject_cast<QLineEdit*>(widget)) initialWholeSelection=editor->selectedText()==editor->text();
                }
                const auto* key=static_cast<QKeyEvent*>(event);
                if (!sawReturn && (key->key()==Qt::Key_Return || key->key()==Qt::Key_Enter))
                    if (auto* editor=qobject_cast<QLineEdit*>(widget)) { beforeReturn=editor->text(); sawReturn=true; }
                deliveredText+=key->text();
            }
        return false;
    }
private:
    QWidget& scope_;
};
QWidget* controlInDocument(AccessFormPreview& document,const QString& canonical) {
    QWidget* found=nullptr;
    for (auto* widget:document.findChildren<QWidget*>("access-control-"+canonical)) {
        auto* owner=widget->parentWidget();
        while (owner && !dynamic_cast<AccessFormPreview*>(owner)) owner=owner->parentWidget();
        if (owner==&document) { assert(!found); found=widget; }
    }
    assert(found); return found;
}
int primeOriginalSearchInputs(const QString& source,QSqlDatabase db,AccessFormPreview& form,const linaccess::AccessFormLayout& layout) {
    const auto name=QString::fromStdString(makePreviewFormObject(form)->member("Name",{}).asText());
    const QRegularExpression references(R"((?:\[Formulare\]|Formulare|\[Forms\]|Forms)\s*!\s*(?:\[([^\]]+)\]|([\p{L}\p{N}_]+))\s*!\s*(?:\[([^\]]+)\]|([\p{L}\p{N}_]+)))",QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression likeOperand(R"rx(\bLIKE\s*\(*\s*(?:(?:'(?:[^']|'')*'|"(?:[^"]|"")*")\s*&\s*)?$)rx",QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression equalsOperand(R"((?:\[([^\]]+)\]|([\p{L}\p{N}_]+))\s*\.\s*(?:\[([^\]]+)\]|([\p{L}\p{N}_]+))\s*\)*\s*=\s*\(*\s*$)",QRegularExpression::CaseInsensitiveOption);
    QStringList queryNames;
    QSqlQuery catalog(db); assert(catalog.exec("SELECT DISTINCT name FROM _la_access_artifacts WHERE kind='query'"));
    while (catalog.next()) queryNames.append(catalog.value(0).toString()); catalog.finish();
    QSet<QString> visited; int inputs=0;
    std::function<void(const QString&,int)> visit=[&](const QString& sql,int depth) {
        assert(depth<24);
        auto matches=references.globalMatch(sql);
        while (matches.hasNext()) {
            const auto reference=matches.next();
            const auto owner=reference.captured(1).isEmpty() ? reference.captured(2) : reference.captured(1);
            const auto field=reference.captured(3).isEmpty() ? reference.captured(4) : reference.captured(3);
            if (owner.compare(name,Qt::CaseInsensitive)!=0) continue;
            const auto prefix=sql.left(reference.capturedStart());
            const bool like=likeOperand.match(prefix).hasMatch(); const auto equals=equalsOperand.match(prefix);
            if (!like && !equals.hasMatch()) continue;
            for (const auto& control:layout.controls) if (control.name.compare(field,Qt::CaseInsensitive)==0 && control.controlSource.isEmpty() && (control.type=="TextBox" || control.type=="ComboBox")) {
                auto value=linaccess::VbaValue::string("*");
                if (!like) {
                    const auto source=equals.captured(1).isEmpty() ? equals.captured(2) : equals.captured(1);
                    const auto column=equals.captured(3).isEmpty() ? equals.captured(4) : equals.captured(3);
                    QSqlQuery actual(db);
                    const auto identifier=db.driver()->escapeIdentifier(column,QSqlDriver::FieldName);
                    assert(actual.exec("SELECT "+identifier+" FROM "+db.driver()->escapeIdentifier(source,QSqlDriver::TableName)+" WHERE "+identifier+" IS NOT NULL LIMIT 1") && actual.next());
                    value=linaccess::qtFieldValue(actual.value(0)); actual.finish();
                }
                makePreviewFormObject(form)->assignItem(control.name.toStdString(),value); ++inputs; break;
            }
        }
        for (const auto& view:queryNames) {
            if (visited.contains(view.toCaseFolded())) continue;
            if (!sql.contains("["+view+"]",Qt::CaseInsensitive) && !sql.contains(QRegularExpression("\\b"+QRegularExpression::escape(view)+"\\b",QRegularExpression::CaseInsensitiveOption))) continue;
            visited.insert(view.toCaseFolded());
            QSqlQuery artifact(db); artifact.prepare("SELECT 1 FROM _la_access_artifacts WHERE kind='query' AND name=? COLLATE NOCASE LIMIT 1"); artifact.addBindValue(view);
            if (artifact.exec() && artifact.next()) visit(linaccess::accessQuerySource(db,view),depth+1);
        }
    };
    visit(source,0); return inputs;
}
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
        QString focusControl;
        for (const auto& control:layout.controls) {
            auto name=control.controlSource.trimmed(); if (name.startsWith('[') && name.endsWith(']')) name=name.mid(1,name.size()-2);
            if (control.type=="TextBox" && control.visible && record.indexOf(name)>=0) { focusControl=control.name; break; }
        }
        if (!checked || focusControl.isEmpty()) { assert(form.close()); continue; }
        native->assign("FilterOn",linaccess::VbaValue::logical(false));
        assert(selector->count()==rows.size()+1 && !makePreviewFormObject(form)->member("FilterOn",{}).truthy());
        assert(form.close());
        // The same original layout and records are opened through real VBA
        // with WhereCondition; CodeBehind stays outside this binding test.
        const auto formName=object["name"].toString();
        QSqlQuery latest(db);
        latest.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba',?,?)");
        latest.addBindValue("Form_"+formName); latest.addBindValue(QByteArray{}); assert(latest.exec()); latest.finish();
        const QByteArray callerCode=R"VBA(Sub OpenOriginal(target As String,criteria As String,controlName As String)
DoCmd.OpenForm target,acNormal,,criteria
If Forms(target).Filter<>criteria Then Err.Raise 5
DoCmd.GoToControl controlName
Me.Tag=Screen.ActiveControl.Name
End Sub
Sub LastOriginal(target As String)
Forms(target).FilterOn=False
DoCmd.GoToRecord , , acLast
DoCmd.Requery
End Sub
)VBA";
        AccessFormPreview caller("original binding caller","Begin Form\nEnd\n",callerCode,db);
        std::vector<linaccess::VbaValue> arguments{linaccess::VbaValue::string(formName.toStdString()),linaccess::VbaValue::string(("["+field+"]="+QString::number(input)).toStdString()),linaccess::VbaValue::string(focusControl.toStdString())};
        assert(caller.executeProcedure("OpenOriginal",arguments).status==linaccess::VbaExecutionStatus::Completed);
        AccessFormPreview* opened=nullptr;
        for (auto* widget:QApplication::topLevelWidgets()) if (auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&caller && candidate->isVisible() && makePreviewFormObject(*candidate)->member("Name",{}).asText()==formName.toStdString()) { assert(!opened); opened=candidate; }
        assert(opened);
        auto* records=opened->findChild<QComboBox*>("access-record-selector"); assert(records && records->count()==expected+1);
        int cells=0;
        for (const auto& control:layout.controls) {
            auto name=control.controlSource.trimmed();
            if (name.startsWith('[') && name.endsWith(']')) name=name.mid(1,name.size()-2);
            const auto column=record.indexOf(name); if (column<0 || control.type!="TextBox") continue;
            auto* widget=opened->findChild<QLineEdit*>("access-control-"+control.name); if (!widget) continue;
            assert(widget->property("la_null").toBool()==rows.front()[column].isNull());
            assert(widget->text()==rows.front()[column].toString()); ++cells;
        }
        assert(cells==checked);
        assert(makePreviewFormObject(caller)->member("Tag",{}).asText()==focusControl.toStdString());
        std::vector<linaccess::VbaValue> last{linaccess::VbaValue::string(formName.toStdString())};
        assert(caller.executeProcedure("LastOriginal",last).status==linaccess::VbaExecutionStatus::Completed);
        assert(records->count()==rows.size()+1 && records->currentIndex()==rows.size());
        int lastCells=0;
        for (const auto& control:layout.controls) {
            auto name=control.controlSource.trimmed(); if (name.startsWith('[') && name.endsWith(']')) name=name.mid(1,name.size()-2);
            const auto column=record.indexOf(name); if (column<0 || control.type!="TextBox") continue;
            auto* widget=opened->findChild<QLineEdit*>("access-control-"+control.name); if (!widget) continue;
            assert(widget->property("la_null").toBool()==rows.back()[column].isNull());
            assert(widget->text()==rows.back()[column].toString()); ++lastCells;
        }
        assert(lastCells==checked);
        std::cout<<"Original active-document GoToControl/Last/Requery: independent bound-cell checks="<<lastCells<<'\n';
        assert(opened->close());
        assert(caller.close()); return checked;
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
void testOriginalCalendar(const QDir&,const QJsonArray&,QSqlDatabase);

void testOriginalKeyboardEvents(const QDir&,const QJsonArray&,QSqlDatabase);
void testOriginalTimer(const QDir&,const QJsonArray&,QSqlDatabase);
void testOriginalTabOrder(const QDir&,const QJsonArray&,QSqlDatabase);
void testOriginalPasswordMask(const QDir&,const QJsonArray&);
void testOriginalControlFonts(const QDir&,const QJsonArray&);
void testOriginalAllowEdits(const QDir&,const QJsonArray&,QSqlDatabase);
void testOriginalControlDefaults(const QDir&,const QJsonArray&);
void testOriginalTextAlignment(const QDir&,const QJsonArray&);
void testOriginalForeColor(const QDir&,const QJsonArray&);
void testOriginalBackgrounds(const QDir&,const QJsonArray&);
void testOriginalOpenReport(const QDir&,const QJsonArray&,QSqlDatabase);

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
        if (object["kind"]!="query" && object["kind"]!="form" && object["kind"]!="report" && object["kind"]!="vba") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES(?,?,?)");
        query.addBindValue(object["kind"].toString()); query.addBindValue(object["name"].toString()); query.addBindValue(bytes); assert(query.exec()); query.finish();
    }
    // Invoke the unchanged public standard-module function through the native
    // project resolver, over every original table and one absent name.
    bool originalTableExists=false;
    for (const auto& entry:objects) {
        const auto artifact=entry.toObject();
        if (artifact["sha256"]!="4be67ece759c7816795451ec075ca3fe4fd87cd5aab632fcaca6ecb534fea5ba") continue;
        QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly));
        const auto bytes=input.readAll(); assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
        const auto analysis=linaccess::analyzeVba(linaccess::decodeVbaSource(bytes).toStdString()); assert(analysis.procedures.size()==1);
        const auto wrapper=QByteArray("Function CheckOriginal(target As String) As Boolean\nCheckOriginal=")+QByteArray::fromStdString(analysis.procedures.front().name)+"(target)\nEnd Function\n";
        AccessFormPreview caller("original TableDefs function","Begin Form\nEnd\n",wrapper,db);
        auto database=linaccess::makeDaoDatabase(db); auto definitions=database->member("TableDefs",{}).object;
        assert(definitions->member("Count",{}).asNumber()==29 && definitions->enumerate().size()==29);
        int checked=0;
        for (const auto& table:db.tables()) {
            if (table.startsWith("_la_") || table.startsWith("sqlite_")) continue;
            std::vector<linaccess::VbaValue> arguments{linaccess::VbaValue::string(table.toStdString())};
            const auto result=caller.executeProcedure("CheckOriginal",arguments);
            assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue);
            assert(result.typedReturnValue->kind==linaccess::VbaValue::Kind::Boolean && result.typedReturnValue->boolean);
            auto definition=database->member("TableDefs",{arguments.front()}).object;
            assert(definition->member("Name",{}).asText()==table.toStdString());
            auto records=definition->member("OpenRecordset",{linaccess::VbaValue::numeric(4)}).object;
            QSqlQuery count(db); assert(count.exec("SELECT COUNT(*) FROM "+db.driver()->escapeIdentifier(table,QSqlDriver::TableName)) && count.next());
            assert(records->member("RecordCount",{}).asNumber()==count.value(0).toLongLong());
            records->member("Close",{}); ++checked;
        }
        assert(checked==29);
        const QString absent="linaccess_missing_table_for_original_function"; assert(!db.tables().contains(absent));
        std::vector<linaccess::VbaValue> missing{linaccess::VbaValue::string(absent.toStdString())};
        const auto result=caller.executeProcedure("CheckOriginal",missing);
        assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->kind==linaccess::VbaValue::Kind::Boolean && !result.typedReturnValue->boolean);
        assert(caller.close()); originalTableExists=true;
    }
    assert(originalTableExists);
    std::cout<<"Original TableDefs function: 29 present and one absent name; record counts independently verified\n";
    // Preserve one originally rejected Unicode-field assignment verbatim,
    // with actual imported values and native DAO records on disposable clones.
    bool originalUnicodeAssignment=false;
    for (const auto& entry:objects) {
        const auto artifact=entry.toObject();
        if (!artifact["sha256"].toString().startsWith("824977e9d7c7")) continue;
        QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly));
        const auto bytes=input.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
        const auto lines=linaccess::decodeVbaSource(bytes).split('\n'); assert(lines.size()>136);
        const auto statement=lines[136].trimmed();
        const QRegularExpression shape(R"(^([\p{L}\p{N}_]+)!([\p{L}\p{N}_]+)\s*=\s*([\p{L}\p{N}_]+)!([\p{L}\p{N}_]+)$)");
        const auto assignment=shape.match(statement); assert(assignment.hasMatch());
        assert(assignment.captured(1)!=assignment.captured(3));
        const auto destinationField=assignment.captured(2),sourceField=assignment.captured(4);
        QVariant expected; QString fieldType;
        for (const auto& table:db.tables()) {
            if (table.startsWith("_la_") || table.startsWith("sqlite_")) continue;
            const auto types=linaccess::sqliteColumnTypes(db,table);
            if (!types.contains(sourceField.toCaseFolded())) continue;
            QSqlQuery source(db);
            const auto column=db.driver()->escapeIdentifier(sourceField,QSqlDriver::FieldName);
            if (!source.exec("SELECT "+column+" FROM "+db.driver()->escapeIdentifier(table,QSqlDriver::TableName)+" WHERE "+column+" IS NOT NULL LIMIT 1") || !source.next()) continue;
            expected=source.value(0); fieldType=types.value(sourceField.toCaseFolded()); break;
        }
        assert(expected.isValid() && !expected.isNull() && !fieldType.isEmpty());
        assert(db.transaction());
        const auto destinationColumn=db.driver()->escapeIdentifier(destinationField,QSqlDriver::FieldName);
        const auto sourceColumn=db.driver()->escapeIdentifier(sourceField,QSqlDriver::FieldName);
        assert(query.exec("CREATE TABLE _la_unicode_source("+sourceColumn+" "+fieldType+")"));
        assert(query.exec("CREATE TABLE _la_unicode_destination("+destinationColumn+" "+fieldType+")"));
        query.prepare("INSERT INTO _la_unicode_source VALUES(?)"); query.addBindValue(expected); assert(query.exec()); query.finish();
        assert(query.exec("INSERT INTO _la_unicode_destination VALUES(NULL)")); query.finish();
        const auto wrapper=("Sub OriginalAssignment("+assignment.captured(1)+" As DAO.Recordset,"+assignment.captured(3)+" As DAO.Recordset)\n"+statement+"\nEnd Sub\n").toUtf8();
        {
            AccessFormPreview caller("original Unicode assignment","Begin Form\nEnd\n",wrapper,db);
            auto database=linaccess::makeDaoDatabase(db);
            auto destination=database->member("OpenRecordset",{linaccess::VbaValue::string("_la_unicode_destination")});
            for (const bool nullSource:{false,true}) {
                if (nullSource) { assert(query.exec("UPDATE _la_unicode_source SET "+sourceColumn+"=NULL")); query.finish(); }
                auto source=database->member("OpenRecordset",{linaccess::VbaValue::string("_la_unicode_source")});
                destination.object->member("Edit",{});
                std::vector<linaccess::VbaValue> arguments{destination,source};
                assert(caller.executeProcedure("OriginalAssignment",arguments).status==linaccess::VbaExecutionStatus::Completed);
                destination.object->member("Update",{});
                assert(query.exec("SELECT "+destinationColumn+" FROM _la_unicode_destination") && query.next());
                assert(nullSource ? query.value(0).isNull() : query.value(0)==expected); query.finish();
                source.object->member("Close",{});
            }
            destination.object->member("Close",{}); assert(caller.close());
        }
        assert(db.rollback()); originalUnicodeAssignment=true;
    }
    assert(originalUnicodeAssignment);
    std::cout<<"Original Unicode DAO assignment: non-Null and Null values independently verified; full handler not executed\n";
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
            // Run the unchanged bracket condition from the original source
            // against controls from its unchanged layout. The surrounding
            // synthetic probe is deliberately not counted as its full handler.
            const QRegularExpression bracketCondition(R"(If\s+(\[([^\]]+)\]\s*=\s*-1)\s+Then)",QRegularExpression::CaseInsensitiveOption);
            const auto condition=bracketCondition.match(linaccess::decodeVbaSource(code)); assert(condition.hasMatch());
            const auto fieldName=condition.captured(2);
            const auto probeCode=("Sub BracketProbe()\nIf "+condition.captured(1)+" Then\nMe.Tag=\"selected\"\nElse\nMe.Tag=\"clear\"\nEnd If\nEnd Sub\n").toUtf8();
            AccessFormPreview bracketProbe("original bracket probe",definition,probeCode,db);
            auto probe=makePreviewFormObject(bracketProbe); std::vector<linaccess::VbaValue> probeArguments;
            const auto actualBracketValue=probe->item(fieldName.toStdString());
            assert(actualBracketValue.kind!=linaccess::VbaValue::Kind::Error);
            // This original control is calculated from a combo column. Preserve
            // its read-only semantics rather than manufacturing an assignment.
            try { probe->assignItem(fieldName.toStdString(),linaccess::VbaValue::logical(true)); assert(false); }
            catch(const linaccess::VbaRuntimeError& error) { assert(error.number==2448); }
            const bool selected=actualBracketValue.kind!=linaccess::VbaValue::Kind::Null && actualBracketValue.kind!=linaccess::VbaValue::Kind::Text && actualBracketValue.asNumber()==-1;
            assert(bracketProbe.executeProcedure("BracketProbe",probeArguments).status==linaccess::VbaExecutionStatus::Completed);
            assert(probe->member("Tag",{}).asText()==(selected ? "selected" : "clear"));
            assert(bracketProbe.close());
            std::cout<<"Original bracket condition: native calculated value executed; full handler not executed by this probe\n";
            // Original With/GoToControl/SendKeys paths: real native focus and
            // queued events, with actual calculated source values kept private.
            const auto originalLines=linaccess::decodeVbaSource(code).split('\n');
            const auto originalAnalysis=linaccess::analyzeVba(linaccess::decodeVbaSource(code).toStdString());
            {
                const auto statement=QRegularExpression(R"sql(\.\s*Execute\s+"((?:[^"]|"")*)"\s*$)sql",QRegularExpression::CaseInsensitiveOption).match(originalLines[2142]);
                assert(statement.hasMatch()); auto sql=statement.captured(1); sql.replace("\"\"","\"");
                const auto plan=linaccess::compileAccessQuery(sql); assert(!plan.makeTable.isEmpty());
                auto projected=sql;
                const QRegularExpression into(R"(\bINTO\s+\[[^\]]+\]\s+(?=FROM\b))",QRegularExpression::CaseInsensitiveOption);
                assert(into.match(projected).hasMatch()); projected.remove(into);
                QStringList projectedColumns;
                const auto rows=[&](const QString& source) {
                    QSqlQuery actual(db); assert(actual.exec(source));
                    assert(actual.record().count()==14);
                    QStringList names; for (int i=0;i<14;++i) names.append(actual.record().fieldName(i));
                    if (projectedColumns.isEmpty()) projectedColumns=names; else assert(names==projectedColumns);
                    std::multiset<QByteArray> result;
                    while (actual.next()) {
                        QJsonArray row;
                        for (int column=0;column<actual.record().count();++column) row.append(QJsonValue::fromVariant(actual.value(column)));
                        result.insert(QJsonDocument(row).toJson(QJsonDocument::Compact));
                    }
                    assert(!actual.lastError().isValid()); return result;
                };
                const auto expected=rows(projected);
                assert(expected.empty()); // No active sale in this unchanged fixture workflow.
                QSqlQuery guard(db); assert(guard.exec("SAVEPOINT original_make_table_probe"));
                const auto target=db.driver()->escapeIdentifier(plan.makeTable,QSqlDriver::TableName);
                assert(guard.exec("DROP TABLE "+target));
                auto database=linaccess::makeDaoDatabase(db);
                database->member("Execute",{linaccess::VbaValue::string(sql.toStdString())});
                assert(database->member("RecordsAffected",{}).asNumber()==expected.size());
                assert(rows("SELECT * FROM "+target)==expected);
                assert(guard.exec("ROLLBACK TO original_make_table_probe") && guard.exec("RELEASE original_make_table_probe"));
                std::cout<<"Original make-table statement: 14 columns, empty-workflow snapshot; "<<expected.size()<<" rows independently compared\n";
            }
            int keyboardHandlers=0,plainTextChecks=0,emptyWorkflowFailures=0;
            for (const int line:{229,244,1943}) {
                QString procedure;
                for (const auto& candidate:originalAnalysis.procedures) if (candidate.line<=line) procedure=QString::fromStdString(candidate.name);
                assert(!procedure.isEmpty());
                const auto input=QRegularExpression(R"(^\s*SendKeys\s+\.([\p{L}\p{N}_]+),\s*False\s*$)",QRegularExpression::CaseInsensitiveOption).match(originalLines[line-1].trimmed());
                const auto target=QRegularExpression(R"rx(DoCmd\.GoToControl\s+"([^"]+)")rx",QRegularExpression::CaseInsensitiveOption).match(originalLines[line-2]);
                assert(input.hasMatch() && target.hasMatch());
                auto keyValue=makePreviewFormObject(initialized)->item(input.captured(1).toStdString());
                if (keyValue.kind==linaccess::VbaValue::Kind::Null) {
                    QString expression;
                    for (const auto& control:layout.controls) if (control.name.compare(input.captured(1),Qt::CaseInsensitive)==0) expression=control.controlSource;
                    const QRegularExpression reference(R"(=\s*\[([^\]]+)\]\s*\.\s*\[(?:Form|Formular)\]\s*!\s*\[([^\]]+)\])",QRegularExpression::CaseInsensitiveOption);
                    const auto source=reference.match(expression);
                    if (source.hasMatch()) {
                        const auto canonical=makePreviewFormObject(initialized)->member("Controls",{linaccess::VbaValue::string(source.captured(1).toStdString())}).object->member("Name",{}).asText();
                        auto* child=dynamic_cast<AccessFormPreview*>(controlInDocument(initialized,QString::fromStdString(canonical)));
                        assert(child && child->hasDataSource());
                        auto* records=child->findChild<QComboBox*>("access-record-selector",Qt::FindDirectChildrenOnly); assert(records);
                        if (records->count()==1) {
                            const auto sourceName=QString::fromStdString(makePreviewFormObject(*child)->member("RecordSource",{}).asText());
                            const auto sql=sourceName.trimmed().startsWith("SELECT",Qt::CaseInsensitive) || sourceName.trimmed().startsWith("PARAMETERS",Qt::CaseInsensitive) ? sourceName : linaccess::accessQuerySource(db,sourceName);
                            // Normal search input over unchanged query definitions;
                            // inspect transitive saved queries for LIKE inputs.
                            const auto prepared=primeOriginalSearchInputs(sql,db,initialized,layout);
                            if (prepared) (void)makePreviewFormObject(*child)->member("Requery",{});
                            std::cerr<<"Original keyboard search preconditions: inputs="<<prepared<<", records="<<records->count()-1<<'\n';
                        }
                        for (int index=1;index<records->count() && keyValue.kind==linaccess::VbaValue::Kind::Null;++index) {
                            records->setCurrentIndex(index); assert(records->currentIndex()==index);
                            keyValue=makePreviewFormObject(initialized)->item(input.captured(1).toStdString());
                        }

                    }
                }
                assert(keyValue.kind!=linaccess::VbaValue::Kind::Null && keyValue.kind!=linaccess::VbaValue::Kind::Error);
                const auto keyText=QString::fromStdString(keyValue.asText()); assert(!keyText.isEmpty());
                auto requested=target.captured(1);
                if (requested.startsWith('[') && requested.endsWith(']')) { requested=requested.mid(1,requested.size()-2); requested.replace("]]","]"); }
                const auto canonical=makePreviewFormObject(initialized)->member("Controls",{linaccess::VbaValue::string(requested.toStdString())}).object->member("Name",{}).asText();
                auto* targetWidget=controlInDocument(initialized,QString::fromStdString(canonical));
                auto* targetLine=qobject_cast<QLineEdit*>(targetWidget);
                if (auto* combo=qobject_cast<QComboBox*>(targetWidget);combo && combo->isEditable()) targetLine=combo->lineEdit();
                const bool plain=targetLine && !targetLine->isReadOnly() && !keyText.contains(QRegularExpression(R"([+^%~{}()\x00-\x1f])"));
                QApplication::setActiveWindow(&initialized); targetWidget->setFocus(); QApplication::processEvents();
                if (plain) targetLine->selectAll();
                QVector<QPair<QPointer<QLineEdit>,QString>> before;
                for (auto* editor:initialized.findChildren<QLineEdit*>()) if (editor->objectName().startsWith("access-control-")) before.append({editor,editor->text()});
                const QPointer<QWidget> previousFocus=QApplication::focusWidget();
                OriginalKeyDelivery delivery(initialized); std::vector<linaccess::VbaValue> args;
                const auto result=initialized.executeProcedure(procedure,args);
                if (result.status!=linaccess::VbaExecutionStatus::Completed)
                    std::cerr<<"Original keyboard handler: status="<<static_cast<int>(result.status)<<", error="<<result.errorNumber<<", line="<<result.line<<'\n';
                assert(result.status==linaccess::VbaExecutionStatus::Completed);
                assert(delivery.presses==0); // Wait=False has not delivered yet.
                const QPointer<QWidget> queuedFocus=QApplication::focusWidget();
                auto* queuedEditor=qobject_cast<QLineEdit*>(queuedFocus.data());
                if (auto* combo=qobject_cast<QComboBox*>(queuedFocus.data());combo && combo->isEditable()) queuedEditor=combo->lineEdit();
                QString expectedEntry;
                if (plain) {
                    assert(queuedEditor);
                    expectedEntry=queuedEditor->text();
                    const auto start=queuedEditor->hasSelectedText() ? queuedEditor->selectionStart() : queuedEditor->cursorPosition();
                    expectedEntry.replace(start,queuedEditor->selectedText().size(),keyText);
                }
                QApplication::processEvents();
                int changed=0; for (const auto& entry:before) if (entry.first && entry.first->text()!=entry.second) ++changed;
                assert(delivery.presses>0 && delivery.dialogs==0);
                assert(changed>0 || QApplication::focusWidget()!=previousFocus || QApplication::activePopupWidget());
                if (plain) {
                    const auto entered=delivery.sawReturn ? delivery.beforeReturn : queuedEditor->text();
                    if (entered!=expectedEntry) std::cerr<<"Original keyboard entry boundary: receiver matches="<<(delivery.firstReceiver==targetLine)<<", whole selection="<<delivery.initialWholeSelection<<", Return observed="<<delivery.sawReturn<<'\n';
                    assert(delivery.firstReceiver==queuedEditor && entered==expectedEntry); ++plainTextChecks;
                    for (const auto& control:layout.controls) if (control.name.compare(QString::fromStdString(canonical),Qt::CaseInsensitive)==0 && !control.afterUpdate.isEmpty()) {
                        const auto updated=initialized.eventResult("control:"+control.name+":AfterUpdate");
                        if (!updated || updated->status!=linaccess::VbaExecutionStatus::Completed)
                            std::cerr<<"Original keyboard AfterUpdate: status="<<(updated ? static_cast<int>(updated->status) : -1)<<", error="<<(updated ? updated->errorNumber : 0)<<", line="<<(updated ? updated->line : 0)<<'\n';
                        assert(updated);
                        if (updated->status!=linaccess::VbaExecutionStatus::Completed) {
                            // These keyboard probes deliberately have no active
                            // sale. The unchanged subsequent event builds an
                            // empty-ID filter; record its real error explicitly.
                            assert(updated->status==linaccess::VbaExecutionStatus::Error && updated->errorNumber==3075 && updated->line==1023);
                            ++emptyWorkflowFailures;
                        }
                    }
                }
                if (auto* popup=QApplication::activePopupWidget()) popup->close();
                ++keyboardHandlers;
            }
            assert(keyboardHandlers==3 && plainTextChecks==3 && emptyWorkflowFailures==2);
            std::cout<<"Original With/GoToControl/SendKeys handlers: "<<keyboardHandlers<<", independent plain-text checks: "<<plainTextChecks<<", subsequent empty-workflow errors: "<<emptyWorkflowFailures<<'\n';
            // Execute the unchanged four-position wizard handler, rather than
            // deriving execution success from syntax recognition.
            QString originalOpenProcedure,targetName; bool targetModal=false;
            const auto analysis=linaccess::analyzeVba(linaccess::decodeVbaSource(code).toStdString());
            for (const auto& procedure:analysis.procedures) if (procedure.line<=2409) originalOpenProcedure=QString::fromStdString(procedure.name);
            for (const auto& candidate:objects) {
                const auto artifact=candidate.toObject();
                if (artifact["sha256"]=="4480943bd1372986a103cc2400deca5b3a0305df6dd7fd8b77f701199363b167") {
                    targetName=artifact["name"].toString(); QFile file(directory.filePath(artifact["file"].toString())); assert(file.open(QIODevice::ReadOnly));
                    targetModal=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(file.readAll())).modal;
                }
            }
            assert(!originalOpenProcedure.isEmpty() && !targetName.isEmpty());
            std::vector<linaccess::VbaValue> noArguments;
            const auto originalOpenResult=initialized.executeProcedure(originalOpenProcedure,noArguments);
            assert(originalOpenResult.status==linaccess::VbaExecutionStatus::Completed);
            AccessFormPreview* originalTarget=nullptr;
            for (auto* widget:QApplication::topLevelWidgets()) if (auto* target=dynamic_cast<AccessFormPreview*>(widget);target && target!=&initialized && target->isVisible() && makePreviewFormObject(*target)->member("Name",{}).asText()==targetName.toStdString()) { assert(!originalTarget); originalTarget=target; }
            assert(originalTarget);
            assert(originalTarget->isModal()==targetModal);
            assert(originalTarget->close());
            std::cout<<"Original four-position OpenForm handler: native target opened and execution completed\n";
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
    acknowledge.stop();
    {
        assert(query.exec("SAVEPOINT original_document_opener"));
        QByteArray openerCode,openerLayout; QString openerName;
        for (const auto& entry:objects) {
            const auto artifact=entry.toObject();
            if (artifact["sha256"]!="8cde46e4573af8191abdbf9eb66d3993b640bbb249fee202872492d3457ad77b") continue;
            QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly)); openerCode=input.readAll();
            assert(QCryptographicHash::hash(openerCode,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
            openerName=artifact["name"].toString().mid(5);
        }
        for (const auto& entry:objects) {
            const auto artifact=entry.toObject();
            if (artifact["kind"]!="form" || artifact["name"].toString()!=openerName) continue;
            QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly)); openerLayout=input.readAll();
            assert(QCryptographicHash::hash(openerLayout,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
        }
        assert(!openerCode.isEmpty() && !openerLayout.isEmpty());
        const auto source=linaccess::decodeVbaSource(openerCode);
        const auto procedures=linaccess::analyzeVba(source.toStdString()).procedures;
        QString procedure; for (const auto& candidate:procedures) if (candidate.line<=16) procedure=QString::fromStdString(candidate.name);
        const QRegularExpression openCall(R"re(DoCmd\.OpenForm\s+"([^"]+)")re",QRegularExpression::CaseInsensitiveOption);
        const auto call=openCall.match(source); assert(call.hasMatch() && !procedure.isEmpty());
        int unexpectedDialogs=0;
        QTimer dismissUnexpected; dismissUnexpected.setInterval(1);
        QObject::connect(&dismissUnexpected,&QTimer::timeout,[&] {
            if (auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) { ++unexpectedDialogs; box->accept(); }
        }); dismissUnexpected.start();
        QTimer openerDeadline; openerDeadline.setSingleShot(true);
        QObject::connect(&openerDeadline,&QTimer::timeout,[] { assert(false && "Original document opener timed out"); }); openerDeadline.start(5000);
        QMainWindow projectWindow;
        auto* projectToolbar=projectWindow.addToolBar("Project"); projectToolbar->addAction("Test command");
        projectWindow.menuBar()->addMenu("Project")->addAction(projectToolbar->actions().first());
        QPointer<AccessFormPreview> target;
        {
            AccessFormPreview opener(openerName,openerLayout,openerCode,db,&projectWindow); opener.show();
            QApplication::setActiveWindow(&opener); QApplication::processEvents();
            std::vector<linaccess::VbaValue> none;
            const auto outcome=opener.executeProcedure(procedure,none);
            if (outcome.status!=linaccess::VbaExecutionStatus::Completed)
                std::cerr<<"Original document opener: status="<<static_cast<int>(outcome.status)<<", error="<<outcome.errorNumber<<", line="<<outcome.line<<'\n';
            assert(outcome.status==linaccess::VbaExecutionStatus::Completed && !opener.isVisible());
            for (auto* widget:QApplication::topLevelWidgets())
                if (auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&opener && candidate->isVisible()
                    && makePreviewFormObject(*candidate)->member("Name",{}).asText()==call.captured(1).toStdString()) { assert(!target); target=candidate; }
            assert(target && target->parentWidget()==&projectWindow);
            for (const auto* event:{"OnOpen","OnLoad","OnActivate","OnCurrent"})
                if (const auto result=target->eventResult(event)) assert(result->status==linaccess::VbaExecutionStatus::Completed);
        }
        QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
        assert(target && target->isVisible()); assert(target->close());
        QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!target);
        openerDeadline.stop(); dismissUnexpected.stop(); assert(unexpectedDialogs==0);
        assert(query.exec("ROLLBACK TO original_document_opener")); assert(query.exec("RELEASE original_document_opener"));
        std::cout<<"Original document opener: unchanged handler closes caller, target survives caller destruction\n";
    }
    {
        assert(query.exec("SAVEPOINT original_login_boundary"));
        QByteArray loginCode,loginLayout; QString loginName;
        for (const auto& entry:objects) {
            const auto artifact=entry.toObject();
            if (artifact["sha256"]!="fa72dad21f4875398da7de601750d963c1d5d1292a36f3db6fbc240ed66959ba") continue;
            QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly)); loginCode=input.readAll();
            assert(QCryptographicHash::hash(loginCode,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
            loginName=artifact["name"].toString().mid(5);
        }
        for (const auto& entry:objects) {
            const auto artifact=entry.toObject();
            if (artifact["kind"]!="form" || artifact["name"].toString()!=loginName) continue;
            QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly)); loginLayout=input.readAll();
            assert(QCryptographicHash::hash(loginLayout,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1());
        }
        assert(!loginCode.isEmpty() && !loginLayout.isEmpty());
        QMainWindow projectWindow;
        auto* toolbar=projectWindow.addToolBar("Project"); toolbar->addAction("Test command");
        projectWindow.menuBar()->addMenu("Project")->addAction(toolbar->actions().first());
        AccessFormPreview login(loginName,loginLayout,loginCode,db,&projectWindow); login.show();
        QApplication::setActiveWindow(&login); QApplication::processEvents();
        auto loginObject=makePreviewFormObject(login);
        auto user=loginObject->member("Controls",{linaccess::VbaValue::string("user")}).object;
        user->member("Requery",{});
        auto* choices=login.findChild<QComboBox*>("access-control-user"); assert(choices && choices->count()>0);
        // Exercise the unchanged original handler with a deliberately synthetic
        // password. It fails on its invalid sixth-column read before comparison;
        // this does not authenticate against or expose original credentials.
        user->assign("Value",linaccess::qtFieldValue(choices->itemData(0)));
        assert(user->member("ColumnCount",{}).asNumber()==5);
        loginObject->assignItem("Psw",linaccess::VbaValue::string("synthetic non-credential input"));
        QString handler;
        for (const auto& procedure:linaccess::analyzeVba(linaccess::decodeVbaSource(loginCode).toStdString()).procedures)
            if (procedure.line<=25) handler=QString::fromStdString(procedure.name);
        assert(!handler.isEmpty());
        int unexpected=0; QTimer dismiss; dismiss.setInterval(1);
        QObject::connect(&dismiss,&QTimer::timeout,[&] { if (auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) { ++unexpected; box->accept(); } }); dismiss.start();
        std::vector<linaccess::VbaValue> none;
        const auto outcome=login.executeProcedure(handler,none);
        std::cerr<<"Original login source-column boundary: status="<<static_cast<int>(outcome.status)<<", error="<<outcome.errorNumber<<", line="<<outcome.line<<", dialogs="<<unexpected<<'\n';
        assert(outcome.status==linaccess::VbaExecutionStatus::Error && outcome.errorNumber==5 && outcome.line==34);
        assert(login.isVisible() && unexpected==0);
        dismiss.stop();
        assert(AccessFormPreview::closeProjectDocuments(db));
        assert(query.exec("ROLLBACK TO original_login_boundary")); assert(query.exec("RELEASE original_login_boundary"));
    }
    testOriginalKeyboardEvents(directory,objects,db);
    testOriginalTabOrder(directory,objects,db);
    testOriginalPasswordMask(directory,objects);
    testOriginalControlFonts(directory,objects);
    testOriginalAllowEdits(directory,objects,db);
    testOriginalControlDefaults(directory,objects);
    testOriginalTextAlignment(directory,objects);
    testOriginalForeColor(directory,objects);
    testOriginalBackgrounds(directory,objects);
    testOriginalTimer(directory,objects,db);
    testOriginalOpenReport(directory,objects,db);
    testOriginalCalendar(directory,objects,db);
}
