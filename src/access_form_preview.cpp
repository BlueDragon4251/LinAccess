#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "native_project_close_guard.hpp"
#include "access_list_box.hpp"
#include "native_send_keys.hpp"
#include "native_input_box.hpp"
#include "native_clock_keys.hpp"
#include "native_cell_zoom.hpp"
#include "native_find_value.hpp"
#include <QVariantMap>

#include "vba_runtime.hpp"
#include "vba_module_context.hpp"
#include "vba_currency.hpp"
#include "vba_calendar.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "access_value_list.hpp"
#include "access_autonumber.hpp"
#include "access_format.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"

#include <QApplication>
#include <QFont>
#include <QCalendarWidget>
#include <QDate>
#include <QDateTime>
#include <QCloseEvent>
#include <QKeyEvent>
#include <QTimer>
#include <QEventLoop>
#include <QScopedValueRollback>
#include <QScopeGuard>
#include <QSignalBlocker>
#include <QCheckBox>
#include <QGroupBox>
#include <QButtonGroup>
#include <QRadioButton>
#include <QWheelEvent>
#include <QComboBox>
#include <QStandardItemModel>
#include <QTreeView>
#include <QMainWindow>
#include <QFormLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QColor>
#include <QPalette>
#include <QLineEdit>
#include <QMessageBox>
#include <QPushButton>
#include <QRegularExpression>
#include <QScrollArea>
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlRecord>
#include <QSqlField>
#include <QVBoxLayout>

#include <algorithm>
#include <cmath>
#include <functional>
#include <limits>
#include <stdexcept>

namespace {
void waitForNativeDialog(QDialog* document) {
    if(!document || !document->isVisible()) return;
    QEventLoop loop;
    class VisibilityWatcher final : public QObject {
    public:
        explicit VisibilityWatcher(QEventLoop& loop):loop_(loop) {}
        bool eventFilter(QObject* watched,QEvent* event) override {
            if(event->type()==QEvent::Hide) loop_.quit();
            return QObject::eventFilter(watched,event);
        }
    private:
        QEventLoop& loop_;
    } watcher(loop);
    document->installEventFilter(&watcher);
    QObject::connect(document,&QDialog::finished,&loop,&QEventLoop::quit);
    QObject::connect(document,&QObject::destroyed,&loop,&QEventLoop::quit);
    loop.exec();
}
QPointer<AccessFormPreview>& activeDocument() {
    static QPointer<AccessFormPreview> document;
    return document;
}
struct DatabaseWindowContext { QPointer<QWidget> window; QString project; };
DatabaseWindowContext& databaseWindowContext() { static DatabaseWindowContext context; return context; }
QSet<QString>& closingProjects() { static QSet<QString> projects; return projects; }
QString projectKey(const QSqlDatabase& db) { return db.connectionName()+QChar(0)+db.databaseName(); }
int& nativeExecutionDepth() { static thread_local int depth=0; return depth; }
QVector<QPointer<AccessFormPreview>>& pendingFormDisposals() {
    static thread_local QVector<QPointer<AccessFormPreview>> pending; return pending;
}
void disposeClosedForms() {
    if (nativeExecutionDepth()!=0) return;
    auto pending=std::move(pendingFormDisposals()); pendingFormDisposals().clear();
    for (const auto& form:pending) if (form) form->deleteLater();
}
void deferFormDisposal(AccessFormPreview* form) {
    auto& pending=pendingFormDisposals();
    if (!pending.contains(form)) pending.append(form);
    disposeClosedForms();
}
linaccess::VbaValue defaultLiteral(QString expression) {
    expression=expression.trimmed(); if(expression.startsWith('=')) expression=expression.mid(1).trimmed();
    if(expression.compare("Null",Qt::CaseInsensitive)==0) return {};
    if(expression.compare("True",Qt::CaseInsensitive)==0 || expression.compare("Yes",Qt::CaseInsensitive)==0) return linaccess::VbaValue::logical(true);
    if(expression.compare("False",Qt::CaseInsensitive)==0 || expression.compare("No",Qt::CaseInsensitive)==0) return linaccess::VbaValue::logical(false);
    static const QRegularExpression quoted(R"literal(^"(?:[^"]|"")*"$)literal");
    static const QRegularExpression number(R"(^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$)");
    if(quoted.match(expression).hasMatch()) return linaccess::VbaValue::string(expression.mid(1,expression.size()-2).replace("\"\"","\"").toUtf8().toStdString());
    if(number.match(expression).hasMatch()) {
        bool valid=false; const auto value=expression.toDouble(&valid);
        if(!valid || !std::isfinite(value)) throw linaccess::VbaRuntimeError(6,"Default numeric literal overflow");
        return linaccess::VbaValue::numeric(value);
    }
    throw linaccess::VbaRuntimeError(3251,"Computed/default expression unsupported");
}
void clearDefaultError(QWidget* widget) {
    widget->setProperty("la_default_error",0);
    auto* editor=qobject_cast<QLineEdit*>(widget);
    if(auto* combo=qobject_cast<QComboBox*>(widget)) editor=combo->lineEdit();
    if(editor && editor->property("la_default_placeholder").isValid()) {
        editor->setPlaceholderText(editor->property("la_default_placeholder").toString()); editor->setProperty("la_default_placeholder",QVariant{});
    }
    if(widget->property("la_default_tooltip").isValid()) {
        widget->setToolTip(widget->property("la_default_tooltip").toString()); widget->setProperty("la_default_tooltip",QVariant{});
    }
}
class AccessComboBox final : public QComboBox {
public:
    using QComboBox::QComboBox;
    std::function<bool()> preparePopup;
    void showPopup() override {
        if (!isVisible() || !isEnabled() || property("la_access_input_readonly").toBool()) return;
        if (preparePopup && !preparePopup()) return;
        QComboBox::showPopup();
    }
protected:
    void wheelEvent(QWheelEvent* event) override {
        if (property("la_access_input_readonly").toBool()) { event->accept(); return; }
        QComboBox::wheelEvent(event);
    }
    void keyPressEvent(QKeyEvent* event) override {
        if (property("la_access_input_readonly").toBool() && (event->key()==Qt::Key_Up || event->key()==Qt::Key_Down || event->key()==Qt::Key_PageUp || event->key()==Qt::Key_PageDown)) { event->accept(); return; }
        QComboBox::keyPressEvent(event);
    }
};
class AccessPushButton final:public QPushButton {
public:
    using QPushButton::QPushButton;
protected:
    void nextCheckState() override { if (!property("la_access_input_readonly").toBool()) QPushButton::nextCheckState(); }
};
class AccessCheckBox final:public QCheckBox {
public:
    using QCheckBox::QCheckBox;
protected:
    void nextCheckState() override { if (!property("la_access_input_readonly").toBool()) QCheckBox::nextCheckState(); }
};
class AccessRadioButton final:public QRadioButton {
public:
    using QRadioButton::QRadioButton;
protected:
    void nextCheckState() override { if (!property("la_access_input_readonly").toBool()) QRadioButton::nextCheckState(); }
};
class AccessMessageBox final : public QMessageBox {
public:
    AccessMessageBox(QWidget* owner,bool dismissible) : QMessageBox(owner),dismissible_(dismissible) {}
protected:
    void keyPressEvent(QKeyEvent* event) override {
        if (event->key()==Qt::Key_Escape && !dismissible_) { event->accept(); return; }
        QMessageBox::keyPressEvent(event);
    }
    void closeEvent(QCloseEvent* event) override {
        if (!dismissible_) { event->ignore(); return; }
        QMessageBox::closeEvent(event);
    }
private:
    bool dismissible_{};
};
int nativeMessageBox(QWidget* owner,std::string_view prompt,int style,const std::optional<std::string>& title) {
    constexpr int allowed=0x0f|0xf0|0x300|0x1000|0x4000|0x10000|0x80000|0x100000;
    const int kind=style&15,icon=style&0xf0,defaultIndex=(style>>8)&3;
    if (style<0 || (style&~allowed)!=0 || kind>5 || (icon!=0 && icon!=16 && icon!=32 && icon!=48 && icon!=64))
        throw linaccess::VbaRuntimeError(5,"Invalid MsgBox style");
    if ((style&0x1000)!=0 || (style&0x4000)!=0) throw linaccess::VbaRuntimeError(3251,"Windows system-modal/Help MsgBox is unsupported");
    const QVector<QVector<QMessageBox::StandardButton>> groups{{QMessageBox::Ok},{QMessageBox::Ok,QMessageBox::Cancel},
        {QMessageBox::Abort,QMessageBox::Retry,QMessageBox::Ignore},{QMessageBox::Yes,QMessageBox::No,QMessageBox::Cancel},
        {QMessageBox::Yes,QMessageBox::No},{QMessageBox::Retry,QMessageBox::Cancel}};
    const auto choices=groups[kind];
    if (defaultIndex>=choices.size()) throw linaccess::VbaRuntimeError(5,"MsgBox default button unavailable");
    const bool dismissible=kind==0 || choices.contains(QMessageBox::Cancel);
    AccessMessageBox box(owner,dismissible);
    box.setWindowModality(Qt::ApplicationModal);
    box.setTextFormat(Qt::PlainText);
    box.setText(QString::fromUtf8(prompt.data(),prompt.size()));
    box.setWindowTitle(title ? QString::fromStdString(*title) : QStringLiteral("LinAccess"));
    box.setIcon(icon==16 ? QMessageBox::Critical : icon==32 ? QMessageBox::Question : icon==48 ? QMessageBox::Warning : icon==64 ? QMessageBox::Information : QMessageBox::NoIcon);
    for (const auto choice : choices) box.addButton(choice);
    box.setDefaultButton(choices[defaultIndex]);
    if (!dismissible) box.setWindowFlag(Qt::WindowCloseButtonHint,false);
    else box.setEscapeButton(kind==0 ? QMessageBox::Ok : QMessageBox::Cancel);
    if ((style&0x100000)!=0) box.setLayoutDirection(Qt::RightToLeft);
    if ((style&0x80000)!=0)
        if (auto* label=box.findChild<QLabel*>("qt_msgbox_label")) label->setAlignment(Qt::AlignRight|Qt::AlignVCenter);
    if ((style&0x10000)!=0) QTimer::singleShot(0,&box,[&box] { box.raise(); box.activateWindow(); });
    const auto result=box.exec();
    auto selected=box.standardButton(box.clickedButton());
    if (selected==QMessageBox::NoButton) {
        for (const auto choice : choices) if (result==choice) selected=choice;
        // Existing OK-only callers may close the dialog through QDialog::accept.
        if (kind==0 && result==QDialog::Accepted) selected=QMessageBox::Ok;
        else if (result==QDialog::Rejected && choices.contains(QMessageBox::Cancel)) selected=QMessageBox::Cancel;
    }
    switch (selected) {
    case QMessageBox::Ok: return 1;
    case QMessageBox::Cancel: return 2;
    case QMessageBox::Abort: return 3;
    case QMessageBox::Retry: return 4;
    case QMessageBox::Ignore: return 5;
    case QMessageBox::Yes: return 6;
    case QMessageBox::No: return 7;
    default: throw linaccess::VbaRuntimeError(5,"MsgBox closed without a supported button");
    }
}
QString sqlIdentifier(QString text) {
    text.replace('"', "\"\"");
    return '"' + text + '"';
}
int pixels(int twips) {
    // 96dpi * twips/1440; do not let a control become unclickably small.
    return qMax(1, qRound(twips / 15.0));
}
bool eventProcedure(const QString& value) {
    return value.compare("[Event Procedure]",Qt::CaseInsensitive) == 0;
}
QString eventName(QString name) {
    name.replace(QRegularExpression("[^\\p{L}\\p{N}_]"),"_"); return name;
}
bool failed(const linaccess::VbaExecutionResult& result) {
    return result.status == linaccess::VbaExecutionStatus::Error || result.status == linaccess::VbaExecutionStatus::Unsupported;
}
QString textOf(QWidget* editor) {
    if(auto* list=dynamic_cast<AccessListBox*>(editor)) return list->valueText();
    if (qobject_cast<QGroupBox*>(editor)) return editor->property("la_option_value").toString();
    if (auto* calendar=qobject_cast<QCalendarWidget*>(editor)) return calendar->selectedDate().toString(Qt::ISODate)+" 00:00:00";
    if (auto* line = qobject_cast<QLineEdit*>(editor)) return line->text();
    if (auto* checkbox = qobject_cast<QCheckBox*>(editor)) return checkbox->isChecked() ? "-1" : "0";
    if (auto* combo = qobject_cast<QComboBox*>(editor)) {
        if (combo->property("la_rows_loaded").toBool() && combo->currentIndex()>=0 && combo->currentText()==combo->itemText(combo->currentIndex()))
            return combo->currentData().toString();
        return combo->currentText();
    }
    return {};
}
void setText(QWidget* editor, const QString& text,bool isNull=false) {
    if(auto* list=dynamic_cast<AccessListBox*>(editor)) { list->setValueText(text,isNull); return; }
    if (auto* group=qobject_cast<QGroupBox*>(editor)) {
        group->setProperty("la_option_value",text);
        if (auto* buttons=group->findChild<QButtonGroup*>("access-options")) {
            buttons->setExclusive(false);
            for (auto* button:buttons->buttons()) button->setChecked(!text.isEmpty() && button->property("la_option_value").toString()==text);
            buttons->setExclusive(true);
        }
        return;
    }
    if (auto* calendar=qobject_cast<QCalendarWidget*>(editor)) {
        const auto date=QDate::fromString(text.left(10),Qt::ISODate); if (date.isValid()) calendar->setSelectedDate(date);
    }
    else if (auto* line = qobject_cast<QLineEdit*>(editor)) line->setText(text);
    else if (auto* checkbox = qobject_cast<QCheckBox*>(editor))
        checkbox->setChecked(text == "-1" || text == "1" || text.compare("true", Qt::CaseInsensitive) == 0);
    else if (auto* combo = qobject_cast<QComboBox*>(editor)) {
        if(isNull) { combo->setCurrentIndex(-1); combo->setEditText({}); return; }
        if (combo->property("la_rows_loaded").toBool()) {
            for (int i=0;i<combo->count();++i) if (!combo->itemData(i).isNull() && combo->itemData(i).toString()==text) { combo->setCurrentIndex(i); return; }
            combo->setCurrentIndex(-1); combo->setEditText(text);
        } else combo->setCurrentText(text);
    }
}
int columnIndex(QString source, const QStringList& columns) {
    source = source.trimmed();
    if (source.startsWith('[') && source.endsWith(']')) source = source.mid(1, source.size() - 2);
    for (int i = 0; i < columns.size(); ++i)
        if (columns[i].compare(source, Qt::CaseInsensitive) == 0) return i;
    return -1;
}
}

QString AccessFormPreview::controlEventPrefix(const QString& name) const {
    for (const auto& control:layout_.controls)
        if (control.name.compare(name,Qt::CaseInsensitive)==0)
            return control.eventProcPrefix.isEmpty() ? eventName(control.name) : control.eventProcPrefix;
    return eventName(name);
}
std::shared_ptr<linaccess::VbaObject> AccessFormPreview::recordsetSnapshot() {
    if (layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Report has no form recordset");
    if (!dataSourceAvailable_) throw linaccess::VbaRuntimeError(3146,"Form record source unavailable");
    if (dataTruncated_) throw linaccess::VbaRuntimeError(3035,"Form recordset is incomplete");
    if (!recordsetSnapshot_) recordsetSnapshot_=linaccess::makeDaoSnapshot(db_,dataColumns_,dataTypes_,recordRows_,rowIds_);
    return recordsetSnapshot_;
}
QString AccessFormPreview::controlMemberName(const QString& name) const {
    if (controls_.contains(name.toCaseFolded())) return name;
    QString resolved;
    for (const auto& control:layout_.controls) {
        if (!controls_.contains(control.name.toCaseFolded()) || eventName(control.name).compare(name,Qt::CaseInsensitive)!=0) continue;
        if (!resolved.isEmpty() && resolved.compare(control.name,Qt::CaseInsensitive)!=0)
            throw linaccess::VbaRuntimeError(2465,"Ambiguous generated control member");
        resolved=control.name;
    }
    return resolved.isEmpty() ? name : resolved;
}

AccessFormPreview* activePreviewDocument() { return activeDocument().data(); }

class PreviewVbaHost final : public linaccess::VbaHost {
public:
    explicit PreviewVbaHost(AccessFormPreview& preview) : preview_(preview) {}
    int compareStrings(std::string_view left,std::string_view right,linaccess::VbaCompareMode mode) override {
        if (mode==linaccess::VbaCompareMode::Binary) return linaccess::VbaHost::compareStrings(left,right,mode);
        return QString::compare(QString::fromUtf8(left.data(),left.size()),QString::fromUtf8(right.data(),right.size()),Qt::CaseInsensitive);
    }
    std::optional<std::string> field(std::string_view name) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (key == "dirty") return preview_.dirty_ ? "-1" : "0";
        if (key == "newrecord") return makePreviewFormObject(preview_)->member("NewRecord",{}).asField();
        if (key == "dataentry" || key == "modal" || key == "popup" || key == "openargs" || key == "currentrecord" || key=="orderby" || key=="orderbyon") return makePreviewFormObject(preview_)->member(name,{}).asField();
        return preview_.getField(name);
    }
    linaccess::VbaValue controlValue(std::string_view name) override { requireInstance(); return preview_.getFieldValue(name); }
    void enterModule(std::string_view source) override { moduleScopes_.push_back(source==preview_.vba_.toUtf8().toStdString()); ++nativeExecutionDepth(); }
    void leaveModule() noexcept override { moduleScopes_.pop_back(); --nativeExecutionDepth(); disposeClosedForms(); }
    bool hasInstanceIdentifier(std::string_view name) override {
        return instanceScope() && (recordProperty(name) || preview_.controls_.contains(preview_.controlMemberName(QString::fromUtf8(name.data(),name.size())).toCaseFolded()));
    }
    std::shared_ptr<linaccess::VbaModuleState> moduleState(std::string_view source) override {
        if (source!=preview_.vba_.toUtf8().toStdString() && !linaccess::isVbaClassModule(source))
            return linaccess::standardVbaModuleState(preview_.db_,source);
        const auto key=std::string(source);
        if (const auto found=preview_.moduleStates_.find(key);found!=preview_.moduleStates_.end()) return found->second;
        auto state=linaccess::createVbaModuleState(source); preview_.moduleStates_.emplace(key,state); return state;
    }
    std::vector<std::shared_ptr<linaccess::VbaModuleState>> globalModuleStates() override {
        return linaccess::publicVbaModuleStates(preview_.db_);
    }
    std::optional<linaccess::VbaValue> resolveIdentifier(std::string_view name,bool reference) override {
        if (!instanceScope()) return std::nullopt;
        if (auto value=preview_.getModuleMember(name,false)) return value;
        const auto canonical=memberControlName(name);
        if (!preview_.controls_.contains(QString::fromStdString(canonical).toCaseFolded())) {
            if(recordProperty(name)) return makePreviewFormObject(preview_)->member(name,{});
            return std::nullopt;
        }
        return reference ? makePreviewFormObject(preview_)->objectItem(canonical) : preview_.getFieldValue(canonical);
    }
    bool assignIdentifier(std::string_view name,const linaccess::VbaValue& value) override {
        if (!instanceScope()) return false;
        const auto canonical=memberControlName(name);
        if (!preview_.controls_.contains(QString::fromStdString(canonical).toCaseFolded())) {
            if(recordProperty(name)) { makePreviewFormObject(preview_)->assign(name,value); return true; }
            return false;
        }
        preview_.setFieldValue(canonical,value); return true;
    }
    linaccess::VbaValue bracketValue(std::string_view name,bool reference) override {
        if (!instanceScope()) throw linaccess::VbaRuntimeError(438,"Unqualified bracket reference requires a native instance context");
        return reference ? makePreviewFormObject(preview_)->objectItem(name) : preview_.getFieldValue(name);
    }
    linaccess::VbaValue expressionValue(std::string_view name,bool reference) override {
        try { return reference ? makePreviewFormObject(preview_)->objectItem(name) : preview_.getFieldValue(name); }
        catch(const linaccess::VbaRuntimeError& failure) {
            if(failure.number==2465 && recordProperty(name)) {
                try { return makePreviewFormObject(preview_)->member(name,{}); }
                catch(const linaccess::VbaRuntimeError& propertyFailure) { return linaccess::VbaValue::errorValue(propertyFailure.number); }
            }
            return linaccess::VbaValue::errorValue(failure.number);
        }
    }
    linaccess::VbaValue domainAggregate(std::string_view function,const std::vector<linaccess::VbaValue>& args) override {
        if (args.size()<2 || args.size()>3) throw linaccess::VbaRuntimeError(450,"Invalid domain aggregate arguments");
        if (args[0].kind==linaccess::VbaValue::Kind::Null || args[1].kind==linaccess::VbaValue::Kind::Null)
            throw linaccess::VbaRuntimeError(94,"Invalid Null domain aggregate argument");
        const auto expression=QString::fromStdString(args[0].asText()).trimmed();
        const auto domain=QString::fromStdString(args[1].asText());
        if (expression.isEmpty() || domain.isEmpty()) throw linaccess::VbaRuntimeError(5,"Empty domain aggregate argument");
        // Domain is an object name, never executable SQL. Saved non-parameter
        // queries are installed as views on the same native connection.
        linaccess::installAccessQueryViews(preview_.db_);
        if (!preview_.db_.tables(QSql::AllTables).contains(domain,Qt::CaseInsensitive))
            throw linaccess::VbaRuntimeError(3078,"Domain object unavailable");
        QString source="SELECT "+QString(function=="dsum" ? "Sum" : "Count")+"("+expression+") FROM "+sqlIdentifier(domain);
        if (args.size()==3 && args[2].kind!=linaccess::VbaValue::Kind::Null) {
            const auto criteria=QString::fromStdString(args[2].asText()).trimmed();
            if (!criteria.isEmpty()) source+=" WHERE "+criteria;
        }
        const auto plan=linaccess::compileAccessQuery(source);
        try {
            const auto result=linaccess::executeAccessQueryPlan(preview_.db_,plan,linaccess::bindAccessQueryParameters(plan,{}));
            if (result.rows.size()!=1 || result.rows.front().size()!=1)
                throw linaccess::VbaRuntimeError(3146,"Invalid domain aggregate result");
            return linaccess::qtFieldValue(result.rows.front().front(),"");
        } catch(const linaccess::VbaRuntimeError&) { throw; }
        catch(...) { throw linaccess::VbaRuntimeError(3146,"Domain aggregate execution failed"); }
    }
    linaccess::VbaValue expressionAggregate(std::string_view function,std::string_view field,bool allRecords) override {
        try {
        if (!preview_.dataSourceAvailable_) throw linaccess::VbaRuntimeError(3146,"Aggregate record source unavailable");
        if (preview_.dataTruncated_) throw linaccess::VbaRuntimeError(3035,"Aggregate would use incomplete records");
        if (allRecords) {
            if (function!="count") throw linaccess::VbaRuntimeError(5,"Aggregate field required");
            return linaccess::VbaValue::numeric(preview_.recordRows_.size());
        }
        const auto name=QString::fromUtf8(field.data(),field.size());
        const auto index=columnIndex(name,preview_.dataColumns_);
        if (index<0) throw linaccess::VbaRuntimeError(2465,"Aggregate field unavailable");
        double total=0; int count=0;
        for (const auto& row : preview_.recordRows_) {
            if (row[index].isNull()) continue;
            ++count;
            if (function=="sum") {
                total+=linaccess::qtFieldValue(row[index],preview_.dataTypes_.value(name.toCaseFolded())).asNumber();
                if (!std::isfinite(total)) throw linaccess::VbaRuntimeError(6,"Aggregate overflow");
            }
        }
        if (function=="count") return linaccess::VbaValue::numeric(count);
        return count ? linaccess::VbaValue::numeric(total) : linaccess::VbaValue{};
        } catch(const linaccess::VbaRuntimeError& failure) { return linaccess::VbaValue::errorValue(failure.number); }
    }
    linaccess::VbaValue fieldValue(std::string_view name) override {
        requireInstance();
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (preview_.layout_.kind==linaccess::AccessDocumentKind::Report && (key=="dirty" || key=="newrecord"))
            throw linaccess::VbaRuntimeError(438,"Report member unsupported");
        if (key == "name") return linaccess::VbaValue::string(preview_.formName_.toUtf8().toStdString());
        if (key == "caption") return makePreviewFormObject(preview_)->member("Caption",{});
        if (key == "tag") return makePreviewFormObject(preview_)->member("Tag",{});
        if (key=="dataentry" || key=="modal" || key=="popup" || key=="openargs" || key=="orderby" || key=="orderbyon" || key=="currentrecord" || key=="visible" || key=="allowadditions" || key=="allowdeletions" || key=="allowedits" || key=="selheight" || key=="filter" || key=="filteron" || key=="keypreview" || key=="timerinterval" || key=="cycle" || key=="bookmark" || key=="recordset") return makePreviewFormObject(preview_)->member(key.toStdString(),{});
        if (key == "recordsource") return linaccess::VbaValue::string(preview_.layout_.recordSource.toUtf8().toStdString());
        if (key == "dirty") return linaccess::VbaValue::logical(preview_.dirty_);
        if (key == "newrecord") return makePreviewFormObject(preview_)->member("NewRecord",{});
        if (auto value=preview_.getModuleMember(name,false)) return *value;
        return preview_.getFieldValue(memberControlName(name));
    }
    void setControlValue(std::string_view name,const linaccess::VbaValue& value) override { requireInstance(); preview_.setFieldValue(name,value); }
    void setFieldValue(std::string_view name,const linaccess::VbaValue& value) override {
        requireInstance();
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (preview_.layout_.kind==linaccess::AccessDocumentKind::Report && key=="dirty")
            throw linaccess::VbaRuntimeError(438,"Report member unsupported");
        if (key == "dataentry" || key == "modal" || key == "popup" || key == "orderby" || key == "orderbyon" || key == "visible" || key == "caption" || key=="tag" || key=="allowadditions" || key=="allowdeletions" || key=="allowedits" || key=="filter" || key=="filteron" || key=="keypreview" || key=="timerinterval" || key=="cycle" || key=="bookmark" || key=="recordset") { makePreviewFormObject(preview_)->assign(key.toUtf8().toStdString(),value); return; }
        if (key == "selheight" || key == "name" || key == "recordsource" || key == "newrecord" || key == "currentrecord" || key=="openargs")
            throw linaccess::VbaRuntimeError(438,"Form property is read-only");
        if (key == "dirty") { setField(name,value.asField()); return; }
        if (preview_.setModuleMember(name,value,false)) return;
        preview_.setFieldValue(memberControlName(name),value);
    }
    void setField(std::string_view name, std::optional<std::string> value) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="dataentry" || key=="modal" || key=="popup") throw linaccess::VbaRuntimeError(3251,"Runtime Modal/PopUp mutation is not implemented");
        if(key=="currentrecord" || key=="openargs") throw linaccess::VbaRuntimeError(438,"Form property is read-only");
        if (key == "dirty") {
            if (value && (*value == "0" || *value == "False")) {
                if (!preview_.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Record save cancelled");
            } else {
                if (!preview_.directTable_) throw linaccess::VbaRuntimeError(2448,"Record source is read-only");
                preview_.markDirty({});
            }
            return;
        }
        preview_.setField(name, std::move(value));
    }
    void message(std::string_view value) override {
        (void)nativeMessageBox(&preview_,value,0,std::nullopt);
    }
    int messageBox(std::string_view prompt,int buttons,const std::optional<std::string>& title) override {
        return nativeMessageBox(&preview_,prompt,buttons,title);
    }
    std::string inputBox(const linaccess::VbaInputBoxArguments& arguments) override {
        return linaccess::nativeInputBox(&preview_,arguments);
    }
    linaccess::VbaValue eval(std::string_view expression) override {
        ++nativeExecutionDepth();
        const auto leave=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
        return linaccess::VbaHost::eval(expression);
    }
    bool usesActiveTableContext() const { return preview_.hasActivated_ || preview_.property("la_explicit_table_context").toBool(); }
    bool tableCommand(int code) {
        if(!usesActiveTableContext()) return false;
        QWidget* owner=nullptr;
        for(auto* ancestor=preview_.parentWidget();ancestor;ancestor=ancestor->parentWidget())
            if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
        if(!owner || owner->metaObject()->indexOfMethod("dispatchTableCommand(int)")<0) return false;
        QVariantMap result;
        if(!QMetaObject::invokeMethod(owner,"dispatchTableCommand",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(int,code)))
            throw linaccess::VbaRuntimeError(3251,"Native table command dispatch is unavailable");
        if(!result.value("handled").toBool()) return false;
        const int error=result.value("error").toInt();
        if(error) throw linaccess::VbaRuntimeError(error,result.value("description").toString().toStdString());
        return true;
    }
    bool findTableRecord(const linaccess::VbaValue& sought,int mode,bool sensitive,bool first,int direction) {
        if(!usesActiveTableContext()) return false;
        QWidget* owner=nullptr;
        for(auto* ancestor=preview_.parentWidget();ancestor;ancestor=ancestor->parentWidget())
            if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
        if(!owner || owner->metaObject()->indexOfMethod("findTableRecord(QVariantMap)")<0) return false;
        auto request=linaccess::nativeFindValuePayload(sought);
        request.insert("match",mode); request.insert("sensitive",sensitive); request.insert("first",first); request.insert("direction",direction);
        QVariantMap result;
        if(!QMetaObject::invokeMethod(owner,"findTableRecord",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(QVariantMap,request)))
            throw linaccess::VbaRuntimeError(3251,"Native FindRecord dispatch is unavailable");
        if(!result.value("handled").toBool()) return false;
        const int error=result.value("error").toInt();
        if(error) throw linaccess::VbaRuntimeError(error,result.value("description").toString().toStdString());
        return true;
    }
    void findFormRecord(const linaccess::VbaValue& sought,int mode,bool sensitive,bool first,int direction) {
        auto& document=commandDocument();
        if(auto* modal=QApplication::activeModalWidget(); modal && modal!=document.window())
            throw linaccess::VbaRuntimeError(2046,"FindRecord is blocked by a modal window");
        QString field;
        const auto* focus=QApplication::focusWidget();
        for(const auto& control:document.layout_.controls) {
            auto* widget=document.controls_.value(control.name.toCaseFolded()).data();
            const auto* combo=qobject_cast<QComboBox*>(widget);
            if(!widget || (widget!=focus && (!combo || combo->lineEdit()!=focus))) continue;
            const auto column=columnIndex(control.controlSource,document.dataColumns_);
            if(column<0) continue;
            const auto type=document.dataTypes_.value(document.dataColumns_[column].toCaseFolded());
            if(!linaccess::nativeFindSupports(sought,type))
                throw linaccess::VbaRuntimeError(3251,"FindRecord value and bound field subtype are not supported together");
            field=document.dataColumns_[column]; break;
        }
        if(field.isEmpty()) throw linaccess::VbaRuntimeError(2046,"No active bound field for FindRecord");
        const bool textValue=sought.kind==linaccess::VbaValue::Kind::Text;
        (void)document.findNextRecord(linaccess::nativeFindValueText(sought),field,mode,sensitive,first ? 1 : 2,true,direction,
            textValue ? std::optional<linaccess::VbaValue>{} : sought);
    }
    void saveRecord() override {
        if(tableCommand(97)) return;
        if (!commandDocument().saveRecord()) throw linaccess::VbaRuntimeError(2501,"Record save cancelled");
    }
    void selectCurrentRecord() { commandDocument().selectCurrentRecordForCommand(); }
    void deleteCurrentRecord() { commandDocument().deleteCurrentRecord(); }
    void copyCurrentRecord() { commandDocument().copyCurrentRecord(); }
    void pasteAppendRecord() { commandDocument().pasteAppendRecord(); }
    void zoomField() { if(!commandDocument().zoomFocusedText()) throw linaccess::VbaRuntimeError(2046,"Zoom is unavailable in the active control"); }
    void findRecords() { commandDocument().openFindDialog(); }
    void repeatFind() { commandDocument().repeatFind(); }
    void sortRecords(bool descending) {
        auto& document=commandDocument();
        if(document.layout_.kind!=linaccess::AccessDocumentKind::Form || !document.dataSourceAvailable_) throw linaccess::VbaRuntimeError(2046,"Sorting requires an active bound form");
        const auto* focus=QApplication::focusWidget();
        for(const auto& control:document.layout_.controls) {
            const auto* widget=document.controls_.value(control.name.toCaseFolded()).data();
            if(!widget || !focus || (widget!=focus && !widget->isAncestorOf(focus))) continue;
            const auto column=columnIndex(control.controlSource,document.dataColumns_);
            if(column<0) continue;
            auto name=document.dataColumns_[column];
            if(document.dataTypes_.value(name.toCaseFolded()).contains("BLOB",Qt::CaseInsensitive)) throw linaccess::VbaRuntimeError(2046,"Binary fields cannot be sorted by this command");
            name.replace(']',"]]"); document.setOrderState('['+name+']'+(descending?" DESC":" ASC"),true); return;
        }
        throw linaccess::VbaRuntimeError(2046,"Sorting requires an active bound field control");
    }
    void selectObject(int type,const QString& name,bool inDatabase) {
        if(type<0 || type>3) throw linaccess::VbaRuntimeError(3251,"SelectObject object type is not implemented");
        if(!inDatabase && name.isEmpty()) throw linaccess::VbaRuntimeError(450,"SelectObject requires an open object name");
        QWidget* owner=nullptr;
        for(auto* ancestor=preview_.parentWidget();ancestor;ancestor=ancestor->parentWidget())
            if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
        const auto selectDatabase=[&](bool activate) {
            int error=0;
            if(!owner || !QMetaObject::invokeMethod(owner,"selectDatabaseObject",Qt::DirectConnection,Q_RETURN_ARG(int,error),Q_ARG(int,type),Q_ARG(QString,name),Q_ARG(bool,activate)))
                throw linaccess::VbaRuntimeError(3251,"Native database-window selection is unavailable");
            if(error) throw linaccess::VbaRuntimeError(error,"Database object selection failed");
        };
        QPointer<AccessFormPreview> selected;
        if(inDatabase) selectDatabase(false);
        else {
            if(type==0) {
                if(QApplication::activeModalWidget()) throw linaccess::VbaRuntimeError(2046,"SelectObject is blocked by a modal window");
                int error=0;
                if(!owner || owner->metaObject()->indexOfMethod("activateOpenTable(QString)")<0
                    || !QMetaObject::invokeMethod(owner,"activateOpenTable",Qt::DirectConnection,Q_RETURN_ARG(int,error),Q_ARG(QString,name)))
                    throw linaccess::VbaRuntimeError(3251,"Native open-table selection is unavailable");
                if(error) throw linaccess::VbaRuntimeError(error,"Open table selection failed");
                preview_.setProperty("la_explicit_table_context",true);
                return;
            }
            if(type!=2 && type!=3) throw linaccess::VbaRuntimeError(3251,"Open query window selection is not implemented");
            for(auto* widget:QApplication::topLevelWidgets()) {
                auto* candidate=dynamic_cast<AccessFormPreview*>(widget);
                if(!candidate || !candidate->isWindow() || !candidate->opened_ || candidate->closed_ || candidate->openCancelled_
                    || (candidate->layout_.kind==linaccess::AccessDocumentKind::Form ? 2 : 3)!=type
                    || candidate->db_.connectionName()!=preview_.db_.connectionName() || candidate->db_.databaseName()!=preview_.db_.databaseName()
                    || candidate->formName_.compare(name,Qt::CaseInsensitive)!=0) continue;
                if(selected) throw linaccess::VbaRuntimeError(5,"SelectObject target is ambiguous"); selected=candidate;
            }
            if(!selected) throw linaccess::VbaRuntimeError(type==3 ? 2451 : 2450,"SelectObject target is not open");
            if(QApplication::activeModalWidget() && QApplication::activeModalWidget()!=selected)
                throw linaccess::VbaRuntimeError(2046,"SelectObject is blocked by a modal window");
        }
        auto* current=activePreviewDocument();
        if(current && current!=selected && !current->closed_ && !current->openCancelled_
            && current->db_.connectionName()==preview_.db_.connectionName() && current->db_.databaseName()==preview_.db_.databaseName()
            && !current->commitControls()) throw linaccess::VbaRuntimeError(2501,"SelectObject control validation cancelled");
        if(inDatabase) { selectDatabase(true); return; }
        if(!selected || selected->closed_ || selected->closing_) throw linaccess::VbaRuntimeError(2501,"SelectObject target closed during validation");
        if(selected->isMinimized()) selected->showNormal(); else selected->show();
        selected->raise(); selected->activateWindow(); selected->activateDocument();
    }
    int doEvents() override {
        ++nativeExecutionDepth();
        const auto leave=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
        QCoreApplication::processEvents(QEventLoop::AllEvents);
        linaccess::drainNativeKeys();
        if(linaccess::vbaExecutionStopRequested()) throw linaccess::VbaExecutionStopped{};
        return 0;
    }
    void quitApplication() override { AccessFormPreview::quitApplication(); }
    void maximizeForm() override { commandDocument().window()->showMaximized(); }
    void minimizeForm() { commandDocument().window()->showMinimized(); }
    void restoreForm() { auto* window=commandDocument().window(); window->showNormal(); window->activateWindow(); }
    bool tableNavigation(int action,int offset,const QString& name,bool explicitTable) {
        if(!explicitTable && !usesActiveTableContext()) return false;
        QWidget* owner=nullptr;
        for(auto* ancestor=preview_.parentWidget();ancestor;ancestor=ancestor->parentWidget())
            if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
        if(!owner || owner->metaObject()->indexOfMethod("navigateTableDocument(QString,int,int,bool)")<0) return false;
        QVariantMap result;
        if(!QMetaObject::invokeMethod(owner,"navigateTableDocument",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),
            Q_ARG(QString,name),Q_ARG(int,action),Q_ARG(int,offset),Q_ARG(bool,explicitTable)))
            throw linaccess::VbaRuntimeError(3251,"Native table navigation dispatch is unavailable");
        if(!result.value("handled").toBool()) return false;
        const int error=result.value("error").toInt();
        if(error) throw linaccess::VbaRuntimeError(error,result.value("description").toString().toStdString());
        preview_.setProperty("la_explicit_table_context",true);
        return true;
    }
    void gotoRecord(int action,int offset=1,const QString& name={}) {
        if(name.isEmpty() && tableNavigation(action,offset,name,false)) return;
        AccessFormPreview* selected=nullptr;
        if (name.isEmpty()) selected=&commandDocument();
        else {
            for (auto* widget:QApplication::topLevelWidgets()) {
                auto* candidate=dynamic_cast<AccessFormPreview*>(widget);
                if (!candidate || !candidate->opened_ || candidate->closed_ || candidate->openCancelled_
                    || candidate->layout_.kind!=linaccess::AccessDocumentKind::Form
                    || candidate->db_.connectionName()!=preview_.db_.connectionName()
                    || candidate->db_.databaseName()!=preview_.db_.databaseName()
                    || candidate->formName_.compare(name,Qt::CaseInsensitive)!=0) continue;
                if (selected) throw linaccess::VbaRuntimeError(5,"GoToRecord target is ambiguous");
                selected=candidate;
            }
            if (!selected) throw linaccess::VbaRuntimeError(2450,"GoToRecord form is unavailable");
        }
        auto& document=*selected;
        if (document.recordCommandRunning_) throw linaccess::VbaRuntimeError(3251,"Navigation during record command unsupported");
        if (document.layout_.kind==linaccess::AccessDocumentKind::Report) throw linaccess::VbaRuntimeError(438,"Report does not support record navigation");
        if (action<0 || action>5) throw linaccess::VbaRuntimeError(5,"Invalid record navigation action");
        if ((action==0 || action==1 || action==4) && offset<1) throw linaccess::VbaRuntimeError(2105,"Invalid record navigation offset");
        if (action==5 && (!document.layout_.allowAdditions || !document.directTable_ || !document.dataSourceAvailable_))
            throw linaccess::VbaRuntimeError(2046,"New records are unavailable");
        if (!document.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Record navigation save cancelled");
        const qint64 count=document.recordSelector_->count()-1;
        const qint64 current=document.currentRecordIndex_ ? document.currentRecordIndex_ : count+1;
        qint64 target=0;
        if (action==0) target=current-offset;
        else if (action==1) target=current+offset;
        else if (action==2) target=1;
        else if (action==3) target=count;
        else if (action==4) target=offset;
        else target=count+1;
        const bool newRecord=target==count+1 && (action==1 || action==5);
        if (newRecord && !document.layout_.allowAdditions) throw linaccess::VbaRuntimeError(2046,"New records are disabled");
        if ((!newRecord && (target<1 || target>count)) || target>count+1) throw linaccess::VbaRuntimeError(2105,"Cannot go to specified record");
        const auto index=newRecord ? 0 : static_cast<int>(target);
        document.selectRecord(index);
        if (document.currentRecordIndex_!=index) throw linaccess::VbaRuntimeError(2501,"Record navigation cancelled");
    }
    void gotoLastRecord() override { gotoRecord(3); }
    void gotoNextRecord() override { gotoRecord(1); }
    void gotoNewRecord() override { gotoRecord(5); }
    void requery() override {
        if(usesActiveTableContext()) {
            QWidget* owner=nullptr;
            for(auto* ancestor=preview_.parentWidget();ancestor;ancestor=ancestor->parentWidget())
                if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
            if(owner && owner->metaObject()->indexOfMethod("requeryTableDocument()")>=0) {
                QVariantMap result;
                if(!QMetaObject::invokeMethod(owner,"requeryTableDocument",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result)))
                    throw linaccess::VbaRuntimeError(3251,"Native table Requery dispatch is unavailable");
                if(result.value("handled").toBool()) {
                    if(const int error=result.value("error").toInt())
                        throw linaccess::VbaRuntimeError(error,result.value("description").toString().toStdString());
                    return;
                }
            }
        }
        (void)makePreviewFormObject(commandDocument())->member("Requery",{});
    }
    void requeryNamedControl(std::string_view name) { commandDocument().requeryControl(name); }
    std::string memberControlName(std::string_view name) override {
        requireInstance();
        return preview_.controlMemberName(QString::fromUtf8(name.data(),name.size())).toUtf8().toStdString();
    }
    void requeryControl(std::string_view name) override { requireInstance(); preview_.requeryControl(name); }

    void focusControl(std::string_view name) override {
        requireInstance();
        (void)makePreviewControlObject(preview_,QString::fromUtf8(name.data(),name.size()))->member("SetFocus",{});
    }
    void setControlVisible(std::string_view name,bool value) override {
        requireInstance();
        makePreviewControlObject(preview_,QString::fromUtf8(name.data(),name.size()))->assign("Visible",linaccess::VbaValue::logical(value));
    }
    std::optional<std::string> moduleForProcedure(std::string_view procedure) override {
        QSqlQuery query(preview_.db_);
        if (!query.exec("SELECT name,source_bytes FROM _la_access_artifacts WHERE kind='vba' ORDER BY id DESC")) return std::nullopt;
        QSet<QString> names; std::optional<std::string> found;
        const auto target=QString::fromUtf8(procedure.data(),procedure.size());
        while (query.next()) {
            const auto name=query.value(0).toString().toCaseFolded(); if (names.contains(name)) continue; names.insert(name);
            const auto source=linaccess::decodeVbaSource(query.value(1).toByteArray());
            const auto code=source.toUtf8().toStdString();
            if (linaccess::isVbaClassModule(code)) continue;
            const auto analysis=linaccess::analyzeVba(code); const auto lines=source.split('\n');
            for (const auto& item:analysis.procedures) {
                if (QString::fromStdString(item.name).compare(target,Qt::CaseInsensitive)!=0) continue;
                if (item.line>0 && item.line<=lines.size() && QRegularExpression("^Private\\b",QRegularExpression::CaseInsensitiveOption).match(lines[item.line-1].trimmed()).hasMatch()) continue;
                if (found) throw linaccess::VbaRuntimeError(5,"Ambiguous public procedure");
                found=code;
            }
        }
        return found;
    }
    std::shared_ptr<linaccess::VbaObject> applicationObject() override { return makePreviewApplicationObject(preview_); }
    std::shared_ptr<linaccess::VbaObject> currentForm() override { requireInstance(); return makePreviewFormObject(preview_); }
    std::shared_ptr<linaccess::VbaObject> reportsCollection() override { return makePreviewReportsObject(preview_); }
    std::shared_ptr<linaccess::VbaObject> formsCollection() override { return makePreviewFormsObject(preview_); }
    linaccess::VbaValue formatValue(const linaccess::VbaValue& value,std::string_view mask,bool stringResult) override {
        return linaccess::formatAccessValue(value,mask,stringResult);
    }
    std::shared_ptr<linaccess::VbaObject> databaseEngine() override { return linaccess::makeDaoEngine(preview_.db_); }
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(preview_.db_); }
    void executeSql(std::string_view statement) override { runSql(statement,true); }
    void runSql(std::string_view statement,bool useTransaction) {
        if (!preview_.db_.isOpen()) throw linaccess::VbaRuntimeError(3420,"No open native database project");
        const auto execute=[&] {
            auto database=linaccess::makeDaoDatabase(preview_.db_);
            try { (void)database->member("Execute",{linaccess::VbaValue::string(std::string(statement))}); }
            catch(const linaccess::VbaRuntimeError& failure) {
                if(failure.number==3065) throw linaccess::VbaRuntimeError(2342,"RunSQL requires an action SQL statement");
                throw;
            }
        };
        if(!useTransaction) { execute(); return; }
        static thread_local quint64 sequence=0;
        const auto savepoint=QString("la_docmd_runsql_%1").arg(++sequence);
        QSqlQuery transaction(preview_.db_);
        const auto failure=[&] {
            const int code=transaction.lastError().nativeErrorCode().toInt() & 255;
            throw linaccess::VbaRuntimeError(code==5 || code==6 ? 3260 : 3146,"Native RunSQL transaction operation failed");
        };
        if(!transaction.exec("SAVEPOINT "+savepoint)) failure();
        try {
            execute();
            if(!transaction.exec("RELEASE "+savepoint)) failure();
        } catch(...) {
            const bool rolledBack=transaction.exec("ROLLBACK TO "+savepoint);
            const bool released=transaction.exec("RELEASE "+savepoint);
            if(!rolledBack || !released) throw linaccess::VbaRuntimeError(3146,"Native RunSQL transaction cleanup failed");
            throw;
        }
    }
    void sendKeys(std::string_view keys,bool wait) override {
        ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
        QPointer<QWidget> owner;
        for(auto* ancestor=preview_.parentWidget();ancestor;ancestor=ancestor->parentWidget())
            if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
        const auto tableContext=[owner]() -> QVariantMap {
            QVariantMap result;
            if(owner && owner->metaObject()->indexOfMethod("tableKeyboardContext()")>=0)
                (void)QMetaObject::invokeMethod(owner,"tableKeyboardContext",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result));
            return result;
        };
        if(usesActiveTableContext() && AccessFormPreview::databaseWindowActive() && !tableContext().value("handled").toBool())
            throw linaccess::VbaRuntimeError(3251,"Database-window SendKeys target is not implemented");
        const auto connection=preview_.db_.connectionName(),database=preview_.db_.databaseName();
        linaccess::sendNativeKeys(&preview_,QString::fromUtf8(keys.data(),keys.size()),wait,
            [preview=QPointer<AccessFormPreview>(&preview_)](int number) {
                if (preview && preview->status_) preview->status_->setText("SendKeys: Fehler "+QString::number(number));
            },[context=QPointer<AccessFormPreview>(&preview_),tableContext,connection,database]() -> QWidget* {
                if (!context || context->closed_ || context->openCancelled_) return nullptr;
                const auto table=tableContext();
                if(table.value("handled").toBool()) {
                    if(table.value("connection").toString()!=connection || table.value("database").toString()!=database) return nullptr;
                    return qobject_cast<QWidget*>(table.value("target").value<QObject*>());
                }
                if ((context->hasActivated_ || context->property("la_explicit_table_context").toBool()) && AccessFormPreview::databaseWindowActive()) return nullptr;
                auto* document=activePreviewDocument();
                if (!context->opened_ || !document || document->closed_ || document->openCancelled_
                    || document->db_.connectionName()!=connection || document->db_.databaseName()!=database) document=context;
                if(document->db_.connectionName()!=connection || document->db_.databaseName()!=database) return nullptr;
                auto* focus=QApplication::focusWidget(); auto* parent=focus;
                while(parent && parent!=document) parent=parent->parentWidget();
                return parent ? focus : nullptr;
            });
    }
    void openFormArguments(const std::vector<linaccess::VbaValue>& args) override {
        const auto request=linaccess::parseVbaOpenFormArguments(args);
        preview_.openOtherForm(QString::fromStdString(request.name),QString::fromStdString(request.where),request.openArgs ? std::optional<QString>(QString::fromStdString(*request.openArgs)) : std::nullopt,request.windowMode,request.dataMode);
    }
    void openForm(std::string_view name) override {
        preview_.openOtherForm(QString::fromUtf8(name.data(), name.size()));
    }
private:
    AccessFormPreview& commandDocument() {
        if (usesActiveTableContext() && AccessFormPreview::databaseWindowActive()) throw linaccess::VbaRuntimeError(2046,"Record command requires an active native document");
        auto* document=activePreviewDocument();
        if ((preview_.opened_ || usesActiveTableContext()) && document && !document->closed_ && !document->openCancelled_
            && document->db_.connectionName()==preview_.db_.connectionName() && document->db_.databaseName()==preview_.db_.databaseName()) return *document;
        if(usesActiveTableContext() && preview_.isHidden()) throw linaccess::VbaRuntimeError(2046,"Record command requires an active native document");
        return preview_;
    }
    std::vector<bool> moduleScopes_;
    static bool recordProperty(std::string_view name) {
        const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        return key=="newrecord" || key=="dataentry" || key=="modal" || key=="popup" || key=="openargs" || key=="currentrecord" || key=="orderby" || key=="orderbyon";
    }
    bool instanceScope() const { return moduleScopes_.empty() || moduleScopes_.back(); }
    void requireInstance() const { if (!instanceScope()) throw linaccess::VbaRuntimeError(438,"Me is unavailable in a standard module"); }
    AccessFormPreview& preview_;
};

linaccess::VbaValue dispatchPreviewDocumentCommand(AccessFormPreview& preview,std::string_view command,const std::vector<linaccess::VbaValue>& actuals) {
    using linaccess::VbaRuntimeError; using linaccess::VbaValue;
    ++nativeExecutionDepth();
    const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    std::vector<VbaValue> args;
    for (const auto& actual:actuals) {
        auto value=linaccess::scalarVbaValue(actual);
        if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null document command argument");
        args.push_back(std::move(value));
    }
    const auto number=[](const VbaValue& value) {
        const auto result=linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
        if (result<std::numeric_limits<std::int32_t>::min() || result>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"Document command argument overflow");
        return static_cast<int>(result);
    };
    PreviewVbaHost host(preview);
    if(command=="sendkeys") {
        if(args.empty() || args.size()>2 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"SendKeys argument count invalid");
        const bool wait=args.size()>1 && args[1].kind!=VbaValue::Kind::Missing ? linaccess::coerceVbaModuleValue(args[1],"boolean").boolean : false;
        host.sendKeys(args[0].asText(),wait); return {};
    }
    if(command=="findrecord") {
        if(args.empty() || args.size()>7 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"FindRecord argument count invalid");
        const auto optionalNumber=[&](std::size_t index,int fallback) { return index>=args.size() || args[index].kind==VbaValue::Kind::Missing ? fallback : number(args[index]); };
        const auto optionalBoolean=[&](std::size_t index,bool fallback) { return index>=args.size() || args[index].kind==VbaValue::Kind::Missing ? fallback : linaccess::coerceVbaModuleValue(args[index],"boolean").boolean; };
        const int match=optionalNumber(1,1),search=optionalNumber(3,2),field=optionalNumber(5,-1);
        const bool sensitive=optionalBoolean(2,false),formatted=optionalBoolean(4,false),first=optionalBoolean(6,true);
        if(match<0 || match>2 || search<0 || search>2 || (field!=0 && field!=-1)) throw VbaRuntimeError(5,"Invalid FindRecord option");
        if(field!=-1 || formatted) throw VbaRuntimeError(3251,"FindRecord all-field/formatted search is not implemented");
        const auto sought=linaccess::normalizeNativeFindValue(args[0]);
        const bool textValue=sought.kind==VbaValue::Kind::Text;
        const auto text=linaccess::nativeFindValueText(sought);
        if(textValue && (text.isEmpty() || text.contains('*') || text.contains('?') || text.contains('#') || text.contains('[') || text.contains(']')))
            throw VbaRuntimeError(3251,"Empty/wildcard FindRecord is not implemented");
        if(!textValue && match!=1) throw VbaRuntimeError(3251,"Typed FindRecord currently requires whole-field matching");
        if(!host.findTableRecord(sought,match,sensitive,first,search)) host.findFormRecord(sought,match,sensitive,first,search);
        return {};
    }
    if(command=="selectobject") {
        if(args.empty() || args.size()>3 || args[0].kind==VbaValue::Kind::Empty) throw VbaRuntimeError(450,"SelectObject argument count invalid");
        const auto name=args.size()>1 && args[1].kind!=VbaValue::Kind::Empty ? QString::fromStdString(args[1].asText()) : QString{};
        const bool inDatabase=args.size()>2 && args[2].kind!=VbaValue::Kind::Empty ? linaccess::coerceVbaModuleValue(args[2],"boolean").boolean : false;
        host.selectObject(number(args[0]),name,inDatabase); return {};
    }
    if (command=="domenuitem") {
        if (args.size()<3 || args.size()>5) throw VbaRuntimeError(450,"DoMenuItem argument count invalid");
        for (std::size_t i=0;i<3;++i) if (args[i].kind==VbaValue::Kind::Empty) throw VbaRuntimeError(450,"DoMenuItem required argument omitted");
        const int version=args.size()==5 && args[4].kind!=VbaValue::Kind::Empty ? number(args[4]) : 11;
        if (version!=70 || number(args[0])!=0 || number(args[1])!=1 || (args.size()>3 && args[3].kind!=VbaValue::Kind::Empty))
            throw VbaRuntimeError(3251,"Legacy menu context unsupported");
        if (number(args[2])==8) host.selectCurrentRecord();
        else if (number(args[2])==6) host.deleteCurrentRecord();
        else if (number(args[2])==2) host.copyCurrentRecord();
        else if (number(args[2])==5) host.pasteAppendRecord();
        else throw VbaRuntimeError(3251,"Legacy menu action unsupported");
        return {};
    }
    if (command=="runsql") {
        if (args.empty() || args.size()>2) throw VbaRuntimeError(450,"RunSQL argument count invalid");
        const bool useTransaction=args.size()<2 || args[1].kind==VbaValue::Kind::Empty
            ? true : coerceVbaModuleValue(scalarVbaValue(args[1]),"boolean").boolean;
        const auto statement=scalarVbaValue(args[0]);
        if(statement.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null RunSQL statement");
        host.runSql(statement.asText(),useTransaction); return {};
    }
    if (command=="runcommand") {
        if (args.size()!=1) throw VbaRuntimeError(450,"RunCommand requires one command");
        const auto code=number(args[0]);
        if(host.tableCommand(code)) return {};
        if (code==28) host.gotoRecord(5);
        else if (code==30) host.findRecords();
        else if (code==341) host.repeatFind();
        else if (code==179) host.zoomField();
        else if(code==163 || code==164) host.sortRecords(code==164);
        else if (code>=65 && code<=68) { const int actions[]{1,0,2,3}; host.gotoRecord(actions[code-65]); }
        else if (code==97) host.saveRecord();
        else if (code==50) host.selectCurrentRecord();
        else if (code==223) host.deleteCurrentRecord();
        else if (code==190) host.copyCurrentRecord();
        else if (code==38) host.pasteAppendRecord();
        else throw VbaRuntimeError(3251,"RunCommand option unsupported");
        return {};
    }
    if (command=="maximize" || command=="minimize" || command=="restore") {
        if (!args.empty()) throw VbaRuntimeError(450,"Window command takes no arguments");
        if(command=="maximize") host.maximizeForm();
        else if(command=="minimize") host.minimizeForm();
        else host.restoreForm();
        return {};
    }
    if (command=="requery") {
        if (args.size()>1) throw VbaRuntimeError(450,"Requery argument count invalid");
        if (!args.empty() && args[0].kind!=VbaValue::Kind::Empty && !args[0].asText().empty()) host.requeryNamedControl(args[0].asText());
        else host.requery();
        return {};
    }
    if (command=="gotorecord") {
        if (args.size()>4) throw VbaRuntimeError(450,"GoToRecord argument count invalid");
        const auto type=!args.empty() && args[0].kind!=VbaValue::Kind::Empty ? number(args[0]) : -1;
        const auto name=args.size()>1 && args[1].kind!=VbaValue::Kind::Empty ? QString::fromStdString(args[1].asText()) : QString{};
        if ((type!=-1 && type!=0 && type!=2) || (!name.isEmpty() && type!=0 && type!=2)) throw VbaRuntimeError(3251,"GoToRecord object type unsupported");
        const auto action=args.size()>2 && args[2].kind!=VbaValue::Kind::Empty ? number(args[2]) : 1;
        const auto offset=args.size()>3 && args[3].kind!=VbaValue::Kind::Empty ? number(args[3]) : 1;
        if(type==0) {
            if(!host.tableNavigation(action,offset,name,true)) throw VbaRuntimeError(3251,"Native table navigation is unavailable");
        } else host.gotoRecord(action,offset,name);
        return {};
    }
    throw VbaRuntimeError(438,"Document command unsupported");
}

AccessFormPreview::AccessFormPreview(QString formName, const QByteArray& source,
                                     const QByteArray& vbaSource, QSqlDatabase database,
                                     QWidget* parent,bool embedded,QString whereCondition,std::optional<QString> openArgs,int windowMode,int dataMode)
    : QDialog(parent), formName_(std::move(formName)),openArgs_(std::move(openArgs)),
      vba_(linaccess::decodeVbaSource(vbaSource)),
      layout_(linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(source))),
      db_(std::move(database)) {
    linaccess::VbaExecutionScope executionScope;
    const auto stopped=[&] {
        if(!linaccess::vbaExecutionStopRequested()) return false;
        openCancelled_=true; return true;
    };
    if(stopped()) return;
    if(windowMode<0 || windowMode>3) throw linaccess::VbaRuntimeError(3251,"Native window mode is not implemented");
    if(dataMode<-1 || dataMode>2) throw linaccess::VbaRuntimeError(3251,"Native data mode is not implemented");
    if(dataMode!=-1 && layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(3251,"DataMode requires a form");
    if(dataMode==0) { layout_.dataEntry=true; layout_.allowAdditions=true; }
    if(dataMode==1 || dataMode==2) {
        if(layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(3251,"DataMode requires a form");
        if(dataMode==1) layout_.dataEntry=false;
        layout_.allowEdits=dataMode==1; layout_.allowAdditions=dataMode==1; layout_.allowDeletions=dataMode==1;
    }
    if(layout_.dataEntry && layout_.allowAdditions && !layout_.recordSource.trimmed().isEmpty()
        && (layout_.kind!=linaccess::AccessDocumentKind::Form || !db_.tables().contains(layout_.recordSource.trimmed(),Qt::CaseInsensitive)))
        throw linaccess::VbaRuntimeError(3251,"Native DataEntry requires a direct table source");
    nativeDialog_=windowMode==3;
    if(nativeDialog_) { layout_.modal=true; layout_.popup=true; }
    static quint64 nextWindowSequence=0;
    windowSequence_=++nextWindowSequence;
    static thread_local QStringList constructionStack;
    const auto identity=db_.connectionName()+QChar(0x1f)+formName_.toCaseFolded();
    if (constructionStack.contains(identity) || constructionStack.size()>=16)
        throw linaccess::VbaRuntimeError(2455,"Recursive subform source");
    QScopedValueRollback<QStringList> constructionGuard(constructionStack);
    constructionStack.append(identity);
    if (embedded) setWindowFlags(Qt::Widget);
    else if (layout_.modal) setWindowModality(Qt::ApplicationModal);
    filter_=whereCondition.trimmed().isEmpty() ? layout_.filter : whereCondition;
    filterOn_=whereCondition.trimmed().isEmpty() ? layout_.filterOnLoad && !filter_.trimmed().isEmpty() : true;
    orderBy_=layout_.orderBy; orderByOn_=layout_.orderByOnLoad && !orderBy_.trimmed().isEmpty();
    initialCurrentDeferred_=!embedded && windowMode!=1 && eventProcedure(layout_.events.value("onactivate"));
    if (db_.isOpen()) connectPreviewSqlObjects(db_);
    static qlonglong openingSequence = 0;
    setProperty("la_form_open_order",++openingSequence);
    const bool queryViewsReady = linaccess::installAccessQueryViews(db_) >= 0;
    setWindowTitle((layout_.caption.isEmpty() ? formName_ : layout_.caption) + " — LinAccess");
    setProperty("la_access_tag",layout_.tag);
    resize(950, 720);
    auto* outer = new QVBoxLayout(this);
    if (layout_.kind==linaccess::AccessDocumentKind::Report) buildPrintToolbar(outer);
    auto* controls = new QHBoxLayout();
    controls->addWidget(new QLabel("Datensatz", this));
    recordSelector_ = new QComboBox(this);
    recordSelector_->setObjectName("access-record-selector");
    recordSelector_->installEventFilter(this);
    controls->addWidget(recordSelector_, 1);
    filterEditor_=new QLineEdit(this); filterEditor_->setObjectName("access-filter");
    filterEditor_->setPlaceholderText("Bedingung eingeben"); filterEditor_->setText(filter_);
    controls->addWidget(filterEditor_,1);
    filterToggle_=new QCheckBox("Filter aktiv",this); filterToggle_->setObjectName("access-filter-on"); filterToggle_->setChecked(filterOn_);
    controls->addWidget(filterToggle_);
    auto* save = new QPushButton("Datensatz speichern", this);
    save->setObjectName("access-save-record");
    controls->addWidget(save);
    outer->addLayout(controls);

    auto* scroll = new QScrollArea(this);
    scroll->setWidgetResizable(false);
    auto* canvas = new QWidget(scroll);
    canvas->setObjectName("access-form-canvas");
    canvas->setAutoFillBackground(true);
    auto canvasPalette=canvas->palette();
    canvasPalette.setColor(QPalette::Window,QColor(240,240,240));
    canvas->setPalette(canvasPalette);
    canvas->setFixedSize(qBound(400, pixels(layout_.width) + 70, 3500),
                         qBound(250, pixels(layout_.height) + 90, 3500));
    buildControls(canvas);
    scroll->setWidget(canvas);
    outer->addWidget(scroll, 1);

    status_ = new QLabel(this);
    status_->setWordWrap(true);
    const auto applyFilter=[this] {
        try { setFilterState(filterEditor_->text(),filterToggle_->isChecked()); }
        catch(const linaccess::VbaRuntimeError& error) {
            const QSignalBlocker blocked(filterToggle_); filterToggle_->setChecked(filterOn_);
            status_->setText("Filter nicht angewendet ("+QString::number(error.number)+")");
        }
    };
    connect(filterToggle_,&QCheckBox::toggled,this,[applyFilter](bool) { applyFilter(); });
    connect(filterEditor_,&QLineEdit::editingFinished,this,applyFilter);
    outer->addWidget(status_);
    connect(recordSelector_, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, [this](int index) { selectRecord(index); });
    connect(save, &QPushButton::clicked, this, [this] { (void)saveRecord(); });
    if (layout_.kind==linaccess::AccessDocumentKind::Form) {
        formTimer_=new QTimer(this); formTimer_->setTimerType(Qt::PreciseTimer);
        connect(formTimer_,&QTimer::timeout,this,[this] {
            if (!opened_ || closed_ || closing_ || openCancelled_ || timerDispatching_ || loading_
                || saving_ || navigating_ || recordCommandRunning_ || insertionEventRunning_) return;
            ++nativeExecutionDepth();
            const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
            QScopedValueRollback<bool> dispatching(timerDispatching_,true);
            const auto binding=layout_.events.value("ontimer");
            if (eventProcedure(binding)) (void)formEvent("OnTimer","Form_Timer");
            else if (!binding.trimmed().isEmpty()) {
                PreviewVbaHost host(*this);
                const auto result=linaccess::executeVbaEventExpression(vba_.toUtf8().toStdString(),binding.toUtf8().toStdString(),host);
                eventResults_.insert("ontimer",result);
                if (failed(result) && status_) status_->setText("Timer-Ereignisfehler "+QString::number(result.errorNumber));
            }
        });
    }
    if (!queryViewsReady) {
        status_->setText("Abfrageansichten werden noch verwendet; andere Ansichten schließen und Formular erneut öffnen.");
        canvas->setEnabled(false); save->setEnabled(false); return;
    }
    const bool report = layout_.kind==linaccess::AccessDocumentKind::Report;
    if (!report) loadRecords();
    for(const auto& control:layout_.controls)
        if(auto* list=dynamic_cast<AccessListBox*>(controls_.value(control.name.toCaseFolded()).data())) {
            try { requeryControl(control.name.toUtf8().toStdString()); }
            catch(const std::exception& error) {
                const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error);
                const int number=typed ? typed->number : 5;
                list->setProperty("la_rows_error",number);
                list->setProperty("la_list_source_tooltip",list->toolTip());
                list->setToolTip("Listendaten nicht verfügbar ("+QString::number(number)+")");
                status_->setText(list->toolTip());
            }
        }
    if(stopped()) return;
    const auto outcome = formEvent("OnOpen","Form_Open");
    if(stopped()) return;
    if (outcome.cancelled) { openCancelled_ = true; return; }
    if (report) {
        loadRecords();
        if (dataSourceAvailable_ && recordRows_.isEmpty()) {
            const auto noData = formEvent("OnNoData","Report_NoData");
            if (noData.cancelled) { openCancelled_ = true; return; }
        }
    }
    (void)formEvent("OnLoad","Form_Load");
    if(stopped()) return;
    opened_ = true;
    if (nativeActivationPending_) activateDocument();
    if (!initialCurrentDeferred_) initialCurrent();
    if(stopped()) return;
    if (!report) setTimerInterval(layout_.timerInterval);
}

AccessFormPreview* AccessFormPreview::showDocument(linaccess::AccessDocumentKind kind,QString name,
    const QByteArray& source,const QByteArray& vba,QSqlDatabase database,QWidget* projectOwner) {
    if (closingProjects().contains(projectKey(database))) throw linaccess::VbaRuntimeError(2501,"Project is closing");
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    QPointer<AccessFormPreview> existing;
    for(auto* widget:QApplication::topLevelWidgets()) {
        auto* document=dynamic_cast<AccessFormPreview*>(widget);
        if(!document || !document->opened_ || document->closed_ || document->openCancelled_ || document->layout_.kind!=kind
            || document->db_.connectionName()!=database.connectionName() || document->db_.databaseName()!=database.databaseName()
            || document->formName_.compare(name,Qt::CaseInsensitive)!=0) continue;
        if(existing) throw linaccess::VbaRuntimeError(5,"Document target instance is ambiguous");
        existing=document;
    }
    if(existing) {
        if(existing->closing_) throw linaccess::VbaRuntimeError(2501,"Document is closing");
        existing->show(); existing->raise(); existing->activateWindow(); existing->activateDocument();
        return existing && !existing->closed_ ? existing.data() : nullptr;
    }
    if(linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(source)).kind!=kind)
        throw linaccess::VbaRuntimeError(kind==linaccess::AccessDocumentKind::Report ? 2103 : 2102,"Document target kind mismatch");
    QWidget* owner=nullptr;
    for(auto* ancestor=projectOwner;ancestor;ancestor=ancestor->parentWidget())
        if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
    auto document=std::make_unique<AccessFormPreview>(std::move(name),source,vba,std::move(database),owner);
    if(document->openCancelled_ || document->closed_) return nullptr;
    document->show(); document->activateWindow(); document->activateDocument();
    if(document->openCancelled_ || document->closed_) return nullptr;
    QObject::connect(document.get(),&QDialog::finished,document.get(),[form=document.get()] { deferFormDisposal(form); });
    return document.release();
}

void AccessFormPreview::setTimerInterval(int interval) {
    if (layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Timer requires a form");
    if (interval<0) throw linaccess::VbaRuntimeError(5,"Invalid TimerInterval");
    layout_.timerInterval=interval;
    if (!formTimer_) return;
    if (interval && opened_ && !closed_ && !openCancelled_) formTimer_->start(interval);
    else formTimer_->stop();
}

void AccessFormPreview::initialCurrent() {
    if (initialCurrentDispatched_ || closed_ || openCancelled_) return;
    initialCurrentDispatched_=true;
    (void)formEvent("OnCurrent","Form_Current");
}

bool AccessFormPreview::event(QEvent* event) {
    const auto type=event->type();
    const auto result=QDialog::event(event);
    if(type==QEvent::Hide && isWindow() && isHidden() && opened_ && !closed_ && !closing_ && !openCancelled_ && activeDocument()==this) {
        ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
        activeDocument().clear();
        (void)formEvent("OnDeactivate","Form_Deactivate");
    }
    if (type==QEvent::WindowActivate && isWindow()) {
        nativeActivationPending_=true;
        if (opened_) activateDocument();
    }
    return result;
}

bool AccessFormPreview::databaseWindowActive(QWidget* window) { return databaseWindowContext().window && (!window || databaseWindowContext().window==window); }

bool AccessFormPreview::projectObjectLoaded(const QSqlDatabase& database,linaccess::AccessDocumentKind kind,const QString& name) {
    if(!database.isOpen()) return false;
    const auto key=projectKey(database);
    for(auto* widget:QApplication::allWidgets()) {
        auto* document=dynamic_cast<AccessFormPreview*>(widget);
        if(document && document->opened_ && !document->closed_ && !document->openCancelled_
            && document->layout_.kind==kind && document->formName_.compare(name,Qt::CaseInsensitive)==0 && projectKey(document->db_)==key) {
            bool liveOwner=true;
            for(auto* parent=document->parentWidget();parent;parent=parent->parentWidget())
                if(auto* ancestor=dynamic_cast<AccessFormPreview*>(parent);ancestor && (ancestor->closed_ || ancestor->openCancelled_)) { liveOwner=false; break; }
            if(liveOwner) return true;
        }
    }
    return false;
}

QVector<QPointer<AccessFormPreview>> AccessFormPreview::projectDocuments(const QSqlDatabase& database) {
    QVector<QPointer<AccessFormPreview>> documents;
    if(!database.isOpen()) return documents;
    const auto key=projectKey(database);
    for(auto* widget:QApplication::topLevelWidgets()) {
        auto* candidate=dynamic_cast<AccessFormPreview*>(widget);
        if(candidate && candidate->opened_ && !candidate->closed_ && !candidate->closing_ && !candidate->openCancelled_
            && candidate->isVisible() && !candidate->isModal() && projectKey(candidate->db_)==key)
            documents.append(candidate);
    }
    std::sort(documents.begin(),documents.end(),[](const auto& left,const auto& right) { return left->windowSequence_<right->windowSequence_; });
    return documents;
}

bool AccessFormPreview::activateProjectDocument(QSqlDatabase database,QWidget* projectWindow,AccessFormPreview* target) {
    if(!database.isOpen() || !projectWindow || QApplication::activeModalWidget()) return false;
    const auto key=projectKey(database);
    if(closingProjects().contains(key)) return false;
    const QPointer<AccessFormPreview> source=activeDocument(),selected=target;
    const QPointer<QWidget> owner=projectWindow;
    const bool databaseTarget=target==nullptr;
    if(source && (projectKey(source->db_)!=key || source->closed_ || source->closing_)) return false;
    if(selected && !projectDocuments(database).contains(selected)) return false;
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    if(source && (!source->commitControls() || source->closed_ || activeDocument()!=source)) return false;
    if(!owner || !database.isOpen() || projectKey(database)!=key || closingProjects().contains(key) || QApplication::activeModalWidget()) return false;
    if(databaseTarget) {
        (void)QMetaObject::invokeMethod(owner,"focusDatabaseWindow",Qt::DirectConnection);
        return databaseWindowActive(owner);
    }
    if(!selected || !projectDocuments(database).contains(selected)) return false;
    if(selected->isMinimized()) selected->showNormal();
    selected->raise(); selected->activateWindow(); selected->activateDocument();
    if(!selected || activeDocument()!=selected || selected->closed_) return false;
    if(auto* focus=selected->focusWidget(); focus && focus->isEnabled()) focus->setFocus(Qt::ShortcutFocusReason);
    return true;
}

void AccessFormPreview::cycleProjectDocuments(QSqlDatabase database,QWidget* projectWindow,bool backwards,AccessFormPreview* current) {
    if(!database.isOpen() || !projectWindow || QApplication::activeModalWidget()) return;
    const auto key=projectKey(database);
    if(closingProjects().contains(key)) return;
    if(current && (projectKey(current->db_)!=key || current->closed_ || current->closing_ || activeDocument()!=current)) return;
    if(projectWindow->metaObject()->indexOfMethod("cycleDocument(bool)")>=0
        && QMetaObject::invokeMethod(projectWindow,"cycleDocument",Qt::DirectConnection,Q_ARG(bool,backwards))) return;
    auto documents=projectDocuments(database);
    documents.prepend(nullptr); // Database window, then native documents in creation order.
    const int index=current ? documents.indexOf(current) : 0;
    if(index<0) return;
    const auto targetIndex=(index+(backwards ? documents.size()-1 : 1))%documents.size();
    (void)activateProjectDocument(database,projectWindow,documents[targetIndex]);
}

void AccessFormPreview::activateDatabaseWindow(QSqlDatabase database,QWidget* window) {
    if (!window || !database.isOpen() || QApplication::activeModalWidget()) return;
    auto& context=databaseWindowContext();
    if(context.window==window && context.project==projectKey(database)) return;
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    context={window,projectKey(database)};
    const auto previous=activeDocument(); activeDocument().clear();
    if(previous && previous->opened_ && !previous->closed_ && !previous->openCancelled_)
        (void)previous->formEvent("OnDeactivate","Form_Deactivate");
}

void AccessFormPreview::activateDocument() {
    if (!opened_ || closed_ || closing_ || openCancelled_ || !isWindow() || isHidden()) return;
    hasActivated_=true;
    databaseWindowContext()={};
    auto& active=activeDocument();
    if (active==this) return;
    const auto previous=active;
    active=this;
    // Only an actual transition between Access documents deactivates the
    // preceding one. Qt WindowDeactivate also covers external applications
    // and message boxes; those do not constitute this Access event.
    if (previous && previous->opened_ && !previous->closed_ && !previous->openCancelled_)
        (void)previous->formEvent("OnDeactivate","Form_Deactivate");
    if (active!=this || closed_) return;
    (void)formEvent("OnActivate","Form_Activate");
    if (active==this && !closed_) initialCurrent();
}

bool AccessFormPreview::supportsControlFont(const linaccess::AccessVisualControl& control) {
    return !(control.locked && control.type!="TextBox" && control.type!="ListBox") &&
        (control.type=="TextBox" || control.type=="Label" || control.type=="CommandButton"
         || control.type=="ComboBox" || control.type=="ListBox" || control.type=="CheckBox"
         || control.type=="OptionButton" || control.type=="ToggleButton");
}
void AccessFormPreview::applyControlFont(QWidget* widget,const linaccess::AccessVisualControl& control) {
    if (!supportsControlFont(control)) return;
    auto font=widget->font();
    if (control.font.name) font.setFamily(*control.font.name);
    font.setPointSize(control.font.size.value_or(control.type=="CommandButton" ? 10 : 8));
    const auto weight=control.font.weight.value_or(400);
    font.setWeight(static_cast<QFont::Weight>(weight ? weight : QFont::Normal));
    font.setItalic(control.font.italic.value_or(false));
    font.setUnderline(control.font.underline.value_or(false));
    widget->setFont(font);
    // Qt's editable combo editor keeps its own explicit default font.
    if (auto* combo=qobject_cast<QComboBox*>(widget)) {
        if (combo->lineEdit()) combo->lineEdit()->setFont(font);
        combo->view()->setFont(font);
    }
}

QColor AccessFormPreview::resolveOleColor(qint32 color) {
    const auto bits=static_cast<quint32>(color);
    if (bits<=0x00ffffffU) return QColor(bits&255U,(bits>>8)&255U,(bits>>16)&255U);
    // These source system-color intents cover the original foreground and
    // background values. Resolve from the application palette, never
    // from a previously overridden instance palette.
    const auto native=QApplication::palette();
    if (bits==0x80000008U) return native.color(QPalette::Active,QPalette::Text);
    if (bits==0x80000012U) return native.color(QPalette::Active,QPalette::ButtonText);
    if (bits==0x80000005U) return native.color(QPalette::Active,QPalette::Base);
    if (bits==0x8000000fU) return native.color(QPalette::Active,QPalette::Button);
    throw linaccess::VbaRuntimeError(3251,"OLE color adapter is not implemented for this encoding");
}

void AccessFormPreview::applyForeColor(QWidget* widget,const linaccess::AccessVisualControl& control) {
    if (!supportsControlFont(control) || !control.foreColor) return;
    const auto color=resolveOleColor(*control.foreColor);
    const auto apply=[&](QWidget* target) {
        target->ensurePolished();
        auto palette=target->palette();
        for (const auto group:{QPalette::Active,QPalette::Inactive})
            for (const auto role:{QPalette::Text,QPalette::WindowText,QPalette::ButtonText}) palette.setColor(group,role,color);
        target->setPalette(palette);
    };
    apply(widget);
    if (auto* combo=qobject_cast<QComboBox*>(widget)) {
        if (combo->lineEdit()) apply(combo->lineEdit());
        if (combo->view()) apply(combo->view());
    }
}

void AccessFormPreview::applyControlBackground(QWidget* widget,const linaccess::AccessVisualControl& control) {
    if (!supportsTextAlignment(control)) return;
    widget->ensurePolished();
    const auto role=control.type=="TextBox" ? QPalette::Base : QPalette::Window;
    const auto background=control.backColor ? resolveOleColor(*control.backColor) : QApplication::palette().color(QPalette::Active,role);
    auto palette=widget->palette();
    for(const auto group:{QPalette::Active,QPalette::Inactive}) palette.setColor(group,role,control.backStyle==0 ? QColor(Qt::transparent) : background);
    widget->setPalette(palette);
    if (auto* label=qobject_cast<QLabel*>(widget)) label->setAutoFillBackground(control.backStyle==1);
}

bool AccessFormPreview::supportsTextAlignment(const linaccess::AccessVisualControl& control) {
    return control.type=="TextBox" || (control.type=="Label" && !control.locked);
}

void AccessFormPreview::applyTextAlignment(QWidget* widget,const linaccess::AccessVisualControl& control) {
    if (!supportsTextAlignment(control)) return;
    if (control.textAlign==4) throw linaccess::VbaRuntimeError(3251,"Distributed text alignment is not implemented");
    Qt::Alignment horizontal=Qt::AlignLeft;
    if (control.textAlign==2) horizontal=Qt::AlignHCenter;
    else if (control.textAlign==3) horizontal=Qt::AlignRight;
    else if (control.textAlign==0 && control.type=="TextBox") {
        // Use retained value/field types, never guess from numeric-looking text
        // or evaluate expressions merely to choose a visual alignment.
        const auto type=widget->property("la_sql_type").toString().toUpper();
        const auto kind=static_cast<linaccess::VbaValue::Kind>(widget->property("la_vba_kind").toInt());
        const bool numericField=type.contains("INT") || type.contains("REAL") || type.contains("FLOA")
            || type.contains("DOUB") || type.contains("NUM") || type.contains("DEC") || type.contains("DATE");
        if ((!type.isEmpty() && numericField) || (type.isEmpty() && (kind==linaccess::VbaValue::Kind::Number || kind==linaccess::VbaValue::Kind::Currency)))
            horizontal=Qt::AlignRight;
    }
    if (auto* editor=qobject_cast<QLineEdit*>(widget)) editor->setAlignment(horizontal|(editor->alignment()&Qt::AlignVertical_Mask)|Qt::AlignAbsolute);
    else if (auto* label=qobject_cast<QLabel*>(widget)) label->setAlignment(horizontal|(label->alignment()&Qt::AlignVertical_Mask)|Qt::AlignAbsolute);
}

void AccessFormPreview::refreshTextAlignments() {
    for (const auto& control:layout_.controls)
        if (auto* widget=controls_.value(control.name.toCaseFolded()).data()) applyTextAlignment(widget,control);
}

void AccessFormPreview::buildControls(QWidget* canvas) {
    for (const auto& control : layout_.controls) {

        QWidget* widget = nullptr;
        if (control.locked && control.type!="TextBox") {
            auto* unavailable=new QLabel("Steuerelement-Sperre nicht unterstützt (3251)",canvas); unavailable->setWordWrap(true); widget=unavailable;
            if (control.type=="Subform") subformErrors_.insert(control.name.toCaseFolded(),3251);
        } else if (control.type=="OptionGroup") {
            auto* group=new QGroupBox(canvas); group->setFocusPolicy(Qt::StrongFocus);
            auto* buttons=new QButtonGroup(group); buttons->setObjectName("access-options");
            widget=group;
        } else if (control.type == "Label") {
            auto* label = new QLabel(control.caption, canvas);
            label->setTextFormat(Qt::PlainText);
            label->setWordWrap(true);
            widget = label;
        } else if (control.type=="CustomControl" && (control.className.compare("MSCAL.Calendar.7",Qt::CaseInsensitive)==0 || control.className.compare("MSCAL.Kalender-Steuerelement.7",Qt::CaseInsensitive)==0)) {
            auto* calendar=new QCalendarWidget(canvas);
            calendar->setDateRange(QDate(100,1,1),QDate(9999,12,31));
            calendar->setMinimumSize(0,0); widget=calendar;
            calendar->setProperty("la_calendar_bound",!control.controlSource.isEmpty());
            if (!control.controlSource.isEmpty()) { calendar->setEnabled(false); calendar->setToolTip("Kalender-Datenbindung wird noch nicht unterstützt."); }
            auto clickBinding=control.onClick;
            // MSCAL's source-interface Click event is stored in CodeBehind,
            // without a SaveAsText OnClick property in the original export.
            if (clickBinding.isEmpty() && linaccess::validateVbaProcedure(vba_.toUtf8().toStdString(),(controlEventPrefix(control.name)+"_Click").toStdString()).status!=linaccess::VbaExecutionStatus::ProcedureMissing)
                clickBinding="[Event Procedure]";
            if (!clickBinding.isEmpty()) connect(calendar,&QCalendarWidget::clicked,this,[this,name=control.name,binding=clickBinding,calendar](const QDate&) {
                calendar->setProperty("la_null",false); markDirty(name,true);
                (void)runControlBinding(name,"Click",binding);
            });
        } else if (control.type == "TextBox") {
            auto* edit = new QLineEdit(canvas);
            // QLineEdit defaults to 32767 UTF-16 units. A bound Memo may be
            // longer, and saving any field writes the bound control values.
            edit->setMaxLength(std::numeric_limits<int>::max());
            edit->setReadOnly(control.controlSource.startsWith('='));
            widget = edit;
        } else if(control.type=="ListBox") {
            widget=new AccessListBox(canvas,control.multiSelect);
        } else if (control.type == "ComboBox") {
            auto* combo = new AccessComboBox(canvas);
            combo->setEditable(true);
            combo->setMaxVisibleItems(control.listRows);
            combo->preparePopup=[this,combo,name=control.name] {
                try {
                    if (!combo->property("la_rows_loaded").toBool()) requeryControl(name.toStdString());
                    return true;
                } catch (const std::exception& error) {
                    const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error);
                    combo->setProperty("la_rows_error",typed ? typed->number : 5);
                    if (status_) status_->setText("Listendaten nicht verfügbar: "+QString::fromUtf8(error.what()));
                    return false;
                }
            };
            widget = combo;
        } else if (control.type=="OptionButton" && control.optionGroupIndex>=0) {
            widget=new AccessRadioButton(control.caption,canvas);
        } else if (control.type == "CheckBox" || control.type == "OptionButton") {
            widget = new AccessCheckBox(control.caption, canvas);
        } else if (control.type == "CommandButton" || control.type == "ToggleButton") {
            auto* button = new AccessPushButton(control.caption.isEmpty() ? control.name : control.caption, canvas);
            widget = button;
            if (control.optionGroupIndex<0 && !control.onClick.isEmpty()) connect(button,&QPushButton::clicked,this,[this,name=control.name,binding=control.onClick] {
                (void)runControlBinding(name,"Click",binding);
            });
        } else if (control.type == "Subform") {
            try {
                if (!control.linkMasterFields.isEmpty() || !control.linkChildFields.isEmpty())
                    throw linaccess::VbaRuntimeError(3251,"Linked subform binding is not implemented");
                auto name=control.sourceObject;
                if (name.startsWith("Form.",Qt::CaseInsensitive)) name=name.mid(5);
                if (name.isEmpty()) throw linaccess::VbaRuntimeError(2455,"Subform source unavailable");
                QSqlQuery source(db_);
                source.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='form' AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1");
                source.addBindValue(name);
                if (!source.exec() || !source.next()) throw linaccess::VbaRuntimeError(2455,"Subform source unavailable");
                const auto definition=source.value(0).toByteArray(); source.finish();
                if (linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(definition)).kind!=linaccess::AccessDocumentKind::Form)
                    throw linaccess::VbaRuntimeError(2455,"Subform source is not a form");
                source.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='vba' AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1");
                source.addBindValue("Form_"+name);
                const auto code=source.exec() && source.next() ? source.value(0).toByteArray() : QByteArray{}; source.finish();
                auto* child=new AccessFormPreview(name,definition,code,db_,canvas,true);
                if (child->cancelledByOpen()) { delete child; throw linaccess::VbaRuntimeError(2501,"Subform opening cancelled"); }
                subforms_.insert(control.name.toCaseFolded(),child); widget=child;
            } catch (const std::exception& error) {
                const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error);
                const int number=typed ? typed->number : 5;
                subformErrors_.insert(control.name.toCaseFolded(),number);
                auto* label=new QLabel("Unterformular nicht verfügbar ("+QString::number(number)+")",canvas);
                label->setWordWrap(true); widget=label;
            }
        } else if (control.type == "Rectangle" || control.type == "Line") {
            auto* shape = new QLabel(canvas);
            shape->setFrameShape(control.type == "Rectangle" ? QFrame::Box : QFrame::HLine);
            widget = shape;
        }
        if (!widget) continue;
        applyControlFont(widget,control);
        applyForeColor(widget,control);
        applyControlBackground(widget,control);
        applyTextAlignment(widget,control);
        widget->setObjectName("access-control-" + control.name);
        widget->setProperty("la_access_tag",control.tag);
        if (!dynamic_cast<AccessFormPreview*>(widget)) widget->installEventFilter(this);
        if (auto* combo=qobject_cast<QComboBox*>(widget); combo && combo->isEditable()) combo->lineEdit()->installEventFilter(this);
        if (auto* button=qobject_cast<QPushButton*>(widget)) { button->setAutoDefault(false); button->setDefault(control.type=="CommandButton" && control.defaultButton); }
        widget->setGeometry(pixels(control.left), pixels(control.top),
                            qMax(12, pixels(control.width)), qMax(14, pixels(control.height)));
        if (control.optionGroupIndex>=0) {
            const auto& parent=layout_.controls[control.optionGroupIndex];
            if (auto* group=qobject_cast<QGroupBox*>(controls_.value(parent.name.toCaseFolded()).data())) {
                const auto geometry=widget->geometry(); widget->setParent(group);
                widget->setGeometry(geometry.translated(-group->pos()));
                if (auto* button=qobject_cast<QAbstractButton*>(widget)) {
                    button->setCheckable(true); button->setProperty("la_option_value",control.optionValue);
                    button->setProperty("la_option_group_name",parent.name);
                    group->findChild<QButtonGroup*>("access-options")->addButton(button);
                    connect(button,&QAbstractButton::clicked,this,[this,group,button,parent] {
                        if (closed_ || openCancelled_ || loading_ || layout_.kind!=linaccess::AccessDocumentKind::Form) return;
                        ++nativeExecutionDepth(); const auto executing=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
                        if (!group->property("la_access_input_readonly").toBool()) {
                            clearDefaultError(group); group->setProperty("la_option_value",button->property("la_option_value"));
                            group->setProperty("la_null",false); group->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Number));
                            markDirty(parent.name,true); if (closed_ || openCancelled_ || group->property("la_insert_rejected").toBool()) return;
                            refreshCalculatedControls(); pendingControls_.insert(parent.name);
                            if (!commitControl(parent.name)) return;
                        }
                        // Access dispatches Click to the option group, not its members.
                        if (!closed_ && !openCancelled_) (void)runControlBinding(parent.name,"Click",parent.onClick);
                    });
                }
            }
        }
        widget->setVisible(control.visible);
        if (!control.name.isEmpty()) controls_.insert(control.name.toCaseFolded(), widget);
        const auto name = control.name;
        widget->setProperty("la_null",!qobject_cast<QCalendarWidget*>(widget));
        if (!control.onDblClick.isEmpty()) {
            widget->setProperty("la_access_dblclick",control.onDblClick);
            widget->setProperty("la_access_control_name",control.name); widget->installEventFilter(this);
            if(auto* list=dynamic_cast<AccessListBox*>(widget)) {
                list->viewport()->setProperty("la_access_dblclick",control.onDblClick);
                list->viewport()->setProperty("la_access_control_name",control.name);
                list->viewport()->installEventFilter(this);
            }
        }
        widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Null));
        if (auto* line = qobject_cast<QLineEdit*>(widget)) {
            connect(line,&QLineEdit::textChanged,this,[this,name,widget,binding=control.onChange] {
                if (loading_) return;
                clearDefaultError(widget); widget->setProperty("la_null",false); widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text));
                markDirty(name,true); refreshCalculatedControls();
                if (!widget->property("la_insert_rejected").toBool() && !binding.isEmpty()) (void)runControlBinding(name,"Change",binding);
            });
            connect(line,&QLineEdit::textEdited,this,[this,name,widget] {
                if (!loading_) { bufferedTextControls_.insert(name); pendingControls_.insert(name); widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text)); }
            });
            connect(line,&QLineEdit::editingFinished,this,[this,name,line] { if (!loading_ && !line->property("la_native_zoom_open").toBool()) (void)commitControl(name); });
        } else if(auto* list=dynamic_cast<AccessListBox*>(widget)) {
            list->selectionEdited=[this,name,list] {
                if(closed_ || openCancelled_ || loading_ || list->property("la_access_input_readonly").toBool()) return;
                ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
                clearDefaultError(list); list->setProperty("la_null",list->multiSelectMode()!=0);
                if(!list->multiSelectMode()) markDirty(name,true);
                if(closed_ || openCancelled_ || list->property("la_insert_rejected").toBool()) return;
                refreshCalculatedControls(); pendingControls_.insert(name); (void)commitControl(name);
            };
            connect(list,&QTreeView::clicked,this,[this,name,list,binding=control.onClick] {
                if(closed_ || openCancelled_ || loading_ || layout_.kind!=linaccess::AccessDocumentKind::Form) return;
                ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
                if(!commitControl(name) || closed_ || openCancelled_) return;
                (void)runControlBinding(name,"Click",binding);
            });
        } else if (auto* combo = qobject_cast<QComboBox*>(widget)) {
            if (auto* editor=combo->lineEdit()) {
                connect(editor,&QLineEdit::textEdited,this,[this,name] {
                    if (!loading_) { bufferedTextControls_.insert(name); pendingControls_.insert(name); }
                });
                connect(editor,&QLineEdit::editingFinished,this,[this,name] { if (!loading_) (void)commitControl(name); });
            }
            connect(combo,&QComboBox::currentTextChanged,this,[this,name,widget,binding=control.onChange] {
                if (loading_) return;
                clearDefaultError(widget); widget->setProperty("la_null",false); widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text));
                markDirty(name,true); refreshCalculatedControls();
                if (!widget->property("la_insert_rejected").toBool() && !binding.isEmpty()) (void)runControlBinding(name,"Change",binding);
            });
            connect(combo,QOverload<int>::of(&QComboBox::activated),this,[this,name](int) {
                if (loading_) return;
                if (auto* widget=controls_.value(name.toCaseFolded()).data(); widget && widget->property("la_access_input_readonly").toBool()) return;
                if (auto* widget=controls_.value(name.toCaseFolded()).data(); widget && widget->property("la_insert_rejected").toBool()) { widget->setProperty("la_insert_rejected",false); return; }
                if (auto* widget=controls_.value(name.toCaseFolded()).data()) widget->setProperty("la_null",false);
                markDirty(name,true);
                if (auto* widget=controls_.value(name.toCaseFolded()).data(); widget && widget->property("la_insert_rejected").toBool()) { widget->setProperty("la_insert_rejected",false); return; }
                refreshCalculatedControls(); pendingControls_.insert(name); (void)commitControl(name);
            });
        } else if (auto* box = qobject_cast<QCheckBox*>(widget);box && control.optionGroupIndex<0) {
            connect(box,&QCheckBox::toggled,this,[this,name,widget] {
                if (closed_ || openCancelled_ || loading_) return;
                ++nativeExecutionDepth(); const auto executing=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
                clearDefaultError(widget); widget->setProperty("la_null",false); markDirty(name,true);
                if (!closed_ && !openCancelled_) refreshCalculatedControls();
            });
            connect(box,&QCheckBox::clicked,this,[this,name,widget,binding=control.onClick] {
                if (closed_ || openCancelled_ || loading_ || layout_.kind!=linaccess::AccessDocumentKind::Form) return;
                ++nativeExecutionDepth(); const auto executing=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
                if (!widget->property("la_access_input_readonly").toBool()) {
                    if (widget->property("la_insert_rejected").toBool()) { widget->setProperty("la_insert_rejected",false); return; }
                    pendingControls_.insert(name); if (!commitControl(name)) return;
                }
                if (!closed_ && !openCancelled_) (void)runControlBinding(name,"Click",binding);
            });
        }
        ++renderedControls_;
    }
    QPointer<QPushButton> cancelButton;
    for (const auto& control:layout_.controls) {
        if(control.type=="CommandButton") {
            if(auto* button=qobject_cast<QPushButton*>(controls_.value(control.name.toCaseFolded()).data())) {
                button->setProperty("la_access_cancel",control.cancelButton);
                if(control.cancelButton) {
                    if(cancelButton) cancelButton->setProperty("la_access_cancel",false);
                    cancelButton=button;
                }
            }
        }
        if (control.type!="Label" || control.attachedControlIndex<0 || control.attachedControlIndex>=layout_.controls.size()) continue;
        const auto& target=layout_.controls[control.attachedControlIndex];
        if (target.type!="TextBox" && target.type!="ComboBox" && target.type!="ListBox") continue;
        auto* label=qobject_cast<QLabel*>(controls_.value(control.name.toCaseFolded()).data());
        auto* buddy=controls_.value(target.name.toCaseFolded()).data();
        if (label && buddy) label->setBuddy(buddy);
    }
    initializeControlDefaults(true);
    applyControlStates();
}

bool AccessFormPreview::supportsControlDefault(const linaccess::AccessVisualControl& control) {
    return control.optionGroupIndex<0 && !control.multiSelect && !(control.locked && control.type!="TextBox" && control.type!="ListBox") &&
        (control.type=="TextBox" || control.type=="ComboBox" || control.type=="ListBox" || control.type=="CheckBox" || control.type=="OptionButton" || control.type=="OptionGroup");
}
void AccessFormPreview::initializeControlDefaults(bool unbound) {
    if(layout_.kind!=linaccess::AccessDocumentKind::Form || (!unbound && (currentRecordIndex_!=0 || !directTable_ || !dataSourceAvailable_ || !layout_.allowAdditions))) return;
    QScopedValueRollback<bool> loading(loading_,true);
    QSet<int> initialized; QSet<QString> generated;
    const auto failed=[&](const linaccess::AccessVisualControl& control,int number) {
        auto* widget=controls_.value(control.name.toCaseFolded()).data(); if(!widget) return;
        widget->setProperty("la_default_error",number);
        auto* editor=qobject_cast<QLineEdit*>(widget); if(auto* combo=qobject_cast<QComboBox*>(widget)) editor=combo->lineEdit();
        if(editor) {
            if(!editor->property("la_default_placeholder").isValid()) editor->setProperty("la_default_placeholder",editor->placeholderText());
            editor->setPlaceholderText("#Fehler ("+QString::number(number)+")");
        } else {
            if(!widget->property("la_default_tooltip").isValid()) widget->setProperty("la_default_tooltip",widget->toolTip());
            widget->setToolTip("Standardwert nicht verfügbar ("+QString::number(number)+")");
        }
        if(status_) status_->setText("Standardwert nicht verfügbar ("+QString::number(number)+")");
    };
    bool generatedLoaded=false;
    for(const auto& control:layout_.controls) {
        if(control.defaultValue.isEmpty() || !supportsControlDefault(control) || control.controlSource.isEmpty()!=unbound || control.controlSource.startsWith('=')) continue;
        const int column=unbound ? -1 : columnIndex(control.controlSource,dataColumns_);
        if(!unbound && column<0) { failed(control,2465); continue; }
        if(!unbound && initialized.contains(column)) continue;
        try {
            if(!unbound) {
                if(!generatedLoaded) { generated=linaccess::generatedAccessColumns(db_,layout_.recordSource.trimmed()); generatedLoaded=true; }
                if(generated.contains(dataColumns_[column].toCaseFolded())) continue;
                initialized.insert(column);
                for(const auto& other:layout_.controls) if(columnIndex(other.controlSource,dataColumns_)==column && !other.defaultValue.isEmpty() && other.defaultValue!=control.defaultValue)
                    throw linaccess::VbaRuntimeError(3251,"Conflicting defaults for one field unsupported");
            }
            const auto value=defaultLiteral(control.defaultValue);
            if(unbound) setFieldValue(control.name.toUtf8().toStdString(),value,true);
            else for(const auto& other:layout_.controls) if(columnIndex(other.controlSource,dataColumns_)==column && supportsControlDefault(other))
                setFieldValue(other.name.toUtf8().toStdString(),value,true);
        } catch(const std::exception& error) {
            const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error); const int number=typed ? typed->number : 3251;
            if(unbound) failed(control,number);
            else for(const auto& other:layout_.controls) if(columnIndex(other.controlSource,dataColumns_)==column) failed(other,number);
        }
    }
}

void AccessFormPreview::applyControlStates() {
    QVector<int> tabOrder;
    for (int index=0;index<layout_.controls.size();++index) {
        const auto& control=layout_.controls[index]; auto* widget=controls_.value(control.name.toCaseFolded()).data();
        if (!widget) continue;
        widget->setVisible(control.visible); widget->setEnabled(control.enabled);
        const bool bound=!control.controlSource.isEmpty();
        const bool parentReadOnly=control.optionGroupIndex>=0 && controls_.value(layout_.controls[control.optionGroupIndex].name.toCaseFolded()) && controls_.value(layout_.controls[control.optionGroupIndex].name.toCaseFolded())->property("la_access_input_readonly").toBool();
        const bool readOnly=parentReadOnly || layout_.kind==linaccess::AccessDocumentKind::Report || control.locked || control.controlSource.startsWith('=') || (bound && (!directTable_ || (currentRecordIndex_==0 && !layout_.allowAdditions) || (currentRecordIndex_>0 && !layout_.allowEdits && !dirty_)));
        widget->setProperty("la_access_input_readonly",readOnly);
        if (auto* line=qobject_cast<QLineEdit*>(widget)) {
            const bool password=control.inputMask.compare("Password",Qt::CaseInsensitive)==0;
            const bool unsupportedMask=!control.inputMask.isEmpty() && !password;
            line->setReadOnly(readOnly || unsupportedMask);
            // Password is an echo policy, not a transformation of the stored value.
            const bool passwordDisplay=password || control.inputMask.section(';',0,0).compare("Password",Qt::CaseInsensitive)==0;
            line->setEchoMode(passwordDisplay ? QLineEdit::Password : QLineEdit::Normal);
            if (passwordDisplay) {
                if (!line->property("la_password_original_style").isValid()) line->setProperty("la_password_original_style",line->styleSheet());
                const auto passwordStyle=line->property("la_password_original_style").toString()+"\nQLineEdit { lineedit-password-character: 42; lineedit-password-mask-delay: 0; }";
                if (line->styleSheet()!=passwordStyle) line->setStyleSheet(passwordStyle);
            } else if (line->property("la_password_original_style").isValid()) {
                line->setStyleSheet(line->property("la_password_original_style").toString()); line->setProperty("la_password_original_style",QVariant{});
            }
            if (unsupportedMask) {
                if (!line->property("la_mask_original_tooltip").isValid()) line->setProperty("la_mask_original_tooltip",line->toolTip());
                line->setToolTip("Access-Eingabemaske nicht unterstützt (3251); Benutzereingabe gesperrt.");
            } else if (line->property("la_mask_original_tooltip").isValid()) {
                line->setToolTip(line->property("la_mask_original_tooltip").toString()); line->setProperty("la_mask_original_tooltip",QVariant{});
            }
        }
        else if (auto* combo=qobject_cast<QComboBox*>(widget);combo && combo->lineEdit()) combo->lineEdit()->setReadOnly(readOnly);
        else if (auto* group=qobject_cast<QGroupBox*>(widget))
            for (auto* button:group->findChildren<QAbstractButton*>()) button->setProperty("la_access_input_readonly",readOnly);
        else if (auto* calendar=qobject_cast<QCalendarWidget*>(widget)) calendar->setEnabled(control.enabled && !calendar->property("la_calendar_bound").toBool());
        applyForeColor(widget,control);
        applyControlBackground(widget,control);
        const bool focusable=!qobject_cast<QLabel*>(widget) && control.type!="Image" && control.type!="Rectangle" && control.type!="Line";
        if (focusable) {
            widget->setFocusPolicy(control.tabStop ? Qt::StrongFocus : Qt::ClickFocus);
            if (auto* combo=qobject_cast<QComboBox*>(widget);combo && combo->lineEdit()) combo->lineEdit()->setFocusPolicy(control.tabStop ? Qt::StrongFocus : Qt::ClickFocus);
            tabOrder.append(index);
        }
    }
    const auto section=[](const QString& name) { return name=="FormHeader" ? 0 : name=="FormFooter" ? 2 : 1; };
    std::stable_sort(tabOrder.begin(),tabOrder.end(),[&](int a,int b) {
        const auto& left=layout_.controls[a]; const auto& right=layout_.controls[b];
        const auto first=section(left.section),second=section(right.section);
        return first!=second ? first<second : left.tabIndex<right.tabIndex;
    });
    for (int index=1;index<tabOrder.size();++index)
        QWidget::setTabOrder(controls_.value(layout_.controls[tabOrder[index-1]].name.toCaseFolded()).data(),controls_.value(layout_.controls[tabOrder[index]].name.toCaseFolded()).data());
}

bool AccessFormPreview::flatTabScope() const {
    for (const auto& control:layout_.controls)
        if (control.optionGroupIndex>=0 || control.type=="OptionGroup" || control.type=="Subform" || control.type=="Tab" || control.type=="CustomControl") return false;
    return true;
}

bool AccessFormPreview::advanceTabFocus(bool next) {
    if (cyclingFocus_) return true;
    if (!isWindow() || !opened_ || closed_ || openCancelled_ || layout_.kind!=linaccess::AccessDocumentKind::Form || !flatTabScope())
        return false;
    auto* origin=QApplication::focusWidget(); if (!origin) return false;
    struct Entry { QPointer<QWidget> widget; QString name,section; int index; };
    const auto collect=[&] {
        QVector<Entry> entries;
        for (const auto& control:layout_.controls) {
            auto* widget=controls_.value(control.name.toCaseFolded()).data();
            if (!widget || qobject_cast<QLabel*>(widget) || widget->focusPolicy()==Qt::NoFocus) continue;
            entries.append({widget,control.name,control.section,control.tabIndex});
        }
        const auto section=[](const QString& name) { return name=="FormHeader" ? 0 : name=="FormFooter" ? 2 : 1; };
        std::stable_sort(entries.begin(),entries.end(),[&](const Entry& a,const Entry& b) {
            return section(a.section)!=section(b.section) ? section(a.section)<section(b.section) : a.index<b.index;
        });
        return entries;
    };
    const auto locate=[&](const QVector<Entry>& entries) {
        for (int index=0;index<entries.size();++index)
            if (entries[index].widget && (entries[index].widget==origin || entries[index].widget->isAncestorOf(origin))) return index;
        return -1;
    };
    auto entries=collect(); auto current=locate(entries);
    if (current<0) return false; // Toolbar/navigation focus stays native.
    QScopedValueRollback<bool> cycling(cyclingFocus_,true);
    ++nativeExecutionDepth();
    const auto unwind=qScopeGuard([&] { --nativeExecutionDepth(); disposeClosedForms(); });
    const QPointer<AccessFormPreview> guard(this);
    try {
        if (!commitControl(entries[current].name)) { status_->setText("Tab-Wechsel abgebrochen (2501)"); return true; }
        if (!guard || closed_ || QApplication::focusWidget()!=origin) return true;
        entries=collect(); current=locate(entries); if (current<0) return true;
        const auto eligible=[](const Entry& entry) { return entry.widget && entry.widget->isVisible() && entry.widget->isEnabled() && (entry.widget->focusPolicy()&Qt::TabFocus); };
        const int step=next ? 1 : -1;
        int target=current+step;
        for (;target>=0 && target<entries.size();target+=step) if (eligible(entries[target])) break;
        if (target<0 || target>=entries.size()) {
            if (layout_.cycle==2) throw linaccess::VbaRuntimeError(3251,"Page cycle unsupported");
            if (layout_.cycle==0 && !layout_.recordSource.trimmed().isEmpty() && !(next && currentRecordIndex_==0 && !dirty_)) {
                if (!dataSourceAvailable_) throw linaccess::VbaRuntimeError(dataSourceErrorNumber_ ? dataSourceErrorNumber_ : 3078,"Record source unavailable");
                if (dataTruncated_) throw linaccess::VbaRuntimeError(3035,"Incomplete record navigation");
                if (next && !directTable_ && currentRecordIndex_==recordRows_.size()) throw linaccess::VbaRuntimeError(2046,"Read-only source has no new record");
                PreviewVbaHost host(*this); host.gotoRecord(next ? 1 : 0);
                if (!guard || closed_ || QApplication::focusWidget()!=origin) return true;
                entries=collect();
            }
            target=next ? 0 : entries.size()-1;
            for (;target>=0 && target<entries.size();target+=step) if (eligible(entries[target])) break;
        }
        if (target>=0 && target<entries.size() && eligible(entries[target])) entries[target].widget->setFocus(next ? Qt::TabFocusReason : Qt::BacktabFocusReason);
    } catch (const linaccess::VbaRuntimeError& error) {
        if (guard && !closed_) status_->setText("Tab-Wechsel abgebrochen ("+QString::number(error.number)+")");
    } catch (const std::exception&) {
        if (guard && !closed_) status_->setText("Tab-Wechsel abgebrochen (5)");
    }
    return true;
}

void AccessFormPreview::loadRecords() {
    recordsetSnapshot_.reset();
    const bool wasNew = initialized_ && currentRecordIndex_ == 0;
    const auto oldId = currentRecordIndex_ > 0 && currentRecordIndex_ <= rowIds_.size()
        ? rowIds_[currentRecordIndex_-1] : QVariant{};
    const int oldIndex = currentRecordIndex_;
    recordSelector_->blockSignals(true);
    recordSelector_->clear(); recordRows_.clear(); rowIds_.clear(); dataColumns_.clear(); dataTypes_.clear();
    dataSourceAvailable_ = false; directTable_ = false; dataTruncated_ = false;
    dataSourceErrorNumber_=0;
    recordSelector_->addItem(layout_.kind==linaccess::AccessDocumentKind::Report ? "(Kein Datensatz)" : "(Neuer Datensatz)");
    if (auto* model=qobject_cast<QStandardItemModel*>(recordSelector_->model());model && model->item(0)) model->item(0)->setEnabled(layout_.kind==linaccess::AccessDocumentKind::Form && layout_.allowAdditions);
    const auto source = layout_.recordSource.trimmed();
    if (!source.isEmpty() && db_.isOpen()) {
        try {
            directTable_ = layout_.kind==linaccess::AccessDocumentKind::Form && db_.tables().contains(source,Qt::CaseInsensitive);
            if (!directTable_) {
                linaccess::AccessQueryPlan plan;
                if (db_.tables().contains(source,Qt::CaseInsensitive) || db_.tables(QSql::Views).contains(source,Qt::CaseInsensitive))
                    plan.sql = "SELECT * FROM " + sqlIdentifier(source);
                else {
                    const auto sql = source.startsWith("SELECT",Qt::CaseInsensitive)
                        || source.startsWith("PARAMETERS",Qt::CaseInsensitive)
                        ? source : linaccess::accessQuerySource(db_,source);
                    plan = linaccess::compileAccessQuery(sql);
                }
                if (filterOn_) plan=linaccess::filterAccessQueryPlan(db_,plan,filter_);
                if(orderByOn_) plan=linaccess::sortAccessQueryPlan(db_,plan,orderBy_);
                auto result = linaccess::executeAccessQueryPlan(db_,plan,linaccess::bindAccessQueryParameters(plan,{}));
                dataColumns_ = result.columns; dataTypes_=result.types;
                dataTruncated_ = result.rows.size() > 10000;
                recordRows_ = result.rows.mid(0,10000);
                for (int i = 0; i < recordRows_.size(); ++i)
                    recordSelector_->addItem("Datensatz " + QString::number(i+1));
                dataSourceAvailable_ = true;
            } else {
                dataTypes_ = linaccess::sqliteColumnTypes(db_,source);
                linaccess::SqliteExecutionBudget budget(db_);
                QSqlQuery query(db_);
                auto name=source; name.replace(']',"]]");
                const auto plan=linaccess::compileAccessQuery("SELECT rowid AS [__la_rowid],* FROM ["+name+"]"+(filterOn_ && !filter_.trimmed().isEmpty() ? " WHERE ("+filter_+")" : QString{}));
                QStringList sortColumns; const auto record=db_.record(source);
                for(int i=0;i<record.count();++i) sortColumns.append(record.fieldName(i));
                const auto order=orderByOn_ ? linaccess::accessOrderByClause(orderBy_,sortColumns,source) : QString{};
                auto statement=plan.sql;
                QVariantList entryBindings;
                if(layout_.dataEntry && layout_.allowAdditions) {
                    const bool hasWhere=filterOn_ && !filter_.trimmed().isEmpty();
                    QStringList entryPlaceholders;
                    for(const auto id:dataEntryRows_) { entryPlaceholders.append("?"); entryBindings.append(id); }
                    statement+=(hasWhere ? " AND " : " WHERE ")+(entryPlaceholders.isEmpty() ? QString("0") : "rowid IN ("+entryPlaceholders.join(',')+")");
                }
                const bool prepared=query.prepare(statement+order);
                if(prepared) for(const auto& id:entryBindings) query.addBindValue(id);
                if (!prepared || !query.exec()) {
                    const auto message=query.lastError().text().toCaseFolded();
                    const auto typed=linaccess::jetSqlFunctionErrorNumber(message.toStdString());
                    const auto code=typed ? *typed : (message.contains("syntax error") || message.contains("incomplete input")) ? 3075 : message.contains("no such table") ? 3078 : message.contains("no such function") ? 3085 : message.contains("interrupted") ? 3035 : 3146;
                    throw linaccess::VbaRuntimeError(code,"Record source query failed");
                }
                for (int i = 1; i < query.record().count(); ++i) dataColumns_ << query.record().fieldName(i);
                while (query.next()) {
                    if (recordRows_.size() >= 10000) { dataTruncated_ = true; break; }
                    QVariantList row;
                    rowIds_.push_back(query.value(0));
                    for (int i = 1; i < query.record().count(); ++i) row << query.value(i);
                    recordRows_.push_back(row);
                    recordSelector_->addItem("Datensatz " + QString::number(recordRows_.size()));
                }
                if (query.lastError().isValid()) throw std::runtime_error("Record source read interrupted");
                dataSourceAvailable_ = true;
            }
        } catch (const std::exception& error) {
            const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error);
            dataSourceErrorNumber_=typed ? (typed->number==3265 ? 3078 : typed->number) : 3146;
            recordRows_.clear(); rowIds_.clear(); dataColumns_.clear();
            while (recordSelector_->count() > 1) recordSelector_->removeItem(recordSelector_->count()-1);
        }
    }
    int selected = wasNew ? 0 : !oldId.isNull() ? rowIds_.indexOf(oldId)+1
        : oldIndex > 0 ? qMin(oldIndex,recordSelector_->count()-1) : recordSelector_->count() > 1 ? 1 : 0;
    if (selected <= 0 && !wasNew && recordSelector_->count() > 1) selected = 1;
    recordSelector_->setCurrentIndex(selected);
    currentRecordIndex_ = selected; initialized_ = true;
    recordSelector_->blockSignals(false);
    loadRecord();
    if (auto* save = findChild<QPushButton*>("access-save-record")) save->setEnabled(dataSourceAvailable_ && directTable_);
    if (status_) status_->setText(dataSourceAvailable_
        ? "Access-Layout: " + QString::number(renderedControls_) + " Elemente, "
            + QString::number(recordRows_.size()) + " Datensätze." + (!directTable_ ? " Datenquelle schreibgeschützt." : "") + (dataTruncated_ ? " Anzeige auf die ersten 10000 begrenzt." : "")
        : "Die Datenquelle ist nicht verfügbar oder ihre Ausführung wurde abgebrochen: " + source);
}

void AccessFormPreview::setDataEntryState(bool enabled) {
    if(layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"DataEntry requires a form");
    if(recordCommandRunning_ || changingDataEntry_ || changingFilter_ || changingOrder_)
        throw linaccess::VbaRuntimeError(3251,"Reentrant DataEntry transition unsupported");
    const bool bound=!layout_.recordSource.trimmed().isEmpty();
    if(enabled && layout_.allowAdditions && bound && !directTable_) throw linaccess::VbaRuntimeError(3251,"Native DataEntry requires a direct table source");
    if(layout_.dataEntry==enabled && (enabled || (!filterOn_ && !orderByOn_))) return;
    QScopedValueRollback<bool> changing(changingDataEntry_,true);
    if(!saveRecord()) throw linaccess::VbaRuntimeError(2501,"DataEntry save cancelled");
    if(closed_ || openCancelled_) throw linaccess::VbaRuntimeError(2450,"Form closed during DataEntry validation");
    layout_.dataEntry=enabled; dataEntryRows_.clear();
    if(!enabled) {
        filterOn_=false; orderByOn_=false;
        if(filterToggle_) { const QSignalBlocker blocked(filterToggle_); filterToggle_->setChecked(false); }
    }
    if(!bound) return;
    if(!enabled && currentRecordIndex_==0) currentRecordIndex_=-1;
    loadRecords(); applyControlStates();
    if(opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
}

void AccessFormPreview::setAdditionPolicy(bool enabled) {
    if(layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"AllowAdditions requires a form");
    if(recordCommandRunning_ || changingDataEntry_) throw linaccess::VbaRuntimeError(3251,"Changing addition policy during record command unsupported");
    if(layout_.allowAdditions==enabled) return;
    if(layout_.dataEntry && !layout_.recordSource.trimmed().isEmpty()) {
        if(enabled && !directTable_) throw linaccess::VbaRuntimeError(3251,"Native DataEntry requires a direct table source");
        QScopedValueRollback<bool> changing(changingDataEntry_,true);
        if(!saveRecord()) throw linaccess::VbaRuntimeError(2501,"Addition policy save cancelled");
        if(closed_ || openCancelled_) throw linaccess::VbaRuntimeError(2450,"Form closed during addition policy validation");
        layout_.allowAdditions=enabled; dataEntryRows_.clear();
        if(!enabled && currentRecordIndex_==0) currentRecordIndex_=-1;
        loadRecords(); applyControlStates();
        if(opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
        return;
    }
    if(!enabled && currentRecordIndex_==0 && dirty_ && !saveRecord()) throw linaccess::VbaRuntimeError(2501,"Addition policy save cancelled");
    layout_.allowAdditions=enabled;
    if(auto* model=qobject_cast<QStandardItemModel*>(recordSelector_->model());model && model->item(0)) model->item(0)->setEnabled(enabled);
    if(currentRecordIndex_==0) loadRecord();
}

void AccessFormPreview::setOrderState(const QString& orderBy,bool enabled) {
    if(layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"OrderBy requires a form");
    if(recordCommandRunning_ || changingOrder_ || changingFilter_) throw linaccess::VbaRuntimeError(3251,"Reentrant record sorting unsupported");
    const bool reload=orderByOn_!=enabled || (enabled && orderBy_!=orderBy);
    if(!reload || (!enabled && layout_.recordSource.trimmed().isEmpty())) { orderBy_=orderBy; orderByOn_=enabled; return; }
    if(layout_.recordSource.trimmed().isEmpty() || !dataSourceAvailable_) throw linaccess::VbaRuntimeError(3251,"Sort source unavailable");
    if(enabled) (void)linaccess::accessOrderByClause(orderBy,dataColumns_,directTable_ ? layout_.recordSource.trimmed() : QString{});
    QScopedValueRollback<bool> changing(changingOrder_,true);
    if(!saveRecord()) throw linaccess::VbaRuntimeError(2501,"Sort change cancelled");
    const auto oldOrder=orderBy_; const bool oldEnabled=orderByOn_; const auto oldIndex=currentRecordIndex_;
    const auto oldId=oldIndex>0 && oldIndex<=rowIds_.size() ? rowIds_[oldIndex-1] : QVariant{};
    orderBy_=orderBy; orderByOn_=enabled; loadRecords();
    if(!dataSourceAvailable_) {
        const auto number=dataSourceErrorNumber_ ? dataSourceErrorNumber_ : 3146;
        orderBy_=oldOrder; orderByOn_=oldEnabled; loadRecords();
        if(dataSourceAvailable_) {
            const auto index=oldIndex==0 ? 0 : !oldId.isNull() ? rowIds_.indexOf(oldId)+1 : qMin(oldIndex,recordSelector_->count()-1);
            const QSignalBlocker blocked(recordSelector_); recordSelector_->setCurrentIndex(qMax(0,index)); currentRecordIndex_=qMax(0,index); loadRecord();
        }
        throw linaccess::VbaRuntimeError(number,"Sort execution failed");
    }
    if(opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
}

void AccessFormPreview::setFilterState(const QString& filter,bool enabled) {
    if (recordCommandRunning_ || changingOrder_) throw linaccess::VbaRuntimeError(3251,"Filter change during record/sort command unsupported");
    const bool reload=filterOn_!=enabled || (enabled && filter_!=filter);
    if (reload && changingFilter_) throw linaccess::VbaRuntimeError(3251,"Reentrant filter change unsupported");
    if (reload && !enabled && layout_.recordSource.trimmed().isEmpty()) {
        filter_=filter; filterOn_=false;
    } else if (reload && !layout_.recordSource.trimmed().isEmpty() && !opened_ && layout_.kind==linaccess::AccessDocumentKind::Report) {
        filter_=filter; filterOn_=enabled;
    } else if (reload) {
        if (enabled && layout_.recordSource.trimmed().isEmpty()) throw linaccess::VbaRuntimeError(3078,"Filter record source unavailable");
        QScopedValueRollback<bool> changing(changingFilter_,true);
        if (!saveRecord()) throw linaccess::VbaRuntimeError(2501,"Filter change cancelled");
        const auto oldFilter=filter_; const bool oldEnabled=filterOn_; const auto oldIndex=currentRecordIndex_;
        const auto oldId=oldIndex>0 && oldIndex<=rowIds_.size() ? rowIds_[oldIndex-1] : QVariant{};
        filter_=filter; filterOn_=enabled; loadRecords();
        if (!dataSourceAvailable_) {
            const auto number=dataSourceErrorNumber_ ? dataSourceErrorNumber_ : 3146;
            filter_=oldFilter; filterOn_=oldEnabled; loadRecords();
            if (dataSourceAvailable_) {
                const auto index=oldIndex==0 ? 0 : !oldId.isNull() ? rowIds_.indexOf(oldId)+1 : qMin(oldIndex,recordSelector_->count()-1);
                const QSignalBlocker blocked(recordSelector_); recordSelector_->setCurrentIndex(qMax(0,index)); currentRecordIndex_=qMax(0,index); loadRecord();
            }
            throw linaccess::VbaRuntimeError(number,"Filter execution failed");
        }
        if (opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
    } else filter_=filter;
    const QSignalBlocker textBlocked(filterEditor_),toggleBlocked(filterToggle_);
    filterEditor_->setText(filter_); filterToggle_->setChecked(filterOn_);
}

void AccessFormPreview::loadRecord() {
    clearRecordSelection(); bufferedTextControls_.clear();
    QScopedValueRollback<bool> loading(loading_,true);
    for (const auto& control : layout_.controls) {
        if (control.controlSource.isEmpty() || (control.type=="ListBox" && control.multiSelect)) continue;
        auto* widget = controls_.value(control.name.toCaseFolded()).data();
        if (!widget) continue;
        clearDefaultError(widget);
        const int index = columnIndex(control.controlSource,dataColumns_);
        const int row = currentRecordIndex_-1;
        QVariant value;
        if (row >= 0 && row < recordRows_.size() && index >= 0) value = recordRows_[row][index];
        auto type = index >= 0 ? dataTypes_.value(dataColumns_[index].toCaseFolded()) : QString{};
        if (type.isEmpty() && value.isValid()) {
            const auto id = value.metaType().id();
            if (id == QMetaType::LongLong || id == QMetaType::Int) type = "INTEGER";
            else if (id == QMetaType::Double) type = "REAL";
            else if (id == QMetaType::Bool) type = "BOOLEAN";
            else type = "TEXT";
        }
        widget->setProperty("la_sql_type",type);
        setText(widget,qobject_cast<QLineEdit*>(widget) ? linaccess::qtFieldText(value,type) : value.toString(),value.isNull()); widget->setProperty("la_null",value.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text));
    }
    dirty_ = false; insertionStarted_=false; pendingControls_.clear(); editOriginals_.clear();
    initializeControlDefaults(false);
    applyControlStates(); refreshCalculatedControls();
    rememberControlEdit(QApplication::focusWidget());
}

void AccessFormPreview::markDirty(const QString& name,bool userEdit) {
    if (loading_ || !directTable_ || !dataSourceAvailable_) return;
    auto* editedWidget=controls_.value(name.toCaseFolded()).data();
    if(auto* list=dynamic_cast<AccessListBox*>(editedWidget); list && list->multiSelectMode()) return;
    if (userEdit && editedWidget) editedWidget->setProperty("la_insert_rejected",false);
    const auto reject=[&] {
        loadRecord();
        if (userEdit && editedWidget) editedWidget->setProperty("la_insert_rejected",true);
    };
    if (!name.isEmpty()) {
        int boundColumn = -1;
        auto* sourceWidget = controls_.value(name.toCaseFolded()).data();
        for (const auto& control : layout_.controls)
            if (control.name.compare(name,Qt::CaseInsensitive) == 0)
                boundColumn = columnIndex(control.controlSource,dataColumns_);
        if (boundColumn < 0 || !sourceWidget) return;
        if (!bufferedTextControls_.contains(name)) {
            QScopedValueRollback<bool> loading(loading_,true);
            for (const auto& control : layout_.controls) {
                if(control.type=="ListBox" && control.multiSelect) continue;
                if (columnIndex(control.controlSource,dataColumns_) != boundColumn) continue;
                auto* target = controls_.value(control.name.toCaseFolded()).data();
                if (!target || target == sourceWidget) continue;
                setText(target,textOf(sourceWidget),sourceWidget->property("la_null").toBool()); target->setProperty("la_null",sourceWidget->property("la_null"));
            }
        }
    }
    if (insertionEventRunning_) return;
    if (userEdit && currentRecordIndex_==0 && !insertionStarted_) {
        insertionStarted_=true;
        linaccess::VbaExecutionResult before;
        {
            QScopedValueRollback<bool> inserting(insertionEventRunning_,true);
            before=formEvent("BeforeInsert","Form_BeforeInsert");
        }
        if (before.cancelled || failed(before)) { reject(); return; }
        if (!insertionStarted_) { reject(); return; } // The handler undid the new row.
    }
    if (!dirty_) {
        dirty_ = true;
        const auto event = formEvent("OnDirty","Form_Dirty");
        if (event.cancelled || failed(event)) reject();
        if (!closed_ && !openCancelled_) applyControlStates();
    }
}

bool AccessFormPreview::commitControl(const QString& name) {
    if (!pendingControls_.contains(name)) return true;
    if (!updatingControl_.isEmpty() && updatingControl_.compare(name,Qt::CaseInsensitive)==0) return false;
    for (const auto& control : layout_.controls) {
        if (control.name.compare(name,Qt::CaseInsensitive) != 0) continue;
        QScopedValueRollback<QString> updating(updatingControl_,control.name);
        if (eventProcedure(control.beforeUpdate)) {
            QScopedValueRollback<QString> validating(validatingControl_,control.name);
            const auto result = runHandler(controlEventPrefix(control.name)+"_BeforeUpdate");
            eventResults_.insert(("control:"+control.name+":BeforeUpdate").toCaseFolded(),result);
            if (result.cancelled || failed(result)) {
                if (auto* widget = controls_.value(control.name.toCaseFolded()).data()) widget->setFocus();
                return false;
            }
        }
        if (!pendingControls_.contains(name)) return true; // Control.Undo inside validation.
        pendingControls_.remove(name); bufferedTextControls_.remove(name); editOriginals_.remove(name);
        markDirty(name); refreshCalculatedControls();
        rememberControlEdit(controls_.value(control.name.toCaseFolded()).data());
        if (eventProcedure(control.afterUpdate)) {
            const auto result = runHandler(controlEventPrefix(control.name)+"_AfterUpdate");
            eventResults_.insert(("control:"+control.name+":AfterUpdate").toCaseFolded(),result);
            if (failed(result)) return false;
        }
    }
    return true;
}
bool AccessFormPreview::commitControls() {
    const auto pending = pendingControls_;
    for (const auto& control : layout_.controls)
        if (pending.contains(control.name) && !commitControl(control.name)) return false;
    return pendingControls_.isEmpty(); // Retain drafts created by a reentrant handler.
}

bool AccessFormPreview::saveRecord() {
    if (saving_ || recordCommandRunning_ || insertionEventRunning_) return false;
    if (currentRecordIndex_==0 && !layout_.allowAdditions) return !dirty_;
    QScopedValueRollback<bool> saving(saving_,true);
    if (!commitControls()) return false;
    if (!dirty_) return true;
    if(layout_.dataEntry && layout_.allowAdditions && currentRecordIndex_==0 && dataEntryRows_.size()>=10000) {
        status_->setText("Die native Eingabesitzung ist auf 10000 Datensätze begrenzt."); return false;
    }
    const auto table = layout_.recordSource.trimmed();
    if (!dataSourceAvailable_ || !directTable_) return false;
    const auto before = formEvent("BeforeUpdate","Form_BeforeUpdate");
    if (before.cancelled || failed(before)) return false;
    const bool userInsertion=currentRecordIndex_==0 && insertionStarted_;
    QSet<QString> generated;
    try { if (currentRecordIndex_==0) generated=linaccess::generatedAccessColumns(db_,table); }
    catch (const std::exception& error) { status_->setText(QString::fromUtf8(error.what())); return false; }
    QStringList columns,placeholders;
    QVariantList values;
    for (const auto& control : layout_.controls) {
        const int fieldIndex = columnIndex(control.controlSource,dataColumns_);
        if(control.type=="ListBox" && control.multiSelect) continue;
        auto* widget = controls_.value(control.name.toCaseFolded()).data();
        if (!widget || fieldIndex < 0) continue;
        const auto column = sqlIdentifier(dataColumns_[fieldIndex]);
        if (columns.contains(column)) continue;
        try {
            const auto value=getFieldValue(control.name.toUtf8().toStdString());
            if (generated.contains(dataColumns_[fieldIndex].toCaseFolded()) && value.kind==linaccess::VbaValue::Kind::Null) continue;
            columns << column; placeholders << "?";
            values << linaccess::sqlFieldValue(value,widget->property("la_sql_type").toString());
        } catch (const std::exception& error) { status_->setText(QString::fromUtf8(error.what())); return false; }
    }
    if (columns.empty() && currentRecordIndex_>0) return false;
    QSqlQuery query(db_);
    const int row = currentRecordIndex_-1;
    const bool updating=row>=0 && row<recordRows_.size();
    const auto previousId=updating ? rowIds_[row] : QVariant{};
    QVariant savedId;
    if (updating) {
        savedId = rowIds_[row]; QStringList assignments;
        for (const auto& column : columns) assignments << column + "=?";
        query.prepare("UPDATE " + sqlIdentifier(table) + " SET " + assignments.join(',') + " WHERE rowid=? RETURNING rowid");
        values << savedId;
    } else if (columns.empty()) query.prepare("INSERT INTO "+sqlIdentifier(table)+" DEFAULT VALUES");
    else query.prepare("INSERT INTO " + sqlIdentifier(table) + " (" + columns.join(',') + ") VALUES (" + placeholders.join(',') + ")");
    for (const auto& value : values) query.addBindValue(value);
    if (!query.exec()) { status_->setText("SQL: " + query.lastError().text()); return false; }
    if(updating) {
        if(!query.next()) { status_->setText("Der Datensatz wurde nicht gespeichert; die Zeile ist nicht verfügbar."); return false; }
        savedId=query.value(0);
        if(query.next() || query.lastError().isValid()) { status_->setText("Die gespeicherte Datensatzidentität ist nicht verfügbar."); return false; }
    } else if (query.numRowsAffected() != 1) {
        status_->setText("Der Datensatz wurde nicht gespeichert; er wurde möglicherweise extern entfernt."); return false;
    }
    if (savedId.isNull()) savedId = query.lastInsertId();
    QVariantMap oldRecord;
    for (int column=0;column<dataColumns_.size();++column)
        oldRecord.insert(dataColumns_[column].toCaseFolded(),row>=0 && row<recordRows_.size() ? recordRows_[row][column] : QVariant{});
    if(layout_.dataEntry) {
        if(!updating && layout_.allowAdditions) dataEntryRows_.insert(savedId.toLongLong());
        else if(updating && previousId!=savedId && dataEntryRows_.remove(previousId.toLongLong())) dataEntryRows_.insert(savedId.toLongLong());
    }
    dirty_ = false;
    const auto snapshot=recordsetSnapshot_;
    QMap<qlonglong,qlonglong> identityChanges;
    if(updating && previousId!=savedId) {
        identityChanges.insert(previousId.toLongLong(),savedId.toLongLong());
        rowIds_[row]=savedId;
    }
    loadRecords();
    if (snapshot && directTable_ && dataSourceAvailable_ && !dataTruncated_) {
        linaccess::refreshDaoSnapshot(snapshot,dataColumns_,dataTypes_,recordRows_,rowIds_,identityChanges);
        recordsetSnapshot_=snapshot;
    }
    const int index = rowIds_.indexOf(savedId)+1;
    recordSelector_->blockSignals(true); recordSelector_->setCurrentIndex(index); recordSelector_->blockSignals(false);
    currentRecordIndex_ = index; loadRecord();
    {
        const auto previous=property("la_after_update_old_values");
        setProperty("la_after_update_old_values",oldRecord);
        const auto restore=qScopeGuard([&] { setProperty("la_after_update_old_values",previous); });
        (void)formEvent("AfterUpdate","Form_AfterUpdate");
    }
    if (userInsertion) (void)formEvent("AfterInsert","Form_AfterInsert");
    return true;
}

void AccessFormPreview::selectRecord(int target) {
    if (recordCommandRunning_ || (target==0 && !layout_.allowAdditions)) {
        const QSignalBlocker blocked(recordSelector_); recordSelector_->setCurrentIndex(currentRecordIndex_); return;
    }
    if (navigating_ || loading_ || target == currentRecordIndex_) return;
    {
        QScopedValueRollback<bool> navigating(navigating_,true);
        const bool newRecord = target == 0;
        const auto id = target > 0 && target <= rowIds_.size() ? rowIds_[target-1] : QVariant{};
        if (!saveRecord()) {
            recordSelector_->blockSignals(true); recordSelector_->setCurrentIndex(currentRecordIndex_); recordSelector_->blockSignals(false);
            return;
        }
        const int destination = newRecord ? 0 : !id.isNull() ? rowIds_.indexOf(id)+1 : qMin(target,recordSelector_->count()-1);
        recordSelector_->blockSignals(true); recordSelector_->setCurrentIndex(destination); recordSelector_->blockSignals(false);
        currentRecordIndex_ = destination; loadRecord();
    }
    if (opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
    // Record navigation is an observable change in an embedded form. Requery
    // itself preserves the parent's displayed snapshot until an explicit recalc.
    for (auto* ancestor=parentWidget();ancestor;ancestor=ancestor->parentWidget())
        if (auto* parent=dynamic_cast<AccessFormPreview*>(ancestor);parent && parent->initialized_ && !parent->loading_) parent->refreshCalculatedControls();
}

bool AccessFormPreview::closeDocuments(const QVector<AccessFormPreview*>& roots,AccessFormPreview* closingRoot) {
    QVector<QPointer<AccessFormPreview>> documents;
    const std::function<void(AccessFormPreview*)> collect=[&](AccessFormPreview* document) {
        if (!document || document->closed_ || document->openCancelled_ || documents.contains(document)) return;
        documents.append(document);
        for (const auto& child : document->subforms_) if (child) collect(child);
    };
    for (auto* root : roots) collect(root);
    for (const auto& document : documents) if (document->closing_) return false;
    for (const auto& document : documents) document->closing_=true;
    const auto reset=qScopeGuard([&] { for (const auto& document : documents) if (document) document->closing_=false; });
    for (const auto& document : documents) if (!document || !document->saveRecord()) return false;
    for (const auto& document : documents) {
        if (!document) return false;
        const auto unload=document->formEvent("OnUnload","Form_Unload");
        if (unload.cancelled || failed(unload)) return false;
    }
    if (activeDocument() && documents.contains(activeDocument())) {
        const auto active=activeDocument(); activeDocument().clear();
        (void)active->formEvent("OnDeactivate","Form_Deactivate");
    }
    for (auto it=documents.crbegin();it!=documents.crend();++it) if (*it) {
        (void)(*it)->formEvent("OnClose","Form_Close");
        if (!*it) continue;
        (*it)->closed_=true;
        if ((*it)->formTimer_) (*it)->formTimer_->stop();
        if (*it!=closingRoot && (*it)->isWindow()) (*it)->QDialog::reject();
    }
    return true;
}
void AccessFormPreview::setDocumentVisible(bool visible) {
    using linaccess::VbaRuntimeError;
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    if(closed_ || openCancelled_) throw VbaRuntimeError(2450,"Form is closed");
    const bool report=layout_.kind==linaccess::AccessDocumentKind::Report;
    const bool firstShow=!hasActivated_ && isHidden() && visible;
    if(!opened_ || (!hasActivated_ && !firstShow) || (!report && !initialCurrentDispatched_) || !isWindow() || closing_ || recordCommandRunning_ || saving_ || navigating_ || loading_ || visibilityChanging_)
        throw VbaRuntimeError(3251,"Changing document visibility during this lifecycle phase is not implemented");
    if(visible==!isHidden()) return;
    auto* modal=QApplication::activeModalWidget();
    if((isModal() && !nativeDialog_) || (modal && modal!=this)) throw VbaRuntimeError(3251,"Modal form visibility change is not implemented for this context");
    if(!visible && (dirty_ || !pendingControls_.isEmpty() || !bufferedTextControls_.isEmpty()))
        throw VbaRuntimeError(3251,"Hiding a form with pending edits is not implemented");
    QScopedValueRollback<bool> changing(visibilityChanging_,true);
    QDialog::setVisible(visible);
    if(firstShow) { activateWindow(); activateDocument(); }
}
bool AccessFormPreview::prepareClose() {
    if (closed_ || openCancelled_) return true;
    return closeDocuments({this},this);
}
bool AccessFormPreview::closeProjectDocuments(const QSqlDatabase& db) {
    for(auto* widget:QApplication::topLevelWidgets())
        if(auto* guard=dynamic_cast<linaccess::NativeProjectCloseGuard*>(widget);guard && !guard->allowProjectClose(db.connectionName())) return false;
    const auto key=projectKey(db);
    if (closingProjects().contains(key)) return false;
    closingProjects().insert(key);
    const auto reset=qScopeGuard([&] { closingProjects().remove(key); });
    QVector<AccessFormPreview*> roots;
    for (auto* widget : QApplication::topLevelWidgets())
        if (auto* document=dynamic_cast<AccessFormPreview*>(widget); document && projectKey(document->db_)==key)
            roots.append(document);
    return closeDocuments(roots);
}
bool AccessFormPreview::closeApplicationDocuments() {
    for(auto* widget:QApplication::topLevelWidgets())
        if(auto* guard=dynamic_cast<linaccess::NativeProjectCloseGuard*>(widget);guard && !guard->allowProjectClose({})) return false;
    QVector<AccessFormPreview*> roots;
    QSet<QString> keys;
    for (auto* widget : QApplication::topLevelWidgets())
        if (auto* document=dynamic_cast<AccessFormPreview*>(widget); document && !document->closed_ && !document->openCancelled_) {
            roots.append(document); keys.insert(projectKey(document->db_));
        }
    for (const auto& key : keys) if (closingProjects().contains(key)) return false;
    closingProjects().unite(keys);
    const auto reset=qScopeGuard([&] { closingProjects().subtract(keys); });
    return closeDocuments(roots);
}
void AccessFormPreview::quitApplication() {
    if (!closeApplicationDocuments()) throw linaccess::VbaRuntimeError(2501,"Quit action cancelled");
    linaccess::requestVbaExecutionStop();
    QApplication::quit();
}
void AccessFormPreview::closeEvent(QCloseEvent* event) {
    if (!prepareClose()) event->ignore(); else QDialog::closeEvent(event);
}
void AccessFormPreview::accept() { if (prepareClose()) QDialog::accept(); }
void AccessFormPreview::reject() { if (prepareClose()) QDialog::reject(); }

std::optional<std::string> AccessFormPreview::getField(std::string_view name) const {
    auto* widget = controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded()).data();
    if (!widget) throw std::runtime_error("Unbekanntes Access-Control: " + std::string(name));
    if(widget->property("la_default_error").toInt()) throw linaccess::VbaRuntimeError(widget->property("la_default_error").toInt(),"Control default unavailable");
    if(auto* list=dynamic_cast<AccessListBox*>(widget); list && list->multiSelectMode()) return {};
    const auto value = textOf(widget);
    if (widget->property("la_null").toBool()) return std::nullopt;
    return value.toUtf8().toStdString();
}

void AccessFormPreview::ensureWritableControl(std::string_view name) const {
    if (!loading_ && !validatingControl_.isEmpty() && validatingControl_.compare(QString::fromUtf8(name.data(),name.size()),Qt::CaseInsensitive)==0)
        throw linaccess::VbaRuntimeError(2115,"Control cannot change its data during BeforeUpdate");
    if (recordCommandRunning_) throw linaccess::VbaRuntimeError(3251,"Editing during record command unsupported");
    const auto key = QString::fromUtf8(name.data(),name.size());
    for (const auto& control : layout_.controls)
        if (control.name.compare(key,Qt::CaseInsensitive)==0 &&
            (control.controlSource.startsWith('=') || (!control.controlSource.isEmpty() && (!directTable_ || (currentRecordIndex_==0 && !layout_.allowAdditions)))))
            throw linaccess::VbaRuntimeError(2448,"Bound control is read-only");
}
void AccessFormPreview::setField(std::string_view name, std::optional<std::string> value) {
    ensureWritableControl(name);
    auto* widget = controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded()).data();
    if (!widget) throw std::runtime_error("Unbekanntes Access-Control: " + std::string(name));
    if(auto* list=dynamic_cast<AccessListBox*>(widget); list && list->multiSelectMode()) throw linaccess::VbaRuntimeError(3251,"Set multi-selection through Selected, not Value");
    clearDefaultError(widget);
    const bool wasNull = widget->property("la_null").toBool();
    const auto previous = textOf(widget);
    {
        QScopedValueRollback<bool> loading(loading_,true);
        setText(widget,value ? QString::fromUtf8(value->data(),value->size()) : QString{},!value);
        widget->setProperty("la_null",!value.has_value());
        widget->setProperty("la_vba_kind",static_cast<int>(value ? linaccess::VbaValue::Kind::Text : linaccess::VbaValue::Kind::Null));
    }
    const auto assignedName=QString::fromUtf8(name.data(),name.size());
    bufferedTextControls_.remove(assignedName); pendingControls_.remove(assignedName); editOriginals_.remove(assignedName);
    rememberControlEdit(widget);
    if (wasNull != !value.has_value() || previous != textOf(widget)) markDirty(assignedName);
    refreshTextAlignments();
}

linaccess::VbaValue AccessFormPreview::getFieldValue(std::string_view name) const {
    const auto key=QString::fromUtf8(name.data(),name.size());
    for (const auto& buffered:bufferedTextControls_)
        if (buffered.compare(key,Qt::CaseInsensitive)==0 && validatingControl_.compare(buffered,Qt::CaseInsensitive)!=0)
            if (const auto original=editOriginals_.value(buffered)) return *original;
    for (const auto& control : layout_.controls)
        if (control.name.compare(key,Qt::CaseInsensitive)==0 && control.controlSource.trimmed().startsWith('=')) {
            const auto cached=calculatedValues_.value(control.name.toCaseFolded());
            if (cached && calculatingControls_.isEmpty()) return *cached;
            return calculatedControlValue(control.name);
        }
    auto* widget = controls_.value(key.toCaseFolded()).data();
    if (!widget) {
        const auto index=columnIndex(key,dataColumns_);
        if (index<0) throw linaccess::VbaRuntimeError(2465,"Unknown form control or field");
        const auto row=currentRecordIndex_-1;
        return row>=0 && row<recordRows_.size() ? linaccess::qtFieldValue(recordRows_[row][index],dataTypes_.value(key.toCaseFolded())) : linaccess::VbaValue{};
    }
    if(auto* list=dynamic_cast<AccessListBox*>(widget); list && list->multiSelectMode()) return {};
    if(widget->property("la_default_error").toInt()) throw linaccess::VbaRuntimeError(widget->property("la_default_error").toInt(),"Control default unavailable");
    if (auto* calendar=qobject_cast<QCalendarWidget*>(widget)) {
        if (calendar->property("la_calendar_bound").toBool()) throw linaccess::VbaRuntimeError(3251,"Bound calendar adapter is unsupported");
        return widget->property("la_null").toBool() ? linaccess::VbaValue{} : linaccess::vbaDateValue(linaccess::VbaValue::string(calendar->selectedDate().toString(Qt::ISODate).toStdString()));
    }
    if (qobject_cast<QGroupBox*>(widget)) {
        if (widget->property("la_option_error").toInt()) throw linaccess::VbaRuntimeError(3251,"Option group default expression unsupported");
        return widget->property("la_null").toBool() ? linaccess::VbaValue{} : linaccess::VbaValue::numeric(widget->property("la_option_value").toInt());
    }
    if (auto* button=qobject_cast<QAbstractButton*>(widget);button && !button->property("la_option_group_name").toString().isEmpty()) return linaccess::VbaValue::logical(button->isChecked());
    if (!qobject_cast<QLineEdit*>(widget) && !qobject_cast<QComboBox*>(widget) && !qobject_cast<QCheckBox*>(widget) && !dynamic_cast<AccessListBox*>(widget))
        throw linaccess::VbaRuntimeError(438,"Control has no Value property");
    if (static_cast<linaccess::VbaValue::Kind>(widget->property("la_vba_kind").toInt())==linaccess::VbaValue::Kind::Error)
        return linaccess::VbaValue::errorValue(widget->property("la_error_number").toInt());
    if (widget->property("la_null").toBool()) return {};
    if(auto* list=dynamic_cast<AccessListBox*>(widget); list && list->currentIndex().isValid())
        return linaccess::qtFieldValue(list->boundData(list->currentIndex().row()),widget->property("la_sql_type").toString());
    if (auto* combo=qobject_cast<QComboBox*>(widget);combo && combo->property("la_rows_loaded").toBool() && combo->currentIndex()>=0 && combo->currentText()==combo->itemText(combo->currentIndex()))
        return linaccess::qtFieldValue(combo->currentData(),widget->property("la_sql_type").toString());
    auto type = widget->property("la_sql_type").toString();
    if (type.isEmpty()) {
        const auto kind = static_cast<linaccess::VbaValue::Kind>(widget->property("la_vba_kind").toInt());
        if (kind == linaccess::VbaValue::Kind::Empty) return linaccess::VbaValue::empty();
        if (kind == linaccess::VbaValue::Kind::Date) return linaccess::VbaValue::dateSerial(widget->property("la_vba_date_serial").toDouble());
        if (kind == linaccess::VbaValue::Kind::Currency) return linaccess::convertCurrency(linaccess::VbaValue::string(textOf(widget).toStdString()));
        if (kind == linaccess::VbaValue::Kind::Number) type = "REAL";
        if (kind == linaccess::VbaValue::Kind::Boolean || qobject_cast<QCheckBox*>(widget)) type = "BOOLEAN";
    }
    return linaccess::qtFieldValue(textOf(widget),type);
}
linaccess::VbaValue AccessFormPreview::calculatedControlValue(const QString& name) const {
    const auto key=name.toCaseFolded();
    auto* widget=controls_.value(key).data();
    if (!widget) throw linaccess::VbaRuntimeError(2465,"Calculated control unavailable");
    const QSignalBlocker blocked(widget);
    const auto remember=[&](linaccess::VbaValue value) {
        calculatedValues_.insert(key,std::make_shared<linaccess::VbaValue>(value)); return value;
    };
    try {
        if (calculatingControls_.contains(key) || calculatingControls_.size()>=64)
            throw linaccess::VbaRuntimeError(2455,"Circular calculated control reference");
        QScopedValueRollback<QSet<QString>> guard(calculatingControls_);
        calculatingControls_.insert(key);
        QString expression;
        for (const auto& control : layout_.controls)
            if (control.name.compare(name,Qt::CaseInsensitive)==0) expression=control.controlSource.trimmed().mid(1);
        PreviewVbaHost host(*const_cast<AccessFormPreview*>(this));
        const auto value=linaccess::evaluateAccessExpression(expression.toUtf8().toStdString(),host);
        if (value.kind==linaccess::VbaValue::Kind::Error) {
            setText(widget,"#Fehler ("+QString::number(value.errorNumber)+")");
            widget->setProperty("la_null",false); widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
            widget->setProperty("la_error_number",value.errorNumber); widget->setProperty("la_calculation_error",value.errorNumber);
            return remember(value);
        }
        const auto converted=linaccess::sqlFieldValue(value,"");
        setText(widget,converted.toString(),converted.isNull()); widget->setProperty("la_null",converted.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
        widget->setProperty("la_calculation_error",0);
        return remember(value);
    } catch (const std::exception& error) {
        const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error);
        const int number=typed ? typed->number : 5;
        setText(widget,"#Fehler ("+QString::number(number)+")"); widget->setProperty("la_calculation_error",number);
        widget->setProperty("la_null",false); widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Error));
        widget->setProperty("la_error_number",number);
        return remember(linaccess::VbaValue::errorValue(number));
    }
}
void AccessFormPreview::refreshCalculatedControls() {
    for (const auto& control : layout_.controls)
        if (control.controlSource.trimmed().startsWith('=')) {
            try { (void)calculatedControlValue(control.name); } catch (const std::exception&) { /* Error is visible on the control. */ }
        }
    refreshTextAlignments();
}
linaccess::VbaValue AccessFormPreview::controlColumnValue(const QString& name,const std::vector<linaccess::VbaValue>& arguments) {
    if (arguments.empty() || arguments.size()>2) throw linaccess::VbaRuntimeError(450,"Column requires column and optional row");
    auto* widget=controls_.value(name.toCaseFolded()).data();
    auto* combo=qobject_cast<QComboBox*>(widget);
    auto* list=dynamic_cast<AccessListBox*>(widget);
    if (!combo && !list) throw linaccess::VbaRuntimeError(438,"Control has no Column property");
    if(list && list->multiSelectMode() && arguments.size()==1) throw linaccess::VbaRuntimeError(3251,"Multi-selection Column requires an explicit row");
    const auto index=[](const linaccess::VbaValue& value) {
        if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null column index");
        return linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
    };
    const auto column=index(arguments[0]);
    int count=1;
    for (const auto& control : layout_.controls) if (control.name.compare(name,Qt::CaseInsensitive)==0) count=control.columnCount;
    if (column<0 || column>=count) throw linaccess::VbaRuntimeError(5,"Column index out of range");
    if (arguments.size()==1 && widget->property("la_null").toBool()) return {};
    if (!widget->property("la_rows_loaded").toBool()) requeryControl(name.toUtf8().toStdString());
    const auto row=arguments.size()==2 ? index(arguments[1]) : list ? list->currentIndex().row() : combo->currentIndex();
    if (row<0 && arguments.size()==1) return {};
    if (row<0 || row>=(list ? list->rowCount() : combo->count())) throw linaccess::VbaRuntimeError(5,"Column row index out of range");
    const auto values=list ? list->rowValues(row) : combo->itemData(row,Qt::UserRole+1).toList();
    if (column>=values.size()) throw linaccess::VbaRuntimeError(5,"Column not present in row source");
    return linaccess::qtFieldValue(values[column]);
}
void AccessFormPreview::setFieldValue(std::string_view name,const linaccess::VbaValue& value,bool initializingDefault) {
    if(!initializingDefault) ensureWritableControl(name);
    else if(!loading_) throw linaccess::VbaRuntimeError(3251,"Default initialization requires loading scope");
    auto* widget = controls_.value(QString::fromUtf8(name.data(),name.size()).toCaseFolded()).data();
    if (!widget) throw linaccess::VbaRuntimeError(2465,"Unknown form control");
    if(auto* list=dynamic_cast<AccessListBox*>(widget); list && list->multiSelectMode()) throw linaccess::VbaRuntimeError(3251,"Set multi-selection through Selected, not Value");
    if (auto* calendar=qobject_cast<QCalendarWidget*>(widget)) {
        if (calendar->property("la_calendar_bound").toBool()) throw linaccess::VbaRuntimeError(3251,"Bound calendar adapter is unsupported");
        if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Calendar requires a date");
        const auto text=value.kind==linaccess::VbaValue::Kind::Text ? QString::fromStdString(value.text) : linaccess::sqlFieldValue(linaccess::vbaDateValue(value),"DATETIME").toString();
        const auto date=QDate::fromString(text.left(10),Qt::ISODate);
        if (!date.isValid() || (text.size()!=10 && (text.size()!=19 || !QDateTime::fromString(text,"yyyy-MM-dd HH:mm:ss").isValid()))
            || date<calendar->minimumDate() || date>calendar->maximumDate()) throw linaccess::VbaRuntimeError(13,"Invalid calendar date");
        if (text.size()==19 && QDateTime::fromString(text,"yyyy-MM-dd HH:mm:ss").time()!=QTime(0,0))
            throw linaccess::VbaRuntimeError(3251,"Calendar time components are unsupported");
        calendar->setSelectedDate(date); calendar->setProperty("la_null",false); markDirty(QString::fromUtf8(name.data(),name.size())); return;
    }
    if (value.kind==linaccess::VbaValue::Kind::Error) {
        bool unbound=false;
        for (const auto& control : layout_.controls)
            if (control.name.compare(QString::fromUtf8(name.data(),name.size()),Qt::CaseInsensitive)==0) unbound=control.controlSource.isEmpty();
        if (!unbound || !qobject_cast<QLineEdit*>(widget)) throw linaccess::VbaRuntimeError(13,"Error Variant cannot be stored in a bound field");
        clearDefaultError(widget);
        const QSignalBlocker blocked(widget);
        setText(widget,"#Fehler ("+QString::number(value.errorNumber)+")");
        widget->setProperty("la_null",false); widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
        widget->setProperty("la_error_number",value.errorNumber);
        const auto assignedName=QString::fromUtf8(name.data(),name.size());
        bufferedTextControls_.remove(assignedName); pendingControls_.remove(assignedName); editOriginals_.remove(assignedName);
        rememberControlEdit(widget); refreshCalculatedControls(); return;
    }
    auto assigned=value;
    if (qobject_cast<QGroupBox*>(widget) && value.kind!=linaccess::VbaValue::Kind::Null) {
        const auto number=value.asNumber();
        if (!std::isfinite(number) || number<std::numeric_limits<int>::min() || number>std::numeric_limits<int>::max()) throw linaccess::VbaRuntimeError(6,"Option group value overflow");
        if (std::floor(number)!=number) throw linaccess::VbaRuntimeError(3251,"Fractional option group values unsupported");
        assigned=linaccess::VbaValue::numeric(number); widget->setProperty("la_option_error",0);
    }
    if (qobject_cast<QGroupBox*>(widget)) widget->setProperty("la_option_error",0);
    if (!widget->property("la_option_group_name").toString().isEmpty()) throw linaccess::VbaRuntimeError(3251,"Individual grouped option assignment unsupported");
    const auto type = widget->property("la_sql_type").toString();
    const auto converted = linaccess::sqlFieldValue(assigned,type);
    clearDefaultError(widget);
    const auto before = textOf(widget); const bool wasNull = widget->property("la_null").toBool();
    {
        QScopedValueRollback<bool> loading(loading_,true);
        setText(widget,qobject_cast<QLineEdit*>(widget) ? linaccess::qtFieldText(converted,type) : converted.toString(),converted.isNull()); widget->setProperty("la_null",converted.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
        if(value.kind==linaccess::VbaValue::Kind::Date) widget->setProperty("la_vba_date_serial",value.number);
    }
    if(initializingDefault) { refreshTextAlignments(); return; }
    const auto assignedName=QString::fromUtf8(name.data(),name.size());
    bufferedTextControls_.remove(assignedName); pendingControls_.remove(assignedName); editOriginals_.remove(assignedName);
    rememberControlEdit(widget);
    if (before != textOf(widget) || wasNull != converted.isNull()) markDirty(assignedName);
    refreshCalculatedControls();
}

void AccessFormPreview::requeryControl(std::string_view name) {
    const auto target = QString::fromUtf8(name.data(),name.size());
    if (auto* child=subforms_.value(target.toCaseFolded()).data()) {
        if (child->closed_) throw linaccess::VbaRuntimeError(2455,"Subform is closed");
        (void)makePreviewFormObject(*child)->member("Requery",{}); return;
    }
    if (subformErrors_.contains(target.toCaseFolded()))
        throw linaccess::VbaRuntimeError(subformErrors_.value(target.toCaseFolded()),"Subform unavailable");
    auto* widget = controls_.value(target.toCaseFolded()).data();
    if (!widget) throw linaccess::VbaRuntimeError(2465,"Control unavailable");
    for (const auto& control : layout_.controls)
        if (control.name.compare(target,Qt::CaseInsensitive)==0 && control.type=="TextBox" && control.controlSource.trimmed().startsWith('=')) {
            (void)calculatedControlValue(control.name); return;
        }
    for (const auto& control : layout_.controls)
        if (control.name.compare(target,Qt::CaseInsensitive)==0 && (control.type=="CommandButton" || (control.type=="TextBox" && control.controlSource.isEmpty()))) {
            // These controls have no record/row source. Invalidate their native
            // presentation while preserving the value and form event state.
            widget->update(); return;
        }
    for (const auto& control : layout_.controls)
        if (control.name.compare(target,Qt::CaseInsensitive)==0 && !control.controlSource.isEmpty()
            && !control.controlSource.startsWith('=') && columnIndex(control.controlSource,dataColumns_)>=0
            && !qobject_cast<QComboBox*>(widget) && !dynamic_cast<AccessListBox*>(widget)) {
            (void)makePreviewFormObject(*this)->member("Requery",{}); return;
        }
    auto* combo = qobject_cast<QComboBox*>(widget);
    auto* list=dynamic_cast<AccessListBox*>(widget);
    if (!combo && !list) throw linaccess::VbaRuntimeError(438,"Requery for this control type is unsupported");
    if(list && list->multiSelectMode()==1 && !list->selectedRows().isEmpty()) throw linaccess::VbaRuntimeError(3251,"Simple-mode selection preservation on Requery is not implemented");
    for (const auto& control : layout_.controls) {
        if (control.name.compare(target,Qt::CaseInsensitive)!=0) continue;
        const auto source = control.rowSource.trimmed();
        linaccess::AccessQueryResult result;
        const auto sourceType=control.rowSourceType.toCaseFolded();
        const bool values=sourceType=="value list" || sourceType=="wertliste";
        const bool fields=sourceType=="field list" || sourceType=="feldliste";
        if(fields && control.columnHeads) throw linaccess::VbaRuntimeError(3251,"Field List ColumnHeads semantics are not implemented");
        if (!values && !fields && !sourceType.isEmpty() && sourceType!="table/query" && sourceType!="tabelle/abfrage")
            throw linaccess::VbaRuntimeError(3251,"Row source type unsupported");
        if (values) result=linaccess::parseAccessValueList(source,control.columnCount);
        else if (!source.isEmpty()) {
            linaccess::AccessQueryPlan plan;
            if (db_.tables().contains(source,Qt::CaseInsensitive)) plan.sql = "SELECT * FROM " + sqlIdentifier(source);
            else plan = linaccess::compileAccessQuery(
                source.startsWith("SELECT",Qt::CaseInsensitive) || source.startsWith("PARAMETERS",Qt::CaseInsensitive)
                    ? source : linaccess::accessQuerySource(db_,source));
            if (fields) {
                result.columns=QStringList{"Field"};
                for (const auto& name:linaccess::accessQueryColumns(db_,plan)) result.rows.append(QVariantList{name});
            } else result = linaccess::executeAccessQueryPlan(db_,plan,linaccess::bindAccessQueryParameters(plan,{}));
            if (result.rows.size()>10000) throw linaccess::VbaRuntimeError(3035,"Control row limit exceeded");
        }
        if (control.boundColumn<0 || control.columnCount<1 || control.columnCount>255 || (!result.columns.empty() && control.boundColumn>result.columns.size()))
            throw linaccess::VbaRuntimeError(5,"Invalid row source column configuration");
        QStringList headings=result.columns;
        int headingRows=control.columnHeads && !headings.isEmpty() ? 1 : 0;
        if(values && control.columnHeads) {
            headings.clear(); headingRows=result.rows.isEmpty() ? 0 : 1;
            if(!result.rows.isEmpty()) {
                const auto first=result.rows.takeFirst();
                for(int column=0;column<control.columnCount;++column) headings.append(first.value(column).toString());
            }
        }
        while(headings.size()<control.columnCount) headings.append(QString{});
        QVector<int> widths;
        const auto specifications=control.columnWidths.split(';');
        if (specifications.size()>255) throw linaccess::VbaRuntimeError(5,"Too many column widths");
        for (auto specification : specifications) {
            specification=specification.trimmed();
            if (specification.isEmpty()) { widths.append(-1); continue; }
            static const QRegularExpression widthPattern(R"(^([0-9]+(?:[.,][0-9]+)?)\s*(cm|mm|in|pt)?$)",QRegularExpression::CaseInsensitiveOption);
            const auto match=widthPattern.match(specification);
            if (!match.hasMatch()) throw linaccess::VbaRuntimeError(5,"Invalid column width");
            auto decimal=match.captured(1); decimal.replace(',','.');
            const auto amount=decimal.toDouble(); const auto unit=match.captured(2).toLower();
            const auto scale=unit=="cm" ? 96.0/2.54 : unit=="mm" ? 96.0/25.4 : unit=="in" ? 96.0 : unit=="pt" ? 96.0/72.0 : 1.0/15.0;
            if (!std::isfinite(amount) || amount*scale>32767) throw linaccess::VbaRuntimeError(5,"Column width out of range");
            widths.append(qRound(amount*scale));
        }
        const bool editing=combo && combo->isEditable() && combo->lineEdit()->isModified();
        const auto draft=combo ? combo->currentText() : QString{};
        QVariant old;
        try { old=linaccess::sqlFieldValue(getFieldValue(target.toStdString()),""); }
        catch (const linaccess::VbaRuntimeError& error) {
            // A user can type a label before the first popup has loaded its
            // numeric IDs. Loading choices must not validate/commit that draft.
            if (!editing || (error.number!=13 && error.number!=6)) throw;
            old=combo->currentText();
        }
        const QSignalBlocker blocked(widget);
        auto* model=new QStandardItemModel(widget);
        model->setColumnCount(control.columnCount);
        model->setHorizontalHeaderLabels(headings.mid(0,control.columnCount));
        for (int index=0;index<result.rows.size();++index) {
            const auto& row=result.rows[index];
            const auto bound=control.boundColumn==0 ? QVariant(index) : row.value(control.boundColumn-1);
            for (int column=0;column<control.columnCount;++column) {
                auto* item=new QStandardItem(row.value(column).toString());
                item->setData(bound,Qt::UserRole); item->setData(row,Qt::UserRole+1);
                model->setItem(index,column,item);
            }
        }
        // QComboBox deletes its previous owned model inside setModel.
        // Do not dereference or schedule that freed model afterwards.
        if(list) list->setModel(model); else combo->setModel(model);
        auto* view=list ? static_cast<QTreeView*>(list) : qobject_cast<QTreeView*>(combo->view());
        if (!view) {
            view=new QTreeView(combo); view->setRootIsDecorated(false); view->setHeaderHidden(true);
            view->setSelectionBehavior(QAbstractItemView::SelectRows); combo->setView(view);
        }
        view->setFont(widget->font());
        view->setHeaderHidden(!control.columnHeads);
        int displayColumn=-1;
        for (int column=0;column<control.columnCount;++column) {
            const auto width=widths.value(column,-1);
            view->setColumnHidden(column,width==0);
            if (width>0) view->setColumnWidth(column,width);
            if (width!=0 && displayColumn<0) displayColumn=column;
        }
        if(list) list->setDisplayColumn(displayColumn<0 ? 0 : displayColumn); else combo->setModelColumn(displayColumn<0 ? 0 : displayColumn);
        widget->setProperty("la_rows_loaded",true);
        widget->setProperty("la_rows_error",0);
        widget->setProperty("la_header_rows",headingRows);
        if(list && widget->property("la_list_source_tooltip").isValid()) {
            widget->setToolTip(widget->property("la_list_source_tooltip").toString()); widget->setProperty("la_list_source_tooltip",QVariant{});
        }
        if (editing) {
            combo->setCurrentIndex(-1); combo->setEditText(draft); combo->setProperty("la_null",false);
        } else { setText(widget,old.toString(),old.isNull()); widget->setProperty("la_null",old.isNull()); }
        if (combo && combo->isEditable()) combo->lineEdit()->setModified(editing);
        applyForeColor(widget,control);
        if(list && list->multiSelectMode()) { pendingControls_.remove(target); editOriginals_.remove(target); }
        return;
    }
}

std::optional<linaccess::VbaValue> AccessFormPreview::getModuleMember(std::string_view name,bool publicOnly) {
    if (closed_ || openCancelled_) throw linaccess::VbaRuntimeError(2450,"CodeBehind instance is unavailable");
    const auto source=vba_.toUtf8().toStdString();
    auto found=moduleStates_.find(source);
    if (found==moduleStates_.end()) found=moduleStates_.emplace(source,linaccess::createVbaModuleState(source)).first;
    const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded().toStdString();
    const auto& frame=found->second;
    if (publicOnly && !frame->publicVariables.contains(key)) return std::nullopt;
    if (const auto field=frame->values.find(key);field!=frame->values.end()) return field->second;
    return std::nullopt;
}
bool AccessFormPreview::setModuleMember(std::string_view name,const linaccess::VbaValue& value,bool publicOnly) {
    if (!getModuleMember(name,publicOnly)) return false;
    const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded().toStdString();
    const auto& frame=moduleStates_.at(vba_.toUtf8().toStdString());
    frame->values.at(key)=linaccess::coerceVbaModuleValue(value,frame->types.at(key)); return true;
}

void AccessFormPreview::closeOtherDocument(int type,const std::optional<QString>& name,int save) {
    if (type!=-1 && type!=2 && type!=3) throw linaccess::VbaRuntimeError(3251,"Closing this object type is unsupported");
    if (save<0 || save>2) throw linaccess::VbaRuntimeError(5,"Invalid Close save option");
    if (hasActivated_ && databaseWindowActive() && (!name || name->isEmpty()))
        throw linaccess::VbaRuntimeError(3251,"Default database-window Close target is not implemented");
    const auto matches=[&](AccessFormPreview* document) {
        return document && document->isWindow() && !document->closed_ && !document->openCancelled_
            && document->db_.connectionName()==db_.connectionName()
            && (type==-1 || (document->layout_.kind==linaccess::AccessDocumentKind::Report ? 3 : 2)==type);
    };
    AccessFormPreview* target=nullptr;
    if (name && !name->isEmpty()) {
        for (auto* widget:QApplication::topLevelWidgets()) {
            auto* document=dynamic_cast<AccessFormPreview*>(widget);
            if (matches(document) && document->formName_.compare(*name,Qt::CaseInsensitive)==0) {
                if (target) throw linaccess::VbaRuntimeError(5,"Close target is ambiguous");
                target=document;
            }
        }
    } else {
        if (matches(activeDocument()) && !activeDocument()->isHidden()) target=activeDocument();
        else if (matches(this) && (!hasActivated_ || !isHidden())) target=this;
        else if(hasActivated_ && isHidden()) throw linaccess::VbaRuntimeError(2046,"Close requires an active native document");
    }
    if (!target) throw linaccess::VbaRuntimeError(type==3 ? 2451 : 2450,"Close target is unavailable");
    // The native preview has no pending design-edit session. Save options apply
    // to object design, not record changes; normal record validation still runs.
    if (target->closing_ || !target->close()) throw linaccess::VbaRuntimeError(2501,"Close action cancelled");
}

void AccessFormPreview::openOtherForm(const QString& name,const QString& where,std::optional<QString> openArgs,int windowMode,int dataMode) {
    openOtherDocument(linaccess::AccessDocumentKind::Form,name,where,std::move(openArgs),windowMode,dataMode);
}

void AccessFormPreview::openOtherDocument(linaccess::AccessDocumentKind kind,const QString& name,const QString& where,std::optional<QString> openArgs,int windowMode,int dataMode) {
    if(windowMode<0 || windowMode>3) throw linaccess::VbaRuntimeError(3251,"Native window mode is not implemented");
    if(dataMode<-1 || dataMode>2) throw linaccess::VbaRuntimeError(3251,"Native data mode is not implemented");
    const bool report=kind==linaccess::AccessDocumentKind::Report;
    if(report && dataMode!=-1) throw linaccess::VbaRuntimeError(3251,"DataMode requires a form");
    if (closing_ || closingProjects().contains(projectKey(db_))) throw linaccess::VbaRuntimeError(2501,"Project is closing");
    AccessFormPreview* existing=nullptr;
    for (auto* widget:QApplication::topLevelWidgets()) {
        auto* document=dynamic_cast<AccessFormPreview*>(widget);
        if (!document || !document->opened_ || document->closed_ || document->openCancelled_
            || document->layout_.kind!=kind
            || document->db_.connectionName()!=db_.connectionName() || document->db_.databaseName()!=db_.databaseName()
            || document->formName_.compare(name,Qt::CaseInsensitive)!=0) continue;
        if (existing) throw linaccess::VbaRuntimeError(5,"Document target instance is ambiguous");
        existing=document;
    }
    if (existing) {
        if(dataMode!=-1) throw linaccess::VbaRuntimeError(3251,"Changing data mode of an open document is not implemented");
        if(windowMode==3) throw linaccess::VbaRuntimeError(3251,"acDialog reuse of an open document is not implemented");
        if(openArgs && existing->openArgs_!=openArgs) throw linaccess::VbaRuntimeError(3251,"Changing OpenArgs of an open document is not implemented");
        if (!where.trimmed().isEmpty()) {
            if (existing->layout_.recordSource.trimmed().isEmpty()) throw linaccess::VbaRuntimeError(3251,"WhereCondition requires a bound document");
            existing->setFilterState(where,true);
        }
        if(windowMode==1) {
            if(!existing->isHidden()) existing->setDocumentVisible(false);
        } else if(windowMode==2) { existing->showMinimized(); existing->activateDocument(); }
        else { existing->show(); existing->raise(); existing->activateWindow(); existing->activateDocument(); }
        return;
    }
    QSqlQuery query(db_);
    query.prepare("SELECT name,source_bytes FROM _la_access_artifacts "
                  "WHERE kind=? AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1");
    query.addBindValue(report ? "report" : "form"); query.addBindValue(name);
    if (!query.exec() || !query.next()) {
        throw linaccess::VbaRuntimeError(report ? 2103 : 2102,"Document target is not imported");
    }
    const auto canonical=query.value(0).toString();
    const auto form = query.value(1).toByteArray(); query.finish();
    QSqlQuery vba(db_);
    vba.prepare("SELECT source_bytes FROM _la_access_artifacts "
                "WHERE kind='vba' AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1");
    vba.addBindValue((report ? "Report_" : "Form_") + canonical);
    const auto code = vba.exec() && vba.next() ? vba.value(0).toByteArray() : QByteArray{}; vba.finish();
    if (linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(form)).kind!=kind)
        throw linaccess::VbaRuntimeError(report ? 2103 : 2102,"Document target kind mismatch");
    // Preserve the actual project UI without owning the new document through
    // its invoking form. Project close validates every independent document.
    QWidget* projectOwner=nullptr;
    for (auto* ancestor=parentWidget();ancestor;ancestor=ancestor->parentWidget())
        if (qobject_cast<QMainWindow*>(ancestor)) { projectOwner=ancestor; break; }
    auto nested=std::make_unique<AccessFormPreview>(canonical,form,code,db_,projectOwner,false,where,std::move(openArgs),windowMode,dataMode);
    const auto validate=[&] {
        for (const auto* event:{"OnOpen","OnLoad","OnActivate","OnCurrent","OnNoData"})
            if (const auto outcome=nested->eventResult(event);outcome && failed(*outcome))
                throw linaccess::VbaRuntimeError(outcome->errorNumber ? outcome->errorNumber : 3251,"Document initialization failed");
        if (nested->cancelledByOpen() || nested->closed_) throw linaccess::VbaRuntimeError(2501,"Document open action cancelled");
        if (!nested->layout_.recordSource.trimmed().isEmpty() && !nested->hasDataSource())
            throw linaccess::VbaRuntimeError(nested->dataSourceErrorNumber_ ? nested->dataSourceErrorNumber_ : 3078,"Document record source unavailable");
        if (!where.trimmed().isEmpty() && nested->layout_.recordSource.trimmed().isEmpty())
            throw linaccess::VbaRuntimeError(3251,"WhereCondition requires a bound document");
    };
    validate();
    // Layout Modal blocks input; acDialog additionally waits below. Keep native
    // ownership/deferred disposal instead of transferring lifetime to exec().
    if(windowMode==2) { nested->showMinimized(); nested->activateDocument(); }
    else if(windowMode!=1) { nested->show(); nested->activateWindow(); nested->activateDocument(); }
    validate();
    // A handler may close its own window and then enter another Qt event
    // loop (e.g. MsgBox). Retain all closing native windows until the outermost
    // VBA call unwinds, including callers with executing child handlers.
    QObject::connect(nested.get(),&QDialog::finished,nested.get(),[form=nested.get()] { deferFormDisposal(form); });
    auto* document=nested.release();
    if(windowMode==3) waitForNativeDialog(document);
}

void AccessFormPreview::rememberControlEdit(QObject* watched) {
    for (const auto& definition : layout_.controls) {
        auto* widget=controls_.value(definition.name.toCaseFolded()).data();
        const auto* combo=qobject_cast<QComboBox*>(widget);
        if (!widget || (widget!=watched && (!combo || combo->lineEdit()!=watched))) continue;
        if (pendingControls_.contains(definition.name) && editOriginals_.contains(definition.name)) return;
        try {
            ensureWritableControl(definition.name.toStdString());
            editOriginals_.insert(definition.name,std::make_shared<linaccess::VbaValue>(getFieldValue(definition.name.toStdString())));
        } catch (const std::exception&) { editOriginals_.remove(definition.name); }
        break;
    }
}
void AccessFormPreview::updateDirtyFromRecord() {
    if (!directTable_ || !dataSourceAvailable_) return;
    dirty_=false;
    const int row=currentRecordIndex_-1;
    for (const auto& definition : layout_.controls) {
        const int column=columnIndex(definition.controlSource,dataColumns_);
        if(definition.type=="ListBox" && definition.multiSelect) continue;
        if (column<0 || !controls_.value(definition.name.toCaseFolded())) continue;
        if (bufferedTextControls_.contains(definition.name)) { dirty_=true; continue; }
        const auto expected=row>=0 && row<recordRows_.size() ? recordRows_[row][column] : QVariant{};
        try {
            const auto actual=linaccess::sqlFieldValue(getFieldValue(definition.name.toStdString()),dataTypes_.value(dataColumns_[column].toCaseFolded()));
            if (actual.isNull()!=expected.isNull() || (!actual.isNull() && actual!=expected)) dirty_=true;
        } catch (const std::exception&) { dirty_=true; }
    }
    if (!dirty_) insertionStarted_=false;
    if (!closed_ && !openCancelled_) applyControlStates();
}
bool AccessFormPreview::undoControlEdit(const QString& name) {
    for (const auto& definition:layout_.controls) {
        if (definition.name.compare(name,Qt::CaseInsensitive)!=0) continue;
        const auto original=editOriginals_.value(definition.name);
        if (!pendingControls_.contains(definition.name) || !original) return false;
        {
            QScopedValueRollback<bool> guard(loading_,true);
            const int column=columnIndex(definition.controlSource,dataColumns_);
            for (const auto& target:layout_.controls)
                if (target.name==definition.name || (column>=0 && columnIndex(target.controlSource,dataColumns_)==column)) {
                    setFieldValue(target.name.toStdString(),*original);
                    pendingControls_.remove(target.name);
                }
        }
        updateDirtyFromRecord(); refreshCalculatedControls(); return true;
    }
    return false;
}
void AccessFormPreview::undoKeyboardEdit() {
    const auto* focus=QApplication::focusWidget();
    for (const auto& definition:layout_.controls) {
        auto* widget=controls_.value(definition.name.toCaseFolded()).data();
        const auto* combo=qobject_cast<QComboBox*>(widget);
        if (!widget || (widget!=focus && (!combo || combo->lineEdit()!=focus))) continue;
        if (undoControlEdit(definition.name)) return;
        break;
    }
    if (dirty_) loadRecord();
}
void AccessFormPreview::openFindDialog() {
    using linaccess::VbaRuntimeError;
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    if(layout_.kind!=linaccess::AccessDocumentKind::Form || !directTable_ || !dataSourceAvailable_ || dataTruncated_)
        throw VbaRuntimeError(3251,"Die Suche ist für diese Datenquelle nicht verfügbar.");
    if(!commitControls() || closed_ || openCancelled_) return;
    if(findDialog_ && findDialog_->isVisible()) { findDialog_->raise(); findDialog_->activateWindow(); return; }
    QStringList searchFields; QMap<QString,QString> typedHints;
    for(const auto& name:dataColumns_) {
        const auto type=dataTypes_.value(name.toCaseFolded());
        if(linaccess::nativeFindFieldKind(type)==linaccess::NativeFindFieldKind::Unsupported) continue;
        searchFields<<name;
        if(!linaccess::nativeFindTextType(type)) typedHints.insert(name,linaccess::nativeFindInputHint(type));
    }
    if(searchFields.isEmpty()) throw VbaRuntimeError(3251,"Keine durchsuchbaren Felder vorhanden.");
    QString initialField,initialText;
    const auto* focus=QApplication::focusWidget();
    for(const auto& control:layout_.controls) {
        auto* widget=controls_.value(control.name.toCaseFolded()).data(); const auto* combo=qobject_cast<QComboBox*>(widget);
        if(widget!=focus && (!combo || combo->lineEdit()!=focus)) continue;
        const auto column=columnIndex(control.controlSource,dataColumns_);
        if(column>=0 && searchFields.contains(dataColumns_[column])) initialField=dataColumns_[column];
        if(const auto* line=qobject_cast<const QLineEdit*>(focus); line && line->echoMode()==QLineEdit::Normal) initialText=line->selectedText();
        break;
    }
    auto* dialog=new QDialog(this); findDialog_=dialog;
    dialog->setObjectName("access-find-dialog"); dialog->setWindowTitle("Suchen"); dialog->setAttribute(Qt::WA_DeleteOnClose);
    auto* box=new QVBoxLayout(dialog); auto* fields=new QFormLayout;
    auto* query=new QLineEdit(dialog); query->setMaxLength(std::numeric_limits<int>::max()); query->setText(initialText); query->setObjectName("access-find-text");
    auto* scope=new QComboBox(dialog); scope->setObjectName("access-find-field");
    if(typedHints.size()<searchFields.size()) scope->addItem("Alle Textfelder",QString{});
    for(const auto& name:searchFields) scope->addItem(name,name);
    if(!initialField.isEmpty()) scope->setCurrentIndex(scope->findData(initialField));
    auto* match=new QComboBox(dialog); match->setObjectName("access-find-match"); match->addItems({"Teil des Feldinhalts","Ganzes Feld","Feldanfang"});
    auto* sensitive=new QCheckBox("Groß-/Kleinschreibung beachten",dialog); sensitive->setObjectName("access-find-case");
    auto* direction=new QComboBox(dialog); direction->setObjectName("access-find-direction");
    direction->addItem("Alle Datensätze",2); direction->addItem("Nach unten",1); direction->addItem("Nach oben",0);
    fields->addRow("Suchen nach:",query); fields->addRow("Suchen in:",scope); fields->addRow("Vergleichen:",match);
    fields->addRow("Suchrichtung:",direction);
    box->addLayout(fields); box->addWidget(sensitive);
    auto* hint=new QLabel(dialog); hint->setObjectName("access-find-hint"); hint->setWordWrap(true); box->addWidget(hint);
    const auto updateField=[scope,match,sensitive,hint,typedHints] {
        const auto field=scope->currentData().toString(); const bool exact=typedHints.contains(field);
        if(exact) match->setCurrentIndex(1);
        match->setEnabled(!exact); sensitive->setEnabled(!exact);
        hint->setText(exact ? typedHints.value(field) : QStringLiteral("Text wird wörtlich gesucht."));
    };
    connect(scope,&QComboBox::currentIndexChanged,dialog,[updateField](int) { updateField(); }); updateField();
    auto* result=new QLabel(dialog); result->setObjectName("access-find-result"); result->setWordWrap(true); box->addWidget(result);
    auto* buttons=new QHBoxLayout; auto* next=new QPushButton("&Weitersuchen",dialog); next->setObjectName("access-find-next"); next->setDefault(true);
    auto* close=new QPushButton("&Schließen",dialog); close->setObjectName("access-find-close"); buttons->addWidget(next); buttons->addWidget(close); box->addLayout(buttons);
    next->setEnabled(!initialText.isEmpty()); connect(query,&QLineEdit::textChanged,next,[next](const QString& text) { next->setEnabled(!text.isEmpty()); });
    connect(close,&QPushButton::clicked,dialog,&QDialog::reject); connect(this,&QDialog::finished,dialog,&QDialog::reject);
    const QPointer<AccessFormPreview> self(this); const QPointer<QDialog> guarded(dialog);
    connect(next,&QPushButton::clicked,dialog,[self,guarded,query,scope,match,sensitive,direction,result] {
        if(!self || !guarded || self->closed_ || self->openCancelled_) return;
        try {
            const auto field=scope->currentData().toString(),type=self->dataTypes_.value(field.toCaseFolded());
            std::optional<linaccess::VbaValue> sought;
            if(!field.isEmpty() && !linaccess::nativeFindTextType(type)) sought=linaccess::nativeFindInput(query->text(),type);
            const auto text=sought ? linaccess::nativeFindValueText(*sought) : query->text();
            const auto message=self->findNextRecord(text,field,match->currentIndex(),sensitive->isChecked(),0,false,direction->currentData().toInt(),sought);
            if(guarded) result->setText(message);
        } catch(const std::exception& error) { if(guarded) result->setText(QString::fromUtf8(error.what())); }
    });
    dialog->show(); dialog->activateWindow(); query->setFocus(); query->selectAll();
}
QString AccessFormPreview::findNextRecord(QString text,QString field,int mode,bool caseSensitive,int origin,bool requireValidation,int direction,std::optional<linaccess::VbaValue> sought) {
    using linaccess::VbaRuntimeError;
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    if(closed_ || openCancelled_ || layout_.kind!=linaccess::AccessDocumentKind::Form || !directTable_ || !dataSourceAvailable_ || dataTruncated_)
        throw VbaRuntimeError(3251,"Die Suche ist für diese Datenquelle nicht verfügbar.");
    if(origin<0 || origin>2 || direction<0 || direction>2) throw VbaRuntimeError(5,"Ungültiger Suchbeginn oder Suchrichtung.");
    if(sought) {
        sought=linaccess::normalizeNativeFindValue(*sought);
        if(mode!=1) throw VbaRuntimeError(3251,"Typed FindRecord requires whole-field matching");
    }
    if(findingRecord_) throw VbaRuntimeError(3251,"Eine Suche wird bereits ausgeführt.");
    QScopedValueRollback<bool> searching(findingRecord_,true);
    const QPointer<AccessFormPreview> self(this); auto* search=&findState_;
    if(self->recordCommandRunning_ || self->dataTruncated_) throw VbaRuntimeError(3251,"Die Suche ist während dieses Vorgangs nicht verfügbar.");
    if(text.isEmpty()) throw VbaRuntimeError(2046,"Keine vorherige Suche vorhanden.");
    if(!self->commitControls() || !self || self->closed_ || self->openCancelled_) {
        if(requireValidation) throw VbaRuntimeError(2501,"FindRecord validation cancelled");
        return "Suche durch Formularvalidierung abgebrochen.";
    }
    if(search->text!=text || search->field!=field || search->mode!=mode || search->sensitive!=caseSensitive || search->direction!=direction || !linaccess::sameNativeFindValue(search->value,sought)) {
        search->text=text; search->field=field; search->mode=mode; search->sensitive=caseSensitive; search->direction=direction; search->value=sought; search->advance=false;
    }
    QVector<int> columns;
    for(int column=0;column<self->dataColumns_.size();++column) {
        const auto& name=self->dataColumns_[column]; const auto type=self->dataTypes_.value(name.toCaseFolded());
        if((field.isEmpty() || name.compare(field,Qt::CaseInsensitive)==0)
            && (sought ? linaccess::nativeFindSupports(*sought,type) : linaccess::nativeFindTextType(type))) columns<<column;
    }
    if(columns.isEmpty()) throw VbaRuntimeError(3265,"Das Suchfeld ist nicht mehr verfügbar.");
    const bool virtualRow=self->currentRecordIndex_==0 && self->layout_.allowAdditions;
    const int count=self->recordRows_.size()+(virtualRow ? 1 : 0);
    if(!count) { return "Keine Datensätze vorhanden."; }
    const int current=self->currentRecordIndex_>0 ? self->currentRecordIndex_-1 : virtualRow ? count-1 : 0;
    const int step=direction==0 ? -1 : 1;
    const int start=origin==1 ? 0 : current+((origin==2 || search->advance) ? step : 0);
    int found=-1;
    for(int offset=0;offset<count && found<0;++offset) {
        const qint64 candidateRow=static_cast<qint64>(start)+static_cast<qint64>(step)*offset;
        if(direction!=2 && (candidateRow<0 || candidateRow>=count)) break;
        const int row=static_cast<int>((candidateRow%count+count)%count);
        for(const auto column:columns) {
            QVariant value=row<self->recordRows_.size() ? self->recordRows_[row][column] : QVariant{};
            std::optional<linaccess::VbaValue> editedValue;
            if(row==current && (self->dirty_ || virtualRow)) {
                for(const auto& control:self->layout_.controls) {
                    if(control.type=="ListBox" && control.multiSelect) continue;
                    if(columnIndex(control.controlSource,self->dataColumns_)!=column) continue;
                    const auto pending=self->getFieldValue(control.name.toStdString()); editedValue=pending;
                    value=pending.kind==linaccess::VbaValue::Kind::Null ? QVariant{} : QVariant(QString::fromStdString(pending.asText())); break;
                }
            }
            if(value.isNull() || value.metaType().id()==QMetaType::QByteArray) continue;
            const auto candidate=value.toString(); const auto cs=caseSensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;
            bool matches=false;
            if(sought) {
                const auto typed=editedValue ? *editedValue : linaccess::qtFieldValue(value,self->dataTypes_.value(self->dataColumns_[column].toCaseFolded()));
                matches=typed.kind!=linaccess::VbaValue::Kind::Null && linaccess::nativeFindEqual(linaccess::normalizeNativeFindValue(typed),*sought);
            } else matches=mode==1 ? candidate.compare(text,cs)==0 : mode==2 ? candidate.startsWith(text,cs) : candidate.contains(text,cs);
            if(matches) { found=row; break; }
        }
    }
    if(found<0) { return "Keine Übereinstimmung gefunden."; }
    if(found!=current) {
        const auto id=self->rowIds_[found];
        if(!self->saveRecord() || !self || self->closed_ || self->openCancelled_) {
            if(requireValidation) throw VbaRuntimeError(2501,"FindRecord record save cancelled");
            return "Suche durch Formularvalidierung abgebrochen.";
        }
        const auto position=self->rowIds_.indexOf(id);
        if(position<0) {
            if(requireValidation) throw VbaRuntimeError(2501,"FindRecord target no longer available");
            return "Gefundener Datensatz ist nicht mehr verfügbar.";
        }
        self->selectRecord(position+1);
        if(!self || self->closed_ || self->currentRecordIndex_<1 || self->rowIds_.value(self->currentRecordIndex_-1)!=id) {
            if(requireValidation) throw VbaRuntimeError(2501,"FindRecord navigation cancelled by form event");
            return "Suche durch Formularereignis abgebrochen.";
        }
    }
    search->advance=true;
    return self->currentRecordIndex_==0 ? "Aktueller neuer Datensatz gefunden." : QString("Datensatz %1 gefunden.").arg(self->currentRecordIndex_);
}
void AccessFormPreview::repeatFind() {
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    if(findState_.text.isEmpty()) throw linaccess::VbaRuntimeError(2046,"Keine vorherige Suche vorhanden.");
    const auto message=findNextRecord(findState_.text,findState_.field,findState_.mode,findState_.sensitive,0,false,findState_.direction,findState_.value);
    if(status_) status_->setText(message);
    if(findDialog_) if(auto* result=findDialog_->findChild<QLabel*>("access-find-result")) result->setText(message);
}
bool AccessFormPreview::zoomFocusedText() {
    ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
    QPointer<QLineEdit> line=qobject_cast<QLineEdit*>(QApplication::focusWidget());
    const linaccess::AccessVisualControl* selected=nullptr;
    for(const auto& control:layout_.controls)
        if(line && controls_.value(control.name.toCaseFolded())==line) { selected=&control; break; }
    if(!selected || selected->type!="TextBox" || !line->isEnabled()) return false;
    if(auto* modal=QApplication::activeModalWidget(); modal && modal!=window() && !modal->isAncestorOf(line)) return false;
    const auto name=selected->name,binding=selected->onChange;
    const auto original=line->text(); const bool readOnly=line->isReadOnly();
    const auto previousZoom=line->property("la_native_zoom_open"); line->setProperty("la_native_zoom_open",true);
    const auto zoomScope=qScopeGuard([line,previousZoom] { if(line) line->setProperty("la_native_zoom_open",previousZoom); });
    const auto result=linaccess::nativeCellZoom(line,original,readOnly);
    if(!result || !line || closed_ || openCancelled_ || readOnly || line->isReadOnly() || !line->isEnabled()) return true;
    if(result->text==original && !result->modified) return true;
    ensureWritableControl(name.toStdString());
    if(result->text.size()>line->maxLength()) throw linaccess::VbaRuntimeError(3251,"Zoom text exceeds the native editor capacity");
    rememberControlEdit(line);
    if(!editOriginals_.contains(name)) throw linaccess::VbaRuntimeError(3251,"Zoom text cannot be buffered safely");
    clearDefaultError(line); bufferedTextControls_.insert(name); pendingControls_.insert(name);
    {
        QScopedValueRollback<bool> loading(loading_,true);
        setText(line,result->text,false); line->setProperty("la_null",false);
        line->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text));
    }
    markDirty(name,true);
    if(closed_ || openCancelled_ || line->property("la_insert_rejected").toBool()) return true;
    refreshCalculatedControls();
    if(!binding.isEmpty()) (void)runControlBinding(name,"Change",binding);
    return true;
}
bool AccessFormPreview::handleDocumentKey(QKeyEvent* event) {
    if (closed_ || openCancelled_) return false;
    const auto modifiers=event->modifiers() & ~Qt::KeypadModifier;
    try {
        if(event->key()==Qt::Key_F6 && (modifiers==Qt::ControlModifier || modifiers==(Qt::ControlModifier|Qt::ShiftModifier))) {
            QWidget* owner=nullptr;
            for(auto* ancestor=parentWidget();ancestor;ancestor=ancestor->parentWidget())
                if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
            if(!owner) return false;
            cycleProjectDocuments(db_,owner,modifiers.testFlag(Qt::ShiftModifier),dynamic_cast<AccessFormPreview*>(window()));
            return true;
        }
        const char* projectCommand=nullptr;
        if(event->key()==Qt::Key_F11 && modifiers==Qt::NoModifier) projectCommand="focusDatabaseWindow";
        else if(event->key()==Qt::Key_O && modifiers==Qt::ControlModifier) projectCommand="chooseDatabase";
        else if(event->key()==Qt::Key_N && modifiers==Qt::ControlModifier) projectCommand="newDatabase";
        if(projectCommand) {
            if(isModal() || QApplication::activeModalWidget()) return true;
            QWidget* owner=nullptr;
            for(auto* ancestor=parentWidget();ancestor;ancestor=ancestor->parentWidget())
                if(qobject_cast<QMainWindow*>(ancestor)) { owner=ancestor; break; }
            if(!owner || owner->metaObject()->indexOfMethod((QByteArray(projectCommand)+"()").constData())<0) return false;
            if(event->key()!=Qt::Key_F11 && nativeExecutionDepth()!=0)
                throw linaccess::VbaRuntimeError(3251,"Project switching from executing native code is not implemented");
            ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
            if(!commitControls() || closed_ || openCancelled_) return true;
            (void)QMetaObject::invokeMethod(owner,projectCommand,Qt::DirectConnection);
            return true;
        }
        if ((event->key()==Qt::Key_PageDown || event->key()==Qt::Key_PageUp) && modifiers==Qt::ControlModifier) {
            (void)dispatchPreviewDocumentCommand(*this,"runcommand",{linaccess::VbaValue::integer(event->key()==Qt::Key_PageDown ? 65 : 66)});
            return true;
        }
        if (event->key()==Qt::Key_F4 && modifiers==Qt::ShiftModifier) { repeatFind(); return true; }
        if (event->key()==Qt::Key_F && modifiers==Qt::ControlModifier) { openFindDialog(); return true; }
        if (event->key()==Qt::Key_C && modifiers==Qt::ControlModifier && recordSelected_) {
            (void)dispatchPreviewDocumentCommand(*this,"runcommand",{linaccess::VbaValue::numeric(190)}); return true;
        }
        if ((event->key()==Qt::Key_Plus && (modifiers==Qt::ControlModifier || modifiers==(Qt::ControlModifier|Qt::ShiftModifier)))
            || (event->key()==Qt::Key_Equal && modifiers==(Qt::ControlModifier|Qt::ShiftModifier))) {
            (void)dispatchPreviewDocumentCommand(*this,"runcommand",{linaccess::VbaValue::numeric(28)}); return true;
        }
        if (event->key()==Qt::Key_Minus && modifiers==Qt::ControlModifier) {
            (void)dispatchPreviewDocumentCommand(*this,"runcommand",{linaccess::VbaValue::numeric(223)}); return true;
        }
        if(event->key()==Qt::Key_F2 && modifiers==Qt::ShiftModifier) return zoomFocusedText();
        if(const auto clockText=linaccess::nativeClockKeyText(*event)) {
            ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
            auto* line=qobject_cast<QLineEdit*>(QApplication::focusWidget());
            const linaccess::AccessVisualControl* selected=nullptr;
            for(const auto& control:layout_.controls)
                if(line && controls_.value(control.name.toCaseFolded())==line) { selected=&control; break; }
            if(!selected || selected->type!="TextBox" || layout_.kind!=linaccess::AccessDocumentKind::Form
                || line->isReadOnly() || !line->isEnabled()) throw linaccess::VbaRuntimeError(2046,"Clock insertion is unavailable");
            if(auto* modal=QApplication::activeModalWidget(); modal && modal!=window() && !modal->isAncestorOf(line)) return true;
            ensureWritableControl(selected->name.toStdString());
            const auto& text=*clockText;
            if(line->text().size()-line->selectedText().size()+text.size()>line->maxLength())
                throw linaccess::VbaRuntimeError(3251,"Clock text exceeds the native editor capacity");
            rememberControlEdit(line);
            if(!editOriginals_.contains(selected->name)) throw linaccess::VbaRuntimeError(3251,"Clock insertion cannot be buffered safely");
            bufferedTextControls_.insert(selected->name); pendingControls_.insert(selected->name);
            line->insert(text);
            return true;
        }
        if (event->key()==Qt::Key_Apostrophe && modifiers==Qt::ControlModifier) {
            ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
            auto* line=qobject_cast<QLineEdit*>(QApplication::focusWidget());
            const linaccess::AccessVisualControl* selected=nullptr;
            for(const auto& control:layout_.controls)
                if(line && controls_.value(control.name.toCaseFolded())==line) { selected=&control; break; }
            if(!selected || selected->type!="TextBox" || layout_.kind!=linaccess::AccessDocumentKind::Form
                || !directTable_ || !dataSourceAvailable_ || line->isReadOnly() || !line->isEnabled())
                throw linaccess::VbaRuntimeError(2046,"Copying the previous field is unavailable");
            if(auto* modal=QApplication::activeModalWidget(); modal && modal!=window() && !modal->isAncestorOf(line)) return true;
            ensureWritableControl(selected->name.toStdString());
            const auto column=columnIndex(selected->controlSource,dataColumns_);
            const auto row=currentRecordIndex_>0 ? currentRecordIndex_-2 : recordRows_.size()-1;
            if(column<0 || row<0 || row>=recordRows_.size()) throw linaccess::VbaRuntimeError(2046,"No previous bound field is available");
            const auto value=recordRows_[row][column];
            if(value.metaType().id()==QMetaType::QByteArray) throw linaccess::VbaRuntimeError(3251,"Copying binary fields is not implemented");
            const auto text=linaccess::qtFieldText(value,dataTypes_.value(dataColumns_[column].toCaseFolded())); const bool null=value.isNull();
            if(text.size()>line->maxLength()) throw linaccess::VbaRuntimeError(3251,"Previous field exceeds the native editor capacity");
            if(line->text()==text && line->property("la_null").toBool()==null) return true;
            const auto name=selected->name,binding=selected->onChange;
            rememberControlEdit(line);
            if(!editOriginals_.contains(name)) throw linaccess::VbaRuntimeError(3251,"Current field value cannot be buffered safely");
            clearDefaultError(line);
            bufferedTextControls_.insert(name); pendingControls_.insert(name);
            {
                QScopedValueRollback<bool> loading(loading_,true);
                setText(line,text,null); line->setProperty("la_null",null);
                line->setProperty("la_vba_kind",static_cast<int>(null ? linaccess::VbaValue::Kind::Null : linaccess::VbaValue::Kind::Text));
            }
            markDirty(name,true);
            if(closed_ || openCancelled_ || line->property("la_insert_rejected").toBool()) return true;
            refreshCalculatedControls();
            if(!binding.isEmpty()) (void)runControlBinding(name,"Change",binding);
            return true;
        }
        if (event->key()==Qt::Key_Delete && modifiers==Qt::NoModifier && recordSelected_) {
            (void)dispatchPreviewDocumentCommand(*this,"runcommand",{linaccess::VbaValue::numeric(223)}); return true;
        }
        if (event->key()==Qt::Key_Escape && modifiers==Qt::NoModifier) {
            if (layout_.kind==linaccess::AccessDocumentKind::Report) { (void)close(); return true; }
            for (const auto& definition:layout_.controls) {
                auto* button=qobject_cast<QPushButton*>(controls_.value(definition.name.toCaseFolded()).data());
                if(definition.type!="CommandButton" || !button || !button->property("la_access_cancel").toBool() || !button->isVisible() || !button->isEnabled()) continue;
                const QPointer<QPushButton> guarded=button;
                auto* focus=QApplication::focusWidget();
                for(const auto& control:layout_.controls)
                    if(control.type=="TextBox" && controls_.value(control.name.toCaseFolded()).data()==focus) { (void)undoControlEdit(control.name); break; }
                if(guarded && !closed_ && !openCancelled_) guarded->click();
                return true;
            }
            undoKeyboardEdit();
            return true;
        }
        if ((event->key()==Qt::Key_Return || event->key()==Qt::Key_Enter) && modifiers==Qt::ShiftModifier) {
            if (!saveRecord()) status_->setText("Speichern abgebrochen"); return true;
        }
        if (event->key()==Qt::Key_F9 && modifiers==Qt::ShiftModifier) {
            (void)makePreviewFormObject(*this)->member("Requery",{}); return true;
        }
        if (event->key()==Qt::Key_F9 && modifiers==Qt::NoModifier) {
            ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
            const auto* focus=QApplication::focusWidget();
            for(const auto& control:layout_.controls) {
                auto* widget=controls_.value(control.name.toCaseFolded()).data();
                if(auto* list=dynamic_cast<AccessListBox*>(widget); list && (focus==list || focus==list->viewport())) { requeryControl(control.name.toUtf8().toStdString()); break; }
                auto* combo=qobject_cast<QComboBox*>(widget);
                if(combo && (focus==combo || focus==combo->lineEdit())) { requeryControl(control.name.toUtf8().toStdString()); break; }
            }
            if(!closed_ && !openCancelled_) refreshCalculatedControls();
            return true;
        }
        if ((event->key()==Qt::Key_W || event->key()==Qt::Key_F4) && modifiers==Qt::ControlModifier) {
            (void)window()->close(); return true;
        }
        if (event->key()==Qt::Key_F2 && modifiers==Qt::NoModifier) {
            auto* line=qobject_cast<QLineEdit*>(QApplication::focusWidget());
            if (line && isAncestorOf(line)) {
                if (line->hasSelectedText()) line->setCursorPosition(line->text().size()); else line->selectAll();
                return true;
            }
        }
        if ((event->key()==Qt::Key_Return || event->key()==Qt::Key_Enter) && modifiers==Qt::NoModifier) {
            auto* focus=QApplication::focusWidget();
            if (auto* button=qobject_cast<QPushButton*>(focus); button && isAncestorOf(button)) { button->click(); return true; }
            if (!commitControls()) return true;
            for (const auto& control:layout_.controls) {
                auto* button=qobject_cast<QPushButton*>(controls_.value(control.name.toCaseFolded()).data());
                if (control.type=="CommandButton" && button && button->isDefault() && button->isEnabled() && button->isVisible()) {
                    button->click(); return true;
                }
            }
            (void)focusNextPrevChild(true); return true;
        }
    } catch (const std::exception& error) { if (status_) status_->setText(QString::fromUtf8(error.what())); return true; }
    return false;
}
bool AccessFormPreview::ownsKeyboardFocus() const {
    const auto* focus=QApplication::focusWidget();
    for (auto* ancestor=focus;ancestor;ancestor=ancestor->parentWidget())
        if (auto* document=dynamic_cast<const AccessFormPreview*>(ancestor)) {
            if (document!=this) return true; // Embedded documents dispatch their own events.
            break;
        }
    for (const auto& widget : controls_) {
        const auto* combo=qobject_cast<QComboBox*>(widget.data());
        if (widget && (widget==focus || (combo && combo->lineEdit()==focus))) return true;
    }
    return false;
}
bool AccessFormPreview::dispatchKeyboardEvent(QObject* watched,QKeyEvent* event) {
    if (deliveringKeyReplacement_ || !opened_ || closed_ || layout_.kind!=linaccess::AccessDocumentKind::Form) return false;
    if (event->type()==QEvent::KeyRelease && event->isAutoRepeat()) return false;
    const linaccess::AccessVisualControl* control=nullptr;
    for (const auto& definition : layout_.controls) {
        const auto widget=controls_.value(definition.name.toCaseFolded());
        const auto* combo=qobject_cast<QComboBox*>(widget.data());
        if (widget && (widget==watched || (combo && combo->lineEdit()==watched))) { control=&definition; break; }
    }
    const auto virtualKey=[&]() -> std::optional<int> {
        const auto key=event->key();
        if ((key>=Qt::Key_A && key<=Qt::Key_Z) || (key>=Qt::Key_0 && key<=Qt::Key_9) || key==Qt::Key_Space) return key;
        if (key>=Qt::Key_F1 && key<=Qt::Key_F24) return 112+key-Qt::Key_F1;
        static const QMap<int,int> codes{{Qt::Key_Backspace,8},{Qt::Key_Tab,9},{Qt::Key_Backtab,9},{Qt::Key_Return,13},{Qt::Key_Enter,13},
            {Qt::Key_Shift,16},{Qt::Key_Control,17},{Qt::Key_Alt,18},{Qt::Key_Pause,19},{Qt::Key_CapsLock,20},{Qt::Key_Escape,27},
            {Qt::Key_PageUp,33},{Qt::Key_PageDown,34},{Qt::Key_End,35},{Qt::Key_Home,36},{Qt::Key_Left,37},{Qt::Key_Up,38},
            {Qt::Key_Right,39},{Qt::Key_Down,40},{Qt::Key_Insert,45},{Qt::Key_Delete,46},{Qt::Key_NumLock,144},{Qt::Key_ScrollLock,145},
            {Qt::Key_Semicolon,186},{Qt::Key_Colon,186},{Qt::Key_Equal,187},{Qt::Key_Plus,187},{Qt::Key_Comma,188},
            {Qt::Key_Minus,189},{Qt::Key_Period,190},{Qt::Key_Slash,191},{Qt::Key_Question,191},{Qt::Key_QuoteLeft,192},
            {Qt::Key_BracketLeft,219},{Qt::Key_Backslash,220},{Qt::Key_BracketRight,221},{Qt::Key_Apostrophe,222}};
        const auto found=codes.constFind(key);
        return found==codes.cend() ? std::nullopt : std::optional<int>(*found);
    };
    const int shift=(event->modifiers().testFlag(Qt::ShiftModifier) ? 1 : 0)
        | (event->modifiers().testFlag(Qt::ControlModifier) ? 2 : 0) | (event->modifiers().testFlag(Qt::AltModifier) ? 4 : 0);
    const auto dispatch=[&](const QString& suffix,int code,bool character) {
        std::vector<linaccess::VbaValue> arguments{linaccess::VbaValue::numeric(code)};
        if (!character) arguments.push_back(linaccess::VbaValue::numeric(shift));
        const auto invoke=[&](bool form) {
            const auto binding=form ? layout_.events.value(("on"+suffix).toCaseFolded()) : !control ? QString{}
                : suffix=="KeyDown" ? control->onKeyDown : suffix=="KeyUp" ? control->onKeyUp : control->onKeyPress;
            if (binding.isEmpty()) return;
            if (!eventProcedure(binding)) throw linaccess::VbaRuntimeError(3251,"Keyboard macro/function bindings unsupported");
            const auto handler=form ? "Form_"+suffix : controlEventPrefix(control->name)+"_"+suffix;
            const auto result=executeProcedure(handler,arguments);
            eventResults_.insert(form ? ("On"+suffix).toCaseFolded() : ("control:"+control->name+":"+suffix).toCaseFolded(),result);
            if (failed(result)) throw linaccess::VbaRuntimeError(result.errorNumber ? result.errorNumber : 3251,"Keyboard handler failed");
        };
        if (layout_.keyPreview) { invoke(true); if (!closed_ && arguments.front().asNumber()!=0) invoke(false); }
        else { invoke(false); if (!closed_ && arguments.front().asNumber()!=0) invoke(true); }
        const auto value=arguments.front().asNumber();
        if (!std::isfinite(value) || value!=std::floor(value) || value<0 || value>255)
            throw linaccess::VbaRuntimeError(5,"Invalid keyboard event result");
        return static_cast<int>(value);
    };
    try {
        const auto key=virtualKey();
        const bool release=event->type()==QEvent::KeyRelease;
        const bool physicalBinding=!layout_.events.value(release ? "onkeyup" : "onkeydown").isEmpty()
            || (control && !(release ? control->onKeyUp : control->onKeyDown).isEmpty());
        if (key) {
            const int result=dispatch(release ? "KeyUp" : "KeyDown",*key,false);
            if (closed_ || result==0) return true;
            if (result!=*key) throw linaccess::VbaRuntimeError(3251,"Virtual key remapping unsupported");
        } else if (physicalBinding) throw linaccess::VbaRuntimeError(3251,"Native virtual-key mapping unsupported");
        if (event->type()==QEvent::KeyRelease) return false;
        if (event->text().isEmpty() || (shift&6)) return false;
        const bool keyPressBound=!layout_.events.value("onkeypress").isEmpty() || (control && !control->onKeyPress.isEmpty());
        if (!keyPressBound) return false;
        if (event->text().size()!=1) throw linaccess::VbaRuntimeError(3251,"Multi-unit KeyPress locale mapping unsupported");
        const auto unicode=event->text().front().unicode();
        if (unicode>127) throw linaccess::VbaRuntimeError(3251,"Non-ASCII KeyPress locale mapping unsupported");
        const auto character=dispatch("KeyPress",unicode,true);
        if (closed_ || character==0) return true;
        if (character!=unicode) {
            if (character>127) throw linaccess::VbaRuntimeError(3251,"Non-ASCII KeyPress replacement unsupported");
            auto* receiver=qobject_cast<QWidget*>(watched);
            if (!receiver) return true;
            const QString text{QChar(character)};
            const auto qtKey=character==13 ? Qt::Key_Return : character==9 ? Qt::Key_Tab : character==8 ? Qt::Key_Backspace : text.toUpper().front().unicode();
            QKeyEvent replacement(QEvent::KeyPress,qtKey,event->modifiers(),text,event->isAutoRepeat(),event->count());
            QScopedValueRollback<bool> guard(deliveringKeyReplacement_,true); QApplication::sendEvent(receiver,&replacement); return true;
        }
    } catch (const std::exception& error) {
        const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error);
        linaccess::VbaExecutionResult result; result.errorNumber=typed ? typed->number : 5;
        result.status=result.errorNumber==3251 ? linaccess::VbaExecutionStatus::Unsupported : linaccess::VbaExecutionStatus::Error;
        result.explanation=error.what(); eventResults_.insert("keyboarderror",result);
        if (status_) status_->setText(QString::fromUtf8(error.what())); return true;
    }
    return false;
}
void AccessFormPreview::keyPressEvent(QKeyEvent* event) {
    if ((!ownsKeyboardFocus() && dispatchKeyboardEvent(this,event)) || handleDocumentKey(event)) event->accept();
    else QDialog::keyPressEvent(event);
}
void AccessFormPreview::keyReleaseEvent(QKeyEvent* event) {
    if (!ownsKeyboardFocus() && dispatchKeyboardEvent(this,event)) event->accept(); else QDialog::keyReleaseEvent(event);
}
bool AccessFormPreview::eventFilter(QObject* watched,QEvent* event) {
    if (auto* document=dynamic_cast<AccessFormPreview*>(watched); document && document!=this
        && (event->type()==QEvent::KeyPress || event->type()==QEvent::KeyRelease)) return QDialog::eventFilter(watched,event);
    if (event->type()==QEvent::Shortcut) {
        if(auto* label=qobject_cast<QLabel*>(watched);label && label->buddy()) {
            ++nativeExecutionDepth(); const auto execution=qScopeGuard([] { --nativeExecutionDepth(); disposeClosedForms(); });
            try {
                if(!commitControls() || closed_ || openCancelled_) { event->accept(); return true; }
            } catch(const std::exception& error) {
                if(status_) status_->setText(QString::fromUtf8(error.what()));
                event->accept(); return true;
            }
        }
    }
    if (event->type()==QEvent::FocusIn) {
        if (watched!=recordSelector_) clearRecordSelection();
        rememberControlEdit(watched);
    }
    if ((event->type()==QEvent::KeyPress || event->type()==QEvent::KeyRelease)
        && dispatchKeyboardEvent(watched,static_cast<QKeyEvent*>(event))) { event->accept(); return true; }
    if (event->type()==QEvent::KeyPress && handleDocumentKey(static_cast<QKeyEvent*>(event))) { event->accept(); return true; }
    if (event->type()==QEvent::KeyPress) {
        const auto* key=static_cast<QKeyEvent*>(event); const auto modifiers=key->modifiers() & ~Qt::KeypadModifier;
        if ((key->key()==Qt::Key_Tab || key->key()==Qt::Key_Backtab) && (modifiers==Qt::NoModifier || modifiers==Qt::ShiftModifier)
            && advanceTabFocus(key->key()==Qt::Key_Tab && modifiers==Qt::NoModifier)) { event->accept(); return true; }
    }

    if (event->type()==QEvent::MouseButtonDblClick) {
        const auto binding=watched->property("la_access_dblclick").toString();
        if (!binding.isEmpty()) {
            const auto result=runControlBinding(watched->property("la_access_control_name").toString(),"DblClick",binding);
            if (result.cancelled || failed(result)) return true;
        }
    }
    return QDialog::eventFilter(watched,event);
}
linaccess::VbaExecutionResult AccessFormPreview::runControlBinding(const QString& control,const QString& event,const QString& binding) {
    if (binding.trimmed().isEmpty()) return {linaccess::VbaExecutionStatus::Completed,{}};
    linaccess::VbaExecutionResult result;
    if (eventProcedure(binding)) result=runHandler(controlEventPrefix(control)+"_"+event);
    else {
        PreviewVbaHost host(*this);
        result=linaccess::executeVbaEventExpression(vba_.toUtf8().toStdString(),binding.toUtf8().toStdString(),host);
        if (failed(result) && status_) status_->setText("Ereignisfehler "+QString::number(result.errorNumber)+" (Zeile "+QString::number(result.line)+")");
    }
    eventResults_.insert(("control:"+control+":"+event).toCaseFolded(),result); return result;
}

linaccess::VbaExecutionResult AccessFormPreview::executeProcedure(const QString& name,std::vector<linaccess::VbaValue>& arguments) {
    PreviewVbaHost host(*this);
    return linaccess::executeVbaProcedureWithArguments(vba_.toUtf8().toStdString(),name.toUtf8().toStdString(),host,arguments);
}
linaccess::VbaValue AccessFormPreview::evalExpression(std::string_view expression) {
    PreviewVbaHost host(*this); return host.eval(expression);
}
linaccess::VbaExecutionResult AccessFormPreview::runHandler(const QString& handler) {
    if (vba_.isEmpty()) return {};
    std::vector<linaccess::VbaValue> arguments;
    const auto result = executeProcedure(handler,arguments);
    if (result.status == linaccess::VbaExecutionStatus::Unsupported
        || result.status == linaccess::VbaExecutionStatus::Error) {
        const auto message = handler + " (Zeile " + QString::number(result.line)
            + "): " + QString::fromStdString(result.explanation);
        if (status_) status_->setText(message);
        else QMessageBox::critical(this, "VBA", message);
    }
    return result;
}
linaccess::VbaExecutionResult AccessFormPreview::formEvent(const QString& property,const QString& handler) {
    if (!eventProcedure(layout_.events.value(property.toCaseFolded()))) return {};

    auto name = handler;
    if (layout_.kind==linaccess::AccessDocumentKind::Report && name.startsWith("Form_")) name.replace(0,5,"Report_");
    const auto result = runHandler(name);
    eventResults_.insert(property.toCaseFolded(),result);
    return result;
}
