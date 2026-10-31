#include "access_form_preview.hpp"
#include "access_form_objects.hpp"

#include "vba_runtime.hpp"
#include "vba_currency.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "access_format.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"

#include <QApplication>
#include <QCloseEvent>
#include <QKeyEvent>
#include <QTimer>
#include <QScopedValueRollback>
#include <QSignalBlocker>
#include <QCheckBox>
#include <QComboBox>
#include <QFormLayout>
#include <QHBoxLayout>
#include <QLabel>
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
#include <stdexcept>

namespace {
QPointer<AccessFormPreview>& activeDocument() {
    static QPointer<AccessFormPreview> document;
    return document;
}
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
    if (auto* line = qobject_cast<QLineEdit*>(editor)) return line->text();
    if (auto* checkbox = qobject_cast<QCheckBox*>(editor)) return checkbox->isChecked() ? "-1" : "0";
    if (auto* combo = qobject_cast<QComboBox*>(editor)) {
        if (combo->property("la_rows_loaded").toBool() && combo->currentIndex()>=0 && combo->currentText()==combo->itemText(combo->currentIndex()))
            return combo->currentData().toString();
        return combo->currentText();
    }
    return {};
}
void setText(QWidget* editor, const QString& text) {
    if (auto* line = qobject_cast<QLineEdit*>(editor)) line->setText(text);
    else if (auto* checkbox = qobject_cast<QCheckBox*>(editor))
        checkbox->setChecked(text == "-1" || text == "1" || text.compare("true", Qt::CaseInsensitive) == 0);
    else if (auto* combo = qobject_cast<QComboBox*>(editor)) {
        if (combo->property("la_rows_loaded").toBool()) {
            for (int i=0;i<combo->count();++i) if (combo->itemData(i).toString()==text) { combo->setCurrentIndex(i); return; }
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

class PreviewVbaHost final : public linaccess::VbaHost {
public:
    explicit PreviewVbaHost(AccessFormPreview& preview) : preview_(preview) {}
    std::optional<std::string> field(std::string_view name) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (key == "dirty") return preview_.dirty_ ? "-1" : "0";
        if (key == "newrecord") return preview_.currentRecordIndex_ == 0 ? "-1" : "0";
        return preview_.getField(name);
    }
    linaccess::VbaValue controlValue(std::string_view name) override { return preview_.getFieldValue(name); }
    std::optional<linaccess::VbaValue> resolveIdentifier(std::string_view name,bool reference) override {
        const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (!preview_.controls_.contains(key)) return std::nullopt;
        return reference ? makePreviewFormObject(preview_)->objectItem(name) : preview_.getFieldValue(name);
    }
    bool assignIdentifier(std::string_view name,const linaccess::VbaValue& value) override {
        const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (!preview_.controls_.contains(key)) return false;
        preview_.setFieldValue(name,value); return true;
    }
    linaccess::VbaValue expressionValue(std::string_view name,bool reference) override {
        try { return reference ? makePreviewFormObject(preview_)->objectItem(name) : preview_.getFieldValue(name); }
        catch(const linaccess::VbaRuntimeError& failure) { return linaccess::VbaValue::errorValue(failure.number); }
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
    void setControlValue(std::string_view name,const linaccess::VbaValue& value) override { preview_.setFieldValue(name,value); }
    linaccess::VbaValue fieldValue(std::string_view name) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (preview_.layout_.kind==linaccess::AccessDocumentKind::Report && (key=="dirty" || key=="newrecord"))
            throw linaccess::VbaRuntimeError(438,"Report member unsupported");
        if (key == "name") return linaccess::VbaValue::string(preview_.formName_.toUtf8().toStdString());
        if (key == "caption") return makePreviewFormObject(preview_)->member("Caption",{});
        if (key == "tag") return makePreviewFormObject(preview_)->member("Tag",{});
        if (key=="filter" || key=="filteron") return makePreviewFormObject(preview_)->member(key.toStdString(),{});
        if (key == "recordsource") return linaccess::VbaValue::string(preview_.layout_.recordSource.toUtf8().toStdString());
        if (key == "dirty") return linaccess::VbaValue::logical(preview_.dirty_);
        if (key == "newrecord") return linaccess::VbaValue::logical(preview_.currentRecordIndex_ == 0);
        return preview_.getFieldValue(name);
    }
    void setFieldValue(std::string_view name,const linaccess::VbaValue& value) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (preview_.layout_.kind==linaccess::AccessDocumentKind::Report && key=="dirty")
            throw linaccess::VbaRuntimeError(438,"Report member unsupported");
        if (key == "caption" || key=="tag" || key=="filter" || key=="filteron") { makePreviewFormObject(preview_)->assign(key.toUtf8().toStdString(),value); return; }
        if (key == "name" || key == "recordsource" || key == "newrecord")
            throw linaccess::VbaRuntimeError(438,"Form property is read-only");
        if (key == "dirty") { setField(name,value.asField()); return; }
        preview_.setFieldValue(name,value);
    }
    void setField(std::string_view name, std::optional<std::string> value) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
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
    void saveRecord() override {
        if (!preview_.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Datensatz konnte nicht gespeichert werden");
    }
    void quitApplication() override { QApplication::quit(); }
    void maximizeForm() override { preview_.window()->showMaximized(); }
    void gotoLastRecord() override {
        if (preview_.recordSelector_->count() > 0)
            preview_.recordSelector_->setCurrentIndex(preview_.recordSelector_->count() - 1);
    }
    void gotoNextRecord() override {
        preview_.recordSelector_->setCurrentIndex(qMin(
            preview_.recordSelector_->currentIndex() + 1,
            preview_.recordSelector_->count() - 1));
    }
    void gotoNewRecord() override {
        if (preview_.layout_.kind==linaccess::AccessDocumentKind::Report) throw linaccess::VbaRuntimeError(438,"Report does not support new records");
        preview_.recordSelector_->setCurrentIndex(0);
    }
    void requery() override {
        (void)makePreviewFormObject(preview_)->member("Requery",{});
    }
    void requeryControl(std::string_view name) override { preview_.requeryControl(name); }

    void focusControl(std::string_view name) override {
        auto* widget = preview_.controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded()).data();
        if (!widget) throw std::runtime_error("Steuerelement nicht gefunden: " + std::string(name));
        widget->setFocus();
    }
    void setControlVisible(std::string_view name, bool value) override {
        auto* widget = preview_.controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded()).data();
        if (!widget) throw std::runtime_error("Steuerelement nicht gefunden: " + std::string(name));
        widget->setVisible(value);
    }
    std::optional<std::string> moduleForProcedure(std::string_view procedure) override {
        QSqlQuery query(preview_.db_);
        if (!query.exec("SELECT source_bytes FROM _la_access_artifacts "
                        "WHERE kind='vba' ORDER BY id")) return std::nullopt;
        const auto target = QString::fromUtf8(procedure.data(), procedure.size());
        while (query.next()) {
            const auto code = linaccess::decodeVbaSource(query.value(0).toByteArray()).toUtf8().toStdString();
            const auto analysis = linaccess::analyzeVba(code);
            for (const auto& item : analysis.procedures)
                if (QString::fromStdString(item.name).compare(target, Qt::CaseInsensitive) == 0)
                    return code;
        }
        return std::nullopt;
    }
    std::shared_ptr<linaccess::VbaObject> applicationObject() override { return makePreviewApplicationObject(preview_); }
    std::shared_ptr<linaccess::VbaObject> currentForm() override { return makePreviewFormObject(preview_); }
    std::shared_ptr<linaccess::VbaObject> reportsCollection() override { return makePreviewReportsObject(preview_); }
    std::shared_ptr<linaccess::VbaObject> formsCollection() override { return makePreviewFormsObject(preview_); }
    linaccess::VbaValue formatValue(const linaccess::VbaValue& value,std::string_view mask,bool stringResult) override {
        return linaccess::formatAccessValue(value,mask,stringResult);
    }
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(preview_.db_); }
    void executeSql(std::string_view statement) override {
        if (!preview_.db_.isOpen()) throw std::runtime_error("Kein offenes Datenbankprojekt");
        auto database = linaccess::makeDaoDatabase(preview_.db_);
        (void)database->member("Execute",{linaccess::VbaValue::string(std::string(statement))});
    }
    void openForm(std::string_view name) override {
        preview_.openOtherForm(QString::fromUtf8(name.data(), name.size()));
    }
private:
    AccessFormPreview& preview_;
};

AccessFormPreview::AccessFormPreview(QString formName, const QByteArray& source,
                                     const QByteArray& vbaSource, QSqlDatabase database,
                                     QWidget* parent,bool embedded)
    : QDialog(parent), formName_(std::move(formName)),
      vba_(linaccess::decodeVbaSource(vbaSource)),
      layout_(linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(source))),
      db_(std::move(database)) {
    static thread_local QStringList constructionStack;
    const auto identity=db_.connectionName()+QChar(0x1f)+formName_.toCaseFolded();
    if (constructionStack.contains(identity) || constructionStack.size()>=16)
        throw linaccess::VbaRuntimeError(2455,"Recursive subform source");
    QScopedValueRollback<QStringList> constructionGuard(constructionStack);
    constructionStack.append(identity);
    if (embedded) setWindowFlags(Qt::Widget);
    filter_=layout_.filter; filterOn_=layout_.filterOnLoad && !filter_.trimmed().isEmpty();
    initialCurrentDeferred_=!embedded && eventProcedure(layout_.events.value("onactivate"));
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
    canvas->setStyleSheet("#access-form-canvas { background: #f0f0f0; }");
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
    if (!queryViewsReady) {
        status_->setText("Abfrageansichten werden noch verwendet; andere Ansichten schließen und Formular erneut öffnen.");
        canvas->setEnabled(false); save->setEnabled(false); return;
    }
    const bool report = layout_.kind==linaccess::AccessDocumentKind::Report;
    if (!report) loadRecords();
    const auto outcome = formEvent("OnOpen","Form_Open");
    if (outcome.cancelled) { openCancelled_ = true; return; }
    if (report) {
        loadRecords();
        if (dataSourceAvailable_ && recordRows_.isEmpty()) {
            const auto noData = formEvent("OnNoData","Report_NoData");
            if (noData.cancelled) { openCancelled_ = true; return; }
        }
    }
    (void)formEvent("OnLoad","Form_Load");
    opened_ = true;
    if (nativeActivationPending_) activateDocument();
    if (!initialCurrentDeferred_) initialCurrent();
}

void AccessFormPreview::initialCurrent() {
    if (initialCurrentDispatched_ || closed_ || openCancelled_) return;
    initialCurrentDispatched_=true;
    (void)formEvent("OnCurrent","Form_Current");
}

bool AccessFormPreview::event(QEvent* event) {
    const auto type=event->type();
    const auto result=QDialog::event(event);
    if (type==QEvent::WindowActivate && isWindow()) {
        nativeActivationPending_=true;
        if (opened_) activateDocument();
    }
    return result;
}

void AccessFormPreview::activateDocument() {
    if (!opened_ || closed_ || openCancelled_ || !isWindow()) return;
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

void AccessFormPreview::buildControls(QWidget* canvas) {
    for (const auto& control : layout_.controls) {

        QWidget* widget = nullptr;
        if (control.type == "Label") {
            auto* label = new QLabel(control.caption, canvas);
            label->setWordWrap(true);
            widget = label;
        } else if (control.type == "TextBox") {
            auto* edit = new QLineEdit(canvas);
            edit->setReadOnly(control.controlSource.startsWith('='));
            widget = edit;
        } else if (control.type == "ComboBox" || control.type == "ListBox") {
            auto* combo = new QComboBox(canvas);
            combo->setEditable(true);
            widget = combo;
        } else if (control.type == "CheckBox" || control.type == "OptionButton") {
            widget = new QCheckBox(control.caption, canvas);
        } else if (control.type == "CommandButton" || control.type == "ToggleButton") {
            auto* button = new QPushButton(control.caption.isEmpty() ? control.name : control.caption, canvas);
            widget = button;
            if (eventProcedure(control.onClick)) {
                const auto event = eventName(control.name) + "_Click";
                connect(button, &QPushButton::clicked, this, [this,event] { (void)runHandler(event); });
            }
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
        widget->setObjectName("access-control-" + control.name);
        widget->setProperty("la_access_tag",control.tag);
        widget->setGeometry(pixels(control.left), pixels(control.top),
                            qMax(12, pixels(control.width)), qMax(14, pixels(control.height)));
        widget->setVisible(control.visible);
        if (!control.name.isEmpty()) controls_.insert(control.name.toCaseFolded(), widget);
        const auto name = control.name;
        widget->setProperty("la_null",true);
        widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Null));
        if (auto* line = qobject_cast<QLineEdit*>(widget)) {
            connect(line,&QLineEdit::textChanged,this,[this,name,widget] {
                if (loading_) return;
                widget->setProperty("la_null",false); markDirty(name);
            });
            connect(line,&QLineEdit::textEdited,this,[this,name,widget] {
                if (!loading_) { pendingControls_.insert(name); widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text)); }
            });
            connect(line,&QLineEdit::editingFinished,this,[this,name] { if (!loading_) (void)commitControl(name); });
        } else if (auto* combo = qobject_cast<QComboBox*>(widget)) {
            connect(combo,&QComboBox::currentTextChanged,this,[this,name,widget] {
                if (loading_) return;
                widget->setProperty("la_null",false); markDirty(name);
            });
            connect(combo,QOverload<int>::of(&QComboBox::activated),this,[this,name](int) {
                if (loading_) return;
                if (auto* widget=controls_.value(name.toCaseFolded()).data()) widget->setProperty("la_null",false);
                markDirty(name); pendingControls_.insert(name); (void)commitControl(name);
            });
        } else if (auto* box = qobject_cast<QCheckBox*>(widget)) {
            connect(box,&QCheckBox::toggled,this,[this,name,widget] {
                if (loading_) return;
                widget->setProperty("la_null",false); markDirty(name);
            });
            connect(box,&QCheckBox::clicked,this,[this,name] { pendingControls_.insert(name); (void)commitControl(name); });
        }
        ++renderedControls_;
    }
}

void AccessFormPreview::loadRecords() {
    const bool wasNew = initialized_ && currentRecordIndex_ == 0;
    const auto oldId = currentRecordIndex_ > 0 && currentRecordIndex_ <= rowIds_.size()
        ? rowIds_[currentRecordIndex_-1] : QVariant{};
    const int oldIndex = currentRecordIndex_;
    recordSelector_->blockSignals(true);
    recordSelector_->clear(); recordRows_.clear(); rowIds_.clear(); dataColumns_.clear(); dataTypes_.clear();
    dataSourceAvailable_ = false; directTable_ = false; dataTruncated_ = false;
    dataSourceErrorNumber_=0;
    recordSelector_->addItem(layout_.kind==linaccess::AccessDocumentKind::Report ? "(Kein Datensatz)" : "(Neuer Datensatz)");
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
                auto result = linaccess::executeAccessQueryPlan(db_,plan,linaccess::bindAccessQueryParameters(plan,{}));
                dataColumns_ = result.columns;
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
                const auto plan=linaccess::compileAccessQuery("SELECT rowid AS [__la_rowid],* FROM ["+name+"]"+(filterOn_ && !filter_.trimmed().isEmpty() ? " WHERE "+filter_ : QString{}));
                if (!query.prepare(plan.sql) || !query.exec()) {
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

void AccessFormPreview::setFilterState(const QString& filter,bool enabled) {
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
    QScopedValueRollback<bool> loading(loading_,true);
    for (const auto& control : layout_.controls) {
        if (control.controlSource.isEmpty()) continue;
        auto* widget = controls_.value(control.name.toCaseFolded()).data();
        if (!widget) continue;
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
        if (auto* line = qobject_cast<QLineEdit*>(widget))
            line->setReadOnly(layout_.kind==linaccess::AccessDocumentKind::Report || control.controlSource.startsWith('=') || (!control.controlSource.isEmpty() && !directTable_));
        else if (qobject_cast<QCheckBox*>(widget) || qobject_cast<QComboBox*>(widget)) widget->setEnabled(layout_.kind!=linaccess::AccessDocumentKind::Report && (control.controlSource.isEmpty() || directTable_));
        setText(widget,value.toString()); widget->setProperty("la_null",value.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text));
    }
    refreshCalculatedControls();
    dirty_ = false; pendingControls_.clear();
}

void AccessFormPreview::markDirty(const QString& name) {
    if (loading_ || !directTable_ || !dataSourceAvailable_) return;
    if (!name.isEmpty()) {
        int boundColumn = -1;
        auto* sourceWidget = controls_.value(name.toCaseFolded()).data();
        for (const auto& control : layout_.controls)
            if (control.name.compare(name,Qt::CaseInsensitive) == 0)
                boundColumn = columnIndex(control.controlSource,dataColumns_);
        if (boundColumn < 0 || !sourceWidget) return;
        QScopedValueRollback<bool> loading(loading_,true);
        for (const auto& control : layout_.controls) {
            if (columnIndex(control.controlSource,dataColumns_) != boundColumn) continue;
            auto* target = controls_.value(control.name.toCaseFolded()).data();
            if (!target || target == sourceWidget) continue;
            setText(target,textOf(sourceWidget)); target->setProperty("la_null",sourceWidget->property("la_null"));
        }
    }
    if (!dirty_) {
        dirty_ = true;
        const auto event = formEvent("OnDirty","Form_Dirty");
        if (event.cancelled || failed(event)) loadRecord();
    }
}

bool AccessFormPreview::commitControl(const QString& name) {
    if (!pendingControls_.contains(name)) return true;
    for (const auto& control : layout_.controls) {
        if (control.name.compare(name,Qt::CaseInsensitive) != 0) continue;
        if (eventProcedure(control.beforeUpdate)) {
            const auto result = runHandler(eventName(control.name)+"_BeforeUpdate");
            if (result.cancelled || failed(result)) {
                if (auto* widget = controls_.value(control.name.toCaseFolded()).data()) widget->setFocus();
                return false;
            }
        }
        pendingControls_.remove(name);
        if (eventProcedure(control.afterUpdate)) {
            const auto result = runHandler(eventName(control.name)+"_AfterUpdate");
            if (failed(result)) return false;
        }
    }
    return true;
}
bool AccessFormPreview::commitControls() {
    const auto pending = pendingControls_;
    for (const auto& control : layout_.controls)
        if (pending.contains(control.name) && !commitControl(control.name)) return false;
    return true;
}

bool AccessFormPreview::saveRecord() {
    if (saving_) return false;
    QScopedValueRollback<bool> saving(saving_,true);
    if (!commitControls()) return false;
    if (!dirty_) return true;
    const auto table = layout_.recordSource.trimmed();
    if (!dataSourceAvailable_ || !directTable_) return false;
    const auto before = formEvent("BeforeUpdate","Form_BeforeUpdate");
    if (before.cancelled || failed(before)) return false;
    QStringList columns,placeholders;
    QVariantList values;
    for (const auto& control : layout_.controls) {
        const int fieldIndex = columnIndex(control.controlSource,dataColumns_);
        auto* widget = controls_.value(control.name.toCaseFolded()).data();
        if (!widget || fieldIndex < 0) continue;
        const auto column = sqlIdentifier(dataColumns_[fieldIndex]);
        if (columns.contains(column)) continue;
        columns << column; placeholders << "?";
        try {
            values << linaccess::sqlFieldValue(getFieldValue(control.name.toUtf8().toStdString()),widget->property("la_sql_type").toString());
        } catch (const std::exception& error) { status_->setText(QString::fromUtf8(error.what())); return false; }
    }
    if (columns.empty()) return false;
    QSqlQuery query(db_);
    const int row = currentRecordIndex_-1;
    QVariant savedId;
    if (row >= 0 && row < recordRows_.size()) {
        savedId = rowIds_[row]; QStringList assignments;
        for (const auto& column : columns) assignments << column + "=?";
        query.prepare("UPDATE " + sqlIdentifier(table) + " SET " + assignments.join(',') + " WHERE rowid=?");
        values << savedId;
    } else query.prepare("INSERT INTO " + sqlIdentifier(table) + " (" + columns.join(',') + ") VALUES (" + placeholders.join(',') + ")");
    for (const auto& value : values) query.addBindValue(value);
    if (!query.exec()) { status_->setText("SQL: " + query.lastError().text()); return false; }
    if (query.numRowsAffected() != 1) {
        status_->setText("Der Datensatz wurde nicht gespeichert; er wurde möglicherweise extern entfernt."); return false;
    }
    if (savedId.isNull()) savedId = query.lastInsertId();
    dirty_ = false;
    loadRecords();
    const int index = rowIds_.indexOf(savedId)+1;
    recordSelector_->blockSignals(true); recordSelector_->setCurrentIndex(index); recordSelector_->blockSignals(false);
    currentRecordIndex_ = index; loadRecord();
    (void)formEvent("AfterUpdate","Form_AfterUpdate");
    return true;
}

void AccessFormPreview::selectRecord(int target) {
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
}

bool AccessFormPreview::prepareClose() {
    if (closed_ || openCancelled_) return true;
    QVector<AccessFormPreview*> documents;
    const std::function<void(AccessFormPreview*)> collect=[&](AccessFormPreview* document) {
        if (document->closed_ || document->openCancelled_) return;
        documents.append(document);
        for (const auto& child : document->subforms_) if (child) collect(child);
    };
    collect(this);
    for (auto* document : documents) if (!document->saveRecord()) return false;
    for (auto* document : documents) {
        const auto unload=document->formEvent("OnUnload","Form_Unload");
        if (unload.cancelled || failed(unload)) return false;
    }
    if (activeDocument() && documents.contains(activeDocument().data())) {
        const auto active=activeDocument();
        activeDocument().clear();
        (void)active->formEvent("OnDeactivate","Form_Deactivate");
    }
    for (auto it=documents.crbegin();it!=documents.crend();++it) {
        (void)(*it)->formEvent("OnClose","Form_Close"); (*it)->closed_=true;
    }
    return true;
}
void AccessFormPreview::closeEvent(QCloseEvent* event) {
    if (!prepareClose()) event->ignore(); else QDialog::closeEvent(event);
}
void AccessFormPreview::accept() { if (prepareClose()) QDialog::accept(); }
void AccessFormPreview::reject() { if (prepareClose()) QDialog::reject(); }

std::optional<std::string> AccessFormPreview::getField(std::string_view name) const {
    auto* widget = controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded()).data();
    if (!widget) throw std::runtime_error("Unbekanntes Access-Control: " + std::string(name));
    const auto value = textOf(widget);
    if (widget->property("la_null").toBool()) return std::nullopt;
    return value.toUtf8().toStdString();
}

void AccessFormPreview::ensureWritableControl(std::string_view name) const {
    const auto key = QString::fromUtf8(name.data(),name.size());
    for (const auto& control : layout_.controls)
        if (control.name.compare(key,Qt::CaseInsensitive)==0 &&
            (control.controlSource.startsWith('=') || (!control.controlSource.isEmpty() && !directTable_)))
            throw linaccess::VbaRuntimeError(2448,"Bound control is read-only");
}
void AccessFormPreview::setField(std::string_view name, std::optional<std::string> value) {
    ensureWritableControl(name);
    auto* widget = controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded()).data();
    if (!widget) throw std::runtime_error("Unbekanntes Access-Control: " + std::string(name));
    const bool wasNull = widget->property("la_null").toBool();
    const auto previous = textOf(widget);
    {
        QScopedValueRollback<bool> loading(loading_,true);
        setText(widget,value ? QString::fromUtf8(value->data(),value->size()) : QString{});
        widget->setProperty("la_null",!value.has_value());
    }
    if (wasNull != !value.has_value() || previous != textOf(widget))
        markDirty(QString::fromUtf8(name.data(),name.size()));
}

linaccess::VbaValue AccessFormPreview::getFieldValue(std::string_view name) const {
    const auto key=QString::fromUtf8(name.data(),name.size());
    for (const auto& control : layout_.controls)
        if (control.name.compare(key,Qt::CaseInsensitive)==0 && control.controlSource.trimmed().startsWith('='))
            return calculatedControlValue(control.name);
    auto* widget = controls_.value(key.toCaseFolded()).data();
    if (!widget) {
        const auto index=columnIndex(key,dataColumns_);
        if (index<0) throw linaccess::VbaRuntimeError(2465,"Unknown form control or field");
        const auto row=currentRecordIndex_-1;
        return row>=0 && row<recordRows_.size() ? linaccess::qtFieldValue(recordRows_[row][index],dataTypes_.value(key.toCaseFolded())) : linaccess::VbaValue{};
    }
    if (!qobject_cast<QLineEdit*>(widget) && !qobject_cast<QComboBox*>(widget) && !qobject_cast<QCheckBox*>(widget))
        throw linaccess::VbaRuntimeError(438,"Control has no Value property");
    if (static_cast<linaccess::VbaValue::Kind>(widget->property("la_vba_kind").toInt())==linaccess::VbaValue::Kind::Error)
        return linaccess::VbaValue::errorValue(widget->property("la_error_number").toInt());
    if (widget->property("la_null").toBool()) return {};
    if (auto* combo=qobject_cast<QComboBox*>(widget);combo && combo->property("la_rows_loaded").toBool() && combo->currentIndex()>=0 && combo->currentText()==combo->itemText(combo->currentIndex()))
        return linaccess::qtFieldValue(combo->currentData(),widget->property("la_sql_type").toString());
    auto type = widget->property("la_sql_type").toString();
    if (type.isEmpty()) {
        const auto kind = static_cast<linaccess::VbaValue::Kind>(widget->property("la_vba_kind").toInt());
        if (kind == linaccess::VbaValue::Kind::Empty) return linaccess::VbaValue::empty();
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
            return value;
        }
        const auto converted=linaccess::sqlFieldValue(value,"");
        setText(widget,converted.toString()); widget->setProperty("la_null",converted.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
        widget->setProperty("la_calculation_error",0);
        return value;
    } catch (const std::exception& error) {
        const auto* typed=dynamic_cast<const linaccess::VbaRuntimeError*>(&error);
        const int number=typed ? typed->number : 5;
        setText(widget,"#Fehler ("+QString::number(number)+")"); widget->setProperty("la_calculation_error",number);
        widget->setProperty("la_null",false); widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Error));
        widget->setProperty("la_error_number",number);
        return linaccess::VbaValue::errorValue(number);
    }
}
void AccessFormPreview::refreshCalculatedControls() {
    for (const auto& control : layout_.controls)
        if (control.controlSource.trimmed().startsWith('=')) {
            try { (void)calculatedControlValue(control.name); } catch (const std::exception&) { /* Error is visible on the control. */ }
        }
}
linaccess::VbaValue AccessFormPreview::controlColumnValue(const QString& name,const std::vector<linaccess::VbaValue>& arguments) {
    if (arguments.empty() || arguments.size()>2) throw linaccess::VbaRuntimeError(450,"Column requires column and optional row");
    auto* combo=qobject_cast<QComboBox*>(controls_.value(name.toCaseFolded()).data());
    if (!combo) throw linaccess::VbaRuntimeError(438,"Control has no Column property");
    const auto index=[](const linaccess::VbaValue& value) {
        if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null column index");
        return linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
    };
    const auto column=index(arguments[0]);
    int count=1;
    for (const auto& control : layout_.controls) if (control.name.compare(name,Qt::CaseInsensitive)==0) count=control.columnCount;
    if (column<0 || column>=count) throw linaccess::VbaRuntimeError(5,"Column index out of range");
    if (arguments.size()==1 && combo->property("la_null").toBool()) return {};
    if (!combo->property("la_rows_loaded").toBool()) requeryControl(name.toUtf8().toStdString());
    const auto row=arguments.size()==2 ? index(arguments[1]) : combo->currentIndex();
    if (row<0 && arguments.size()==1) return {};
    if (row<0 || row>=combo->count()) throw linaccess::VbaRuntimeError(5,"Column row index out of range");
    const auto values=combo->itemData(row,Qt::UserRole+1).toList();
    if (column>=values.size()) throw linaccess::VbaRuntimeError(5,"Column not present in row source");
    return linaccess::qtFieldValue(values[column]);
}
void AccessFormPreview::setFieldValue(std::string_view name,const linaccess::VbaValue& value) {
    ensureWritableControl(name);
    auto* widget = controls_.value(QString::fromUtf8(name.data(),name.size()).toCaseFolded()).data();
    if (!widget) throw linaccess::VbaRuntimeError(2465,"Unknown form control");
    if (value.kind==linaccess::VbaValue::Kind::Error) {
        bool unbound=false;
        for (const auto& control : layout_.controls)
            if (control.name.compare(QString::fromUtf8(name.data(),name.size()),Qt::CaseInsensitive)==0) unbound=control.controlSource.isEmpty();
        if (!unbound || !qobject_cast<QLineEdit*>(widget)) throw linaccess::VbaRuntimeError(13,"Error Variant cannot be stored in a bound field");
        const QSignalBlocker blocked(widget);
        setText(widget,"#Fehler ("+QString::number(value.errorNumber)+")");
        widget->setProperty("la_null",false); widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
        widget->setProperty("la_error_number",value.errorNumber); refreshCalculatedControls(); return;
    }
    const auto type = widget->property("la_sql_type").toString();
    const auto converted = linaccess::sqlFieldValue(value,type);
    const auto before = textOf(widget); const bool wasNull = widget->property("la_null").toBool();
    {
        QScopedValueRollback<bool> loading(loading_,true);
        setText(widget,converted.toString()); widget->setProperty("la_null",converted.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
    }
    if (before != textOf(widget) || wasNull != converted.isNull()) markDirty(QString::fromUtf8(name.data(),name.size()));
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
        if (control.name.compare(target,Qt::CaseInsensitive)==0 && control.type=="CommandButton") {
            // A command button has no record/row source to execute. Requery
            // invalidates its presentation after Caption/Visible/Tag changes.
            widget->update(); return;
        }
    auto* combo = qobject_cast<QComboBox*>(widget);
    if (!combo) throw linaccess::VbaRuntimeError(438,"Requery for this control type is unsupported");
    for (const auto& control : layout_.controls) {
        if (control.name.compare(target,Qt::CaseInsensitive)!=0) continue;
        const auto source = control.rowSource.trimmed();
        linaccess::AccessQueryResult result;
        if (!control.rowSourceType.isEmpty() && control.rowSourceType.compare("Table/Query",Qt::CaseInsensitive)!=0 && control.rowSourceType.compare("Tabelle/Abfrage",Qt::CaseInsensitive)!=0)
            throw linaccess::VbaRuntimeError(3251,"Row source type unsupported");
        if (!source.isEmpty()) {
            linaccess::AccessQueryPlan plan;
            if (db_.tables().contains(source,Qt::CaseInsensitive)) plan.sql = "SELECT * FROM " + sqlIdentifier(source);
            else plan = linaccess::compileAccessQuery(
                source.startsWith("SELECT",Qt::CaseInsensitive) || source.startsWith("PARAMETERS",Qt::CaseInsensitive)
                    ? source : linaccess::accessQuerySource(db_,source));
            result = linaccess::executeAccessQueryPlan(db_,plan,linaccess::bindAccessQueryParameters(plan,{}));
            if (result.rows.size()>10000) throw linaccess::VbaRuntimeError(3035,"Control row limit exceeded");
        }
        if (control.boundColumn<0 || control.columnCount<1 || (!result.columns.empty() && control.boundColumn>result.columns.size()))
            throw linaccess::VbaRuntimeError(5,"Invalid row source column configuration");
        const auto oldValue=getFieldValue(target.toUtf8().toStdString());
        const auto old=linaccess::sqlFieldValue(oldValue,"");
        const QSignalBlocker blocked(combo);
        combo->clear();
        for (const auto& row : result.rows) if (!row.isEmpty()) {
            const auto bound=control.boundColumn==0 ? QVariant(combo->count()) : row[control.boundColumn-1];
            combo->addItem(row.front().toString(),bound);
            combo->setItemData(combo->count()-1,row,Qt::UserRole+1);
        }
        combo->setProperty("la_rows_loaded",true);
        setText(combo,old.toString()); combo->setProperty("la_null",old.isNull());
        return;
    }
}

void AccessFormPreview::openOtherForm(const QString& name) {
    QSqlQuery query(db_);
    query.prepare("SELECT source_bytes FROM _la_access_artifacts "
                  "WHERE kind='form' AND name=? ORDER BY id DESC LIMIT 1");
    query.addBindValue(name);
    if (!query.exec() || !query.next()) {
        status_->setText("DoCmd.OpenForm: Formular '" + name + "' nicht importiert.");
        return;
    }
    const auto form = query.value(0).toByteArray();
    QSqlQuery vba(db_);
    vba.prepare("SELECT source_bytes FROM _la_access_artifacts "
                "WHERE kind='vba' AND name=? ORDER BY id DESC LIMIT 1");
    vba.addBindValue("Form_" + name);
    const auto code = vba.exec() && vba.next() ? vba.value(0).toByteArray() : QByteArray{};
    AccessFormPreview nested(name, form, code, db_, this);
    if (!nested.cancelledByOpen()) nested.exec();
}

linaccess::VbaExecutionResult AccessFormPreview::executeProcedure(const QString& name,std::vector<linaccess::VbaValue>& arguments) {
    PreviewVbaHost host(*this);
    return linaccess::executeVbaProcedureWithArguments(vba_.toUtf8().toStdString(),name.toUtf8().toStdString(),host,arguments);
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
