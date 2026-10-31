#include "access_form_preview.hpp"
#include "access_form_objects.hpp"

#include "vba_runtime.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"

#include <QApplication>
#include <QCloseEvent>
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
#include <functional>
#include <stdexcept>

namespace {
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
    if (auto* combo = qobject_cast<QComboBox*>(editor)) return combo->currentText();
    return {};
}
void setText(QWidget* editor, const QString& text) {
    if (auto* line = qobject_cast<QLineEdit*>(editor)) line->setText(text);
    else if (auto* checkbox = qobject_cast<QCheckBox*>(editor))
        checkbox->setChecked(text == "-1" || text == "1" || text.compare("true", Qt::CaseInsensitive) == 0);
    else if (auto* combo = qobject_cast<QComboBox*>(editor)) combo->setCurrentText(text);
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
    void setControlValue(std::string_view name,const linaccess::VbaValue& value) override { preview_.setFieldValue(name,value); }
    linaccess::VbaValue fieldValue(std::string_view name) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (key == "name") return linaccess::VbaValue::string(preview_.formName_.toUtf8().toStdString());
        if (key == "caption") return makePreviewFormObject(preview_)->member("Caption",{});
        if (key == "recordsource") return linaccess::VbaValue::string(preview_.layout_.recordSource.toUtf8().toStdString());
        if (key == "dirty") return linaccess::VbaValue::logical(preview_.dirty_);
        if (key == "newrecord") return linaccess::VbaValue::logical(preview_.currentRecordIndex_ == 0);
        return preview_.getFieldValue(name);
    }
    void setFieldValue(std::string_view name,const linaccess::VbaValue& value) override {
        const auto key = QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if (key == "caption") { makePreviewFormObject(preview_)->assign("Caption",value); return; }
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
        QMessageBox::information(&preview_, "Access VBA", QString::fromUtf8(value.data(), value.size()));
    }
    void saveRecord() override {
        if (!preview_.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Datensatz konnte nicht gespeichert werden");
    }
    void quitApplication() override { QApplication::quit(); }
    void maximizeForm() override { preview_.showMaximized(); }
    void gotoLastRecord() override {
        if (preview_.recordSelector_->count() > 0)
            preview_.recordSelector_->setCurrentIndex(preview_.recordSelector_->count() - 1);
    }
    void gotoNextRecord() override {
        preview_.recordSelector_->setCurrentIndex(qMin(
            preview_.recordSelector_->currentIndex() + 1,
            preview_.recordSelector_->count() - 1));
    }
    void gotoNewRecord() override { preview_.recordSelector_->setCurrentIndex(0); }
    void requery() override {
        if (!preview_.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Requery cancelled by validation");
        preview_.loadRecords();
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
    std::shared_ptr<linaccess::VbaObject> formsCollection() override { return makePreviewFormsObject(preview_); }
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
                                     QWidget* parent)
    : QDialog(parent), formName_(std::move(formName)),
      vba_(linaccess::decodeVbaSource(vbaSource)),
      layout_(linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(source))),
      db_(std::move(database)) {
    static qlonglong openingSequence = 0;
    setProperty("la_form_open_order",++openingSequence);
    const bool queryViewsReady = linaccess::installAccessQueryViews(db_) >= 0;
    setWindowTitle((layout_.caption.isEmpty() ? formName_ : layout_.caption) + " — LinAccess");
    resize(950, 720);
    auto* outer = new QVBoxLayout(this);
    if (linaccess::decodeVbaSource(source).contains(QRegularExpression("^\\s*Begin\\s+Report\\b",QRegularExpression::CaseInsensitiveOption|QRegularExpression::MultilineOption)))
        buildPrintToolbar(outer);
    auto* controls = new QHBoxLayout();
    controls->addWidget(new QLabel("Datensatz", this));
    recordSelector_ = new QComboBox(this);
    recordSelector_->setObjectName("access-record-selector");
    controls->addWidget(recordSelector_, 1);
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
    outer->addWidget(status_);
    connect(recordSelector_, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, [this](int index) { selectRecord(index); });
    connect(save, &QPushButton::clicked, this, [this] { (void)saveRecord(); });
    if (!queryViewsReady) {
        status_->setText("Abfrageansichten werden noch verwendet; andere Ansichten schließen und Formular erneut öffnen.");
        canvas->setEnabled(false); save->setEnabled(false); return;
    }
    loadRecords();

    const auto outcome = formEvent("OnOpen","Form_Open");
    if (outcome.cancelled) { openCancelled_ = true; return; }
    (void)formEvent("OnLoad","Form_Load");
    opened_ = true;
    (void)formEvent("OnCurrent","Form_Current");
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
            auto* button = new QPushButton(control.name + " (Unterformular)", canvas);
            widget = button;
            const auto other = control.sourceObject;
            connect(button, &QPushButton::clicked, this, [this, other] {
                auto name = other;
                if (name.startsWith("Form.", Qt::CaseInsensitive)) name = name.mid(5);
                openOtherForm(name);
            });
        } else if (control.type == "Rectangle" || control.type == "Line") {
            auto* shape = new QLabel(canvas);
            shape->setFrameShape(control.type == "Rectangle" ? QFrame::Box : QFrame::HLine);
            widget = shape;
        }
        if (!widget) continue;
        widget->setObjectName("access-control-" + control.name);
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
                pendingControls_.insert(name); (void)commitControl(name);
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
    recordSelector_->addItem("(Neuer Datensatz)");
    const auto source = layout_.recordSource.trimmed();
    if (!source.isEmpty() && db_.isOpen()) {
        try {
            directTable_ = db_.tables().contains(source,Qt::CaseInsensitive);
            if (!directTable_) {
                linaccess::AccessQueryPlan plan;
                if (db_.tables(QSql::Views).contains(source,Qt::CaseInsensitive))
                    plan.sql = "SELECT * FROM " + sqlIdentifier(source);
                else {
                    const auto sql = source.startsWith("SELECT",Qt::CaseInsensitive)
                        || source.startsWith("PARAMETERS",Qt::CaseInsensitive)
                        ? source : linaccess::accessQuerySource(db_,source);
                    plan = linaccess::compileAccessQuery(sql);
                }
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
                if (!query.exec("SELECT rowid AS __la_rowid,* FROM " + sqlIdentifier(source)))
                    throw std::runtime_error("Record source unavailable");
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
        } catch (const std::exception&) {
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
            line->setReadOnly(control.controlSource.startsWith('=') || (index >= 0 && !directTable_));
        else if (qobject_cast<QCheckBox*>(widget) || qobject_cast<QComboBox*>(widget)) widget->setEnabled(index < 0 || directTable_);
        setText(widget,value.toString()); widget->setProperty("la_null",value.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(linaccess::VbaValue::Kind::Text));
    }
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
    if (opened_) (void)formEvent("OnCurrent","Form_Current");
}

bool AccessFormPreview::prepareClose() {
    if (closed_ || openCancelled_) return true;
    if (!saveRecord()) return false;
    const auto unload = formEvent("OnUnload","Form_Unload");
    if (unload.cancelled || failed(unload)) return false;
    (void)formEvent("OnClose","Form_Close"); closed_ = true; return true;
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

void AccessFormPreview::setField(std::string_view name, std::optional<std::string> value) {
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
    auto* widget = controls_.value(QString::fromUtf8(name.data(),name.size()).toCaseFolded()).data();
    if (!widget) throw linaccess::VbaRuntimeError(2465,"Unknown form control");
    if (widget->property("la_null").toBool()) return {};
    auto type = widget->property("la_sql_type").toString();
    if (type.isEmpty()) {
        const auto kind = static_cast<linaccess::VbaValue::Kind>(widget->property("la_vba_kind").toInt());
        if (kind == linaccess::VbaValue::Kind::Empty) return linaccess::VbaValue::empty();
        if (kind == linaccess::VbaValue::Kind::Number) type = "REAL";
        if (kind == linaccess::VbaValue::Kind::Boolean || qobject_cast<QCheckBox*>(widget)) type = "BOOLEAN";
    }
    return linaccess::qtFieldValue(textOf(widget),type);
}
void AccessFormPreview::setFieldValue(std::string_view name,const linaccess::VbaValue& value) {
    auto* widget = controls_.value(QString::fromUtf8(name.data(),name.size()).toCaseFolded()).data();
    if (!widget) throw linaccess::VbaRuntimeError(2465,"Unknown form control");
    const auto type = widget->property("la_sql_type").toString();
    const auto converted = linaccess::sqlFieldValue(value,type);
    const auto before = textOf(widget); const bool wasNull = widget->property("la_null").toBool();
    {
        QScopedValueRollback<bool> loading(loading_,true);
        setText(widget,converted.toString()); widget->setProperty("la_null",converted.isNull());
        widget->setProperty("la_vba_kind",static_cast<int>(value.kind));
    }
    if (before != textOf(widget) || wasNull != converted.isNull()) markDirty(QString::fromUtf8(name.data(),name.size()));
}

void AccessFormPreview::requeryControl(std::string_view name) {
    const auto target = QString::fromUtf8(name.data(),name.size());
    auto* widget = controls_.value(target.toCaseFolded()).data();
    if (!widget) throw linaccess::VbaRuntimeError(2465,"Control unavailable");
    auto* combo = qobject_cast<QComboBox*>(widget);
    if (!combo) throw linaccess::VbaRuntimeError(438,"Requery for this control type is unsupported");
    for (const auto& control : layout_.controls) {
        if (control.name.compare(target,Qt::CaseInsensitive)!=0) continue;
        const auto source = control.rowSource.trimmed();
        linaccess::AccessQueryResult result;
        if (!source.isEmpty()) {
            linaccess::AccessQueryPlan plan;
            if (db_.tables().contains(source,Qt::CaseInsensitive)) plan.sql = "SELECT * FROM " + sqlIdentifier(source);
            else plan = linaccess::compileAccessQuery(
                source.startsWith("SELECT",Qt::CaseInsensitive) || source.startsWith("PARAMETERS",Qt::CaseInsensitive)
                    ? source : linaccess::accessQuerySource(db_,source));
            result = linaccess::executeAccessQueryPlan(db_,plan,linaccess::bindAccessQueryParameters(plan,{}));
            if (result.rows.size()>10000) throw linaccess::VbaRuntimeError(3035,"Control row limit exceeded");
        }
        const auto old = combo->currentText();
        const QSignalBlocker blocked(combo);
        combo->clear();
        for (const auto& row : result.rows) if (!row.isEmpty()) combo->addItem(row.front().toString());
        combo->setCurrentText(old);
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
    return runHandler(handler);
}
