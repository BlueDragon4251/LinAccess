#include "access_form_preview.hpp"

#include "vba_runtime.hpp"
#include "dao_database.hpp"
#include "access_query.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"

#include <QApplication>
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
    std::optional<std::string> field(std::string_view name) override { return preview_.getField(name); }
    void setField(std::string_view name, std::optional<std::string> value) override {
        preview_.setField(name, std::move(value));
    }
    void message(std::string_view value) override {
        QMessageBox::information(&preview_, "Access VBA", QString::fromUtf8(value.data(), value.size()));
    }
    void saveRecord() override {
        if (!preview_.saveRecord()) throw std::runtime_error("Datensatz konnte nicht gespeichert werden");
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
    void requery() override { preview_.loadRecords(); }
    void requeryControl(std::string_view name) override {
        const auto target = QString::fromUtf8(name.data(), name.size());
        auto* widget = preview_.controls_.value(target.toCaseFolded());
        if (!widget) throw std::runtime_error("Steuerelement nicht gefunden: " + std::string(name));
        auto* combo = qobject_cast<QComboBox*>(widget);
        if (!combo) { widget->update(); return; }
        for (const auto& control : preview_.layout_.controls) {
            if (control.name.compare(target, Qt::CaseInsensitive) != 0) continue;
            if (!control.rowSource.startsWith("SELECT ", Qt::CaseInsensitive)) return;
            const auto old = combo->currentText();
            QSqlQuery query(preview_.db_);
            if (!query.exec(control.rowSource))
                throw std::runtime_error(query.lastError().text().toStdString());
            combo->clear();
            while (combo->count() < 10000 && query.next()) combo->addItem(query.value(0).toString());
            combo->setCurrentText(old);
            return;
        }
    }
    void focusControl(std::string_view name) override {
        auto* widget = preview_.controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded());
        if (!widget) throw std::runtime_error("Steuerelement nicht gefunden: " + std::string(name));
        widget->setFocus();
    }
    void setControlVisible(std::string_view name, bool value) override {
        auto* widget = preview_.controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded());
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
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(preview_.db_); }
    void executeSql(std::string_view statement) override {
        if (!preview_.db_.isOpen()) throw std::runtime_error("Kein offenes Datenbankprojekt");
        auto database = linaccess::makeDaoDatabase(preview_.db_);
        (void)database->member("Execute",{linaccess::VbaValue::string(std::string(statement))});
        preview_.loadRecords();
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
    const bool queryViewsReady = linaccess::installAccessQueryViews(db_) >= 0;
    setWindowTitle((layout_.caption.isEmpty() ? formName_ : layout_.caption) + " — LinAccess");
    resize(950, 720);
    auto* outer = new QVBoxLayout(this);
    auto* controls = new QHBoxLayout();
    controls->addWidget(new QLabel("Datensatz", this));
    recordSelector_ = new QComboBox(this);
    recordSelector_->setObjectName("access-record-selector");
    controls->addWidget(recordSelector_, 1);
    auto* save = new QPushButton("Datensatz speichern", this);
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
            this, [this] { loadRecord(); });
    connect(save, &QPushButton::clicked, this, [this] { (void)saveRecord(); });
    if (!queryViewsReady) {
        status_->setText("Abfrageansichten werden noch verwendet; andere Ansichten schließen und Formular erneut öffnen.");
        canvas->setEnabled(false); save->setEnabled(false); return;
    }
    loadRecords();

    // Access opens the form before its Form_Load event and honors Form_Open(Cancel).
    // No additional LinAccess consent dialog is shown for these handlers.
    if (!vba_.isEmpty()) {
        PreviewVbaHost host(*this);
        const auto outcome = linaccess::executeVbaProcedure(
            vba_.toUtf8().toStdString(), "Form_Open", host);
        if (outcome.cancelled) { openCancelled_ = true; return; }
        if (outcome.status == linaccess::VbaExecutionStatus::Unsupported
            || outcome.status == linaccess::VbaExecutionStatus::Error)
            status_->setText("Form_Open: " + QString::fromStdString(outcome.explanation));
        runHandler("Form_Load");
        runHandler("Form_Current");
    }
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
            auto eventName = control.name;
            eventName.replace(QRegularExpression("[^\\p{L}\\p{N}_]"), "_");
            const auto event = eventName + "_Click";
            connect(button, &QPushButton::clicked, this, [this, event] { runHandler(event); });
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
        ++renderedControls_;
    }
}

void AccessFormPreview::loadRecords() {
    recordSelector_->blockSignals(true);
    recordSelector_->clear();
    recordRows_.clear();
    rowIds_.clear();
    dataColumns_.clear();
    dataSourceAvailable_ = false;
    directTable_ = false;
    recordSelector_->addItem("(Neuer Datensatz)");
    const auto source = layout_.recordSource.trimmed();
    if (!source.isEmpty() && db_.isOpen()) {
        QString sql;
        if (source.startsWith("SELECT ", Qt::CaseInsensitive)) sql = source;
        else {
            // The record source may be a saved Access query that was imported
            // only as an artifact. We do not substitute a different query.
            const auto tables = db_.tables();
            directTable_ = tables.contains(source, Qt::CaseInsensitive);
            sql = directTable_ ? "SELECT rowid AS __la_rowid, * FROM " + sqlIdentifier(source)
                               : "SELECT * FROM " + sqlIdentifier(source);
        }
        QSqlQuery query(db_);
        if (query.exec(sql)) {
            dataSourceAvailable_ = true;
            const int offset = directTable_ ? 1 : 0;
            for (int i = offset; i < query.record().count(); ++i)
                dataColumns_ << query.record().fieldName(i);
            while (recordRows_.size() < 10000 && query.next()) {
                QVariantList row;
                if (directTable_) rowIds_.push_back(query.value(0));
                for (int i = 0; i < dataColumns_.size(); ++i) row << query.value(i + offset);
                recordRows_.push_back(row);
                recordSelector_->addItem("Datensatz " + QString::number(recordRows_.size()));
            }
        }
    }
    recordSelector_->blockSignals(false);
    loadRecord();
    if (status_) {
        status_->setText(dataSourceAvailable_
            ? "Access-Layout: " + QString::number(renderedControls_)
                + " sichtbare Elemente, " + QString::number(recordRows_.size()) + " Datensätze."
            : "Access-Layout geladen; die Datenquelle '" + source
                + "' ist im SQLite-Projekt nicht als ausführbare Abfrage verfügbar.");
    }
}

void AccessFormPreview::loadRecord() {
    for (const auto& control : layout_.controls) {
        if (control.controlSource.isEmpty()) continue;
        auto* widget = controls_.value(control.name.toCaseFolded());
        if (!widget) continue;
        QString value;
        const int index = columnIndex(control.controlSource, dataColumns_);
        const int row = recordSelector_->currentIndex() - 1;
        if (row >= 0 && row < recordRows_.size() && index >= 0)
            value = recordRows_[row][index].toString();
        setText(widget, value);
    }
}

bool AccessFormPreview::saveRecord() {
    const auto table = layout_.recordSource.trimmed();
    if (!dataSourceAvailable_ || !directTable_) {
        status_->setText("Speichern: Das Formular benötigt eine direkt gebundene, schreibbare Tabelle.");
        return false;
    }
    QStringList columns, placeholders;
    QVariantList values;
    for (const auto& control : layout_.controls) {
        const auto fieldIndex = columnIndex(control.controlSource, dataColumns_);
        auto* widget = controls_.value(control.name.toCaseFolded());
        if (!widget || fieldIndex < 0) continue;
        const auto column = dataColumns_[fieldIndex];
        if (columns.contains(sqlIdentifier(column))) continue;
        columns << sqlIdentifier(column);
        placeholders << "?";
        values << textOf(widget);
    }
    if (columns.empty()) return false;
    QSqlQuery query(db_);
    const int row = recordSelector_->currentIndex() - 1;
    if (row >= 0 && row < recordRows_.size()) {
        QStringList assignments;
        for (const auto& name : columns) assignments << name + "=?";
        query.prepare("UPDATE " + sqlIdentifier(table) + " SET " + assignments.join(',')
                      + " WHERE rowid=?");
        values << rowIds_.at(row);
    } else {
        query.prepare("INSERT INTO " + sqlIdentifier(table) + " (" + columns.join(',')
                      + ") VALUES (" + placeholders.join(',') + ")");
    }
    for (const auto& value : values) query.addBindValue(value);
    if (!query.exec()) {
        status_->setText("SQL: " + query.lastError().text());
        return false;
    }
    const auto savedRow = row >= 0 ? row : static_cast<int>(recordRows_.size());
    loadRecords();
    recordSelector_->setCurrentIndex(qMin(savedRow + 1, recordSelector_->count() - 1));
    return true;
}

std::optional<std::string> AccessFormPreview::getField(std::string_view name) const {
    auto* widget = controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded());
    if (!widget) throw std::runtime_error("Unbekanntes Access-Control: " + std::string(name));
    const auto value = textOf(widget);
    if (value.isEmpty()) return std::nullopt;
    return value.toUtf8().toStdString();
}

void AccessFormPreview::setField(std::string_view name, std::optional<std::string> value) {
    auto* widget = controls_.value(QString::fromUtf8(name.data(), name.size()).toCaseFolded());
    if (!widget) throw std::runtime_error("Unbekanntes Access-Control: " + std::string(name));
    setText(widget, value ? QString::fromUtf8(value->data(), value->size()) : QString{});
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

void AccessFormPreview::runHandler(const QString& handler) {
    if (vba_.isEmpty()) return;
    PreviewVbaHost host(*this);
    const auto result = linaccess::executeVbaProcedure(
        vba_.toUtf8().toStdString(), handler.toUtf8().toStdString(), host);
    if (result.status == linaccess::VbaExecutionStatus::Unsupported
        || result.status == linaccess::VbaExecutionStatus::Error) {
        const auto message = handler + " (Zeile " + QString::number(result.line)
            + "): " + QString::fromStdString(result.explanation);
        if (status_) status_->setText(message);
        else QMessageBox::critical(this, "VBA", message);
    }
}
