#pragma once

#include "access_form_import.hpp"

#include <QDialog>
#include <QMap>
#include <QSqlDatabase>
#include <QVariant>

#include <optional>

class QComboBox;
class QLabel;
class QWidget;

// Renders controls from Access 2003 SaveAsText instead of placing them in an
// unrelated auto-generated QFormLayout. Only supported VBA handlers can run.
class AccessFormPreview final : public QDialog {
public:
    AccessFormPreview(QString formName, const QByteArray& source,
                      const QByteArray& vbaSource, QSqlDatabase database,
                      QWidget* parent = nullptr);

    int renderedControls() const { return renderedControls_; }
    bool hasDataSource() const { return dataSourceAvailable_; }
    bool cancelledByOpen() const { return openCancelled_; }

private:
    void buildControls(QWidget* canvas);
    void loadRecords();
    void loadRecord();
    bool saveRecord();
    void runHandler(const QString& handler);
    void openOtherForm(const QString& name);
    std::optional<std::string> getField(std::string_view name) const;
    void setField(std::string_view name, std::optional<std::string> value);
    friend class PreviewVbaHost;

    QString formName_;
    QString vba_;
    linaccess::AccessFormLayout layout_;
    QSqlDatabase db_;
    QComboBox* recordSelector_{};
    QLabel* status_{};
    QMap<QString, QWidget*> controls_;
    QStringList dataColumns_;
    QVector<QVariantList> recordRows_;
    QVector<QVariant> rowIds_;
    int renderedControls_{};
    bool dataSourceAvailable_{};
    bool directTable_{};
    bool openCancelled_{};
};
