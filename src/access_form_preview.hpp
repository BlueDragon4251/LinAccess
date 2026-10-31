#pragma once

#include "access_form_import.hpp"
#include "vba_runtime.hpp"

#include <QDialog>
#include <QMap>
#include <QPointer>
#include <QSet>
#include <QSqlDatabase>
#include <QVariant>

#include <optional>

class QComboBox;
class QLabel;
class QWidget;
class QCloseEvent;
class QToolBar;
class QVBoxLayout;
class QPrinter;
class QAction;

// Renders controls from Access 2003 SaveAsText instead of placing them in an
// unrelated auto-generated QFormLayout. Only supported VBA handlers can run.
class AccessFormPreview final : public QDialog {
public:
    AccessFormPreview(QString formName, const QByteArray& source,
                      const QByteArray& vbaSource, QSqlDatabase database,
                      QWidget* parent = nullptr);

    linaccess::VbaExecutionResult executeProcedure(const QString& name,std::vector<linaccess::VbaValue>& arguments);
    bool exportLayoutPdf(const QString& path);
    std::optional<linaccess::VbaExecutionResult> eventResult(const QString& property) const {
        const auto found = eventResults_.constFind(property.toCaseFolded());
        return found==eventResults_.cend() ? std::nullopt : std::optional<linaccess::VbaExecutionResult>(*found);
    }
    int renderedControls() const { return renderedControls_; }
    bool hasDataSource() const { return dataSourceAvailable_; }
    bool cancelledByOpen() const { return openCancelled_; }
    bool hasUnsavedChanges() const { return dirty_; }
    void accept() override;
    void reject() override;

private:
    void buildControls(QWidget* canvas);
    void loadRecords();
    void loadRecord();
    bool saveRecord();
    linaccess::VbaExecutionResult runHandler(const QString& handler);
    linaccess::VbaExecutionResult formEvent(const QString& property,const QString& handler);
    void markDirty(const QString& control);
    bool commitControl(const QString& control);
    bool commitControls();
    void selectRecord(int target);
    bool prepareClose();
    void closeEvent(QCloseEvent* event) override;
    void openOtherForm(const QString& name);
    std::optional<std::string> getField(std::string_view name) const;
    void ensureWritableControl(std::string_view name) const;
    void setField(std::string_view name, std::optional<std::string> value);
    linaccess::VbaValue getFieldValue(std::string_view name) const;
    void setFieldValue(std::string_view name,const linaccess::VbaValue& value);
    friend class PreviewVbaHost;
    friend void connectPreviewSqlObjects(QSqlDatabase db);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewFormObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewReportsObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewFormsObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewControlsObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewControlObject(AccessFormPreview&,const QString&);
    friend class PreviewApplicationObject;
    friend class PreviewCommandBarsObject;
    friend class PreviewCommandBarObject;
    friend class PreviewCommandControlsObject;
    friend class PreviewCommandControlObject;
    friend std::shared_ptr<linaccess::VbaObject> makePreviewCommandControlObject(AccessFormPreview&,QAction*);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewApplicationObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewCommandBarObject(AccessFormPreview&,QToolBar*);
    friend class PreviewObjectBase;
    friend class PreviewFormObject;
    friend class PreviewFormsObject;
    friend class PreviewControlsObject;
    friend class PreviewControlObject;
    void requeryControl(std::string_view name);

    std::weak_ptr<linaccess::VbaObject> formObject_, formsObject_, reportsObject_, controlsObject_;
    QMap<QString,std::weak_ptr<linaccess::VbaObject>> controlObjects_;
    std::weak_ptr<linaccess::VbaObject> applicationObject_;
    QMap<QToolBar*,std::weak_ptr<linaccess::VbaObject>> commandBarObjects_;
    QMap<QAction*,std::weak_ptr<linaccess::VbaObject>> commandControlObjects_;
    QToolBar* printToolbar_{};
    void buildPrintToolbar(QVBoxLayout* layout);
    bool renderLayout(QPrinter& printer);
    QMap<QString,linaccess::VbaExecutionResult> eventResults_;
    QString formName_;
    QString vba_;
    linaccess::AccessFormLayout layout_;
    QSqlDatabase db_;
    QComboBox* recordSelector_{};
    QLabel* status_{};
    QMap<QString,QPointer<QWidget>> controls_;
    QStringList dataColumns_;
    QMap<QString,QString> dataTypes_;
    QVector<QVariantList> recordRows_;
    QVector<QVariant> rowIds_;
    int renderedControls_{};
    bool dataSourceAvailable_{};
    bool directTable_{};
    bool openCancelled_{};
    bool opened_{};
    bool closed_{};
    bool dirty_{};
    bool loading_{};
    bool saving_{};
    bool navigating_{};
    bool initialized_{};
    bool dataTruncated_{};
    int currentRecordIndex_{};
    QSet<QString> pendingControls_;
};
