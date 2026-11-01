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
class QLineEdit;
class QCheckBox;
class QKeyEvent;

// Renders controls from Access 2003 SaveAsText instead of placing them in an
// unrelated auto-generated QFormLayout. Only supported VBA handlers can run.
class AccessFormPreview final : public QDialog {
public:
    AccessFormPreview(QString formName, const QByteArray& source,
                      const QByteArray& vbaSource, QSqlDatabase database,
                      QWidget* parent = nullptr,bool embedded=false,QString whereCondition={});

    linaccess::VbaExecutionResult executeProcedure(const QString& name,std::vector<linaccess::VbaValue>& arguments);
    bool exportLayoutPdf(const QString& path);
    std::optional<linaccess::VbaExecutionResult> eventResult(const QString& property) const {
        const auto found = eventResults_.constFind(property.toCaseFolded());
        return found==eventResults_.cend() ? std::nullopt : std::optional<linaccess::VbaExecutionResult>(*found);
    }
    int renderedControls() const { return renderedControls_; }
    bool hasDataSource() const { return dataSourceAvailable_; }
    int dataSourceErrorNumber() const { return dataSourceErrorNumber_; }
    bool cancelledByOpen() const { return openCancelled_; }
    bool hasUnsavedChanges() const { return dirty_; }
    static bool closeProjectDocuments(const QSqlDatabase& database);
    static bool closeApplicationDocuments();
    static void quitApplication();
    void accept() override;
    void reject() override;

private:
    QString controlMemberName(const QString& name) const;
    QString controlEventPrefix(const QString& name) const;
    std::shared_ptr<linaccess::VbaObject> recordsetSnapshot();
    linaccess::VbaValue evalExpression(std::string_view expression);
    void buildControls(QWidget* canvas);
    void loadRecords();
    void setFilterState(const QString& filter,bool enabled);
    void loadRecord();
    bool saveRecord();
    void selectCurrentRecordForCommand();
    void deleteCurrentRecord();
    void clearRecordSelection();
    linaccess::VbaExecutionResult runHandler(const QString& handler);
    linaccess::VbaExecutionResult runControlBinding(const QString& control,const QString& event,const QString& binding);
    bool eventFilter(QObject* watched,QEvent* event) override;
    linaccess::VbaExecutionResult formEvent(const QString& property,const QString& handler);
    void markDirty(const QString& control);
    bool commitControl(const QString& control);
    bool commitControls();
    void selectRecord(int target);
    bool prepareClose();
    static bool closeDocuments(const QVector<AccessFormPreview*>& roots,AccessFormPreview* closingRoot=nullptr);
    bool event(QEvent* event) override;
    void activateDocument();
    void initialCurrent();
    void closeEvent(QCloseEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;
    void keyReleaseEvent(QKeyEvent* event) override;
    bool dispatchKeyboardEvent(QObject* watched,QKeyEvent* event);
    bool ownsKeyboardFocus() const;
    bool handleDocumentKey(QKeyEvent* event);
    void rememberControlEdit(QObject* watched);
    void undoKeyboardEdit();
    void updateDirtyFromRecord();
    void openOtherForm(const QString& name,const QString& where={});
    void closeOtherDocument(int type,const std::optional<QString>& name,int save);
    std::optional<linaccess::VbaValue> getModuleMember(std::string_view name,bool publicOnly);
    bool setModuleMember(std::string_view name,const linaccess::VbaValue& value,bool publicOnly);
    std::optional<std::string> getField(std::string_view name) const;
    void ensureWritableControl(std::string_view name) const;
    void setField(std::string_view name, std::optional<std::string> value);
    linaccess::VbaValue getFieldValue(std::string_view name) const;
    linaccess::VbaValue calculatedControlValue(const QString& name) const;
    void refreshCalculatedControls();
    linaccess::VbaValue controlColumnValue(const QString& name,const std::vector<linaccess::VbaValue>& arguments);
    void setFieldValue(std::string_view name,const linaccess::VbaValue& value);
    friend class PreviewVbaHost;
    friend void connectPreviewSqlObjects(QSqlDatabase db);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewFormObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewReportsObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewFormsObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewControlsObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewControlObject(AccessFormPreview&,const QString&);
    friend class PreviewApplicationObject;
    friend class PreviewDoCmdObject;
    friend class PreviewScreenObject;
    friend class PreviewCommandBarsObject;
    friend class PreviewCommandBarObject;
    friend class PreviewCommandControlsObject;
    friend class PreviewCommandControlObject;
    friend std::shared_ptr<linaccess::VbaObject> makePreviewCommandControlObject(AccessFormPreview&,QAction*);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewApplicationObject(AccessFormPreview&);
    friend std::shared_ptr<linaccess::VbaObject> makePreviewCommandBarObject(AccessFormPreview&,QWidget*);
    friend class PreviewObjectBase;
    friend class PreviewFormObject;
    friend class PreviewRecordsetObject;
    friend class PreviewFormsObject;
    friend class PreviewControlsObject;
    friend class PreviewControlObject;
    void requeryControl(std::string_view name);

    std::weak_ptr<linaccess::VbaObject> formObject_, formsObject_, reportsObject_, controlsObject_;
    QMap<QString,std::weak_ptr<linaccess::VbaObject>> controlObjects_;
    std::unordered_map<std::string,std::shared_ptr<linaccess::VbaModuleState>> moduleStates_;
    std::weak_ptr<linaccess::VbaObject> applicationObject_,doCmdObject_,screenObject_;
    QMap<QWidget*,std::weak_ptr<linaccess::VbaObject>> commandBarObjects_;
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
    QLineEdit* filterEditor_{};
    QCheckBox* filterToggle_{};
    QString filter_;
    bool filterOn_{};
    bool changingFilter_{};
    QLabel* status_{};
    QMap<QString,QPointer<QWidget>> controls_;
    QMap<QString,QPointer<AccessFormPreview>> subforms_;
    QMap<QString,int> subformErrors_;
    mutable QSet<QString> calculatingControls_;
    mutable QMap<QString,std::shared_ptr<linaccess::VbaValue>> calculatedValues_;
    int dataSourceErrorNumber_{};
    QStringList dataColumns_;
    QMap<QString,QString> dataTypes_;
    QVector<QVariantList> recordRows_;
    QVector<QVariant> rowIds_;
    std::shared_ptr<linaccess::VbaObject> recordsetSnapshot_;
    int renderedControls_{};
    bool dataSourceAvailable_{};
    bool directTable_{};
    bool openCancelled_{};
    bool opened_{};
    bool nativeActivationPending_{};
    bool initialCurrentDeferred_{};
    bool initialCurrentDispatched_{};
    bool closed_{};
    bool closing_{};
    bool dirty_{};
    bool loading_{};
    bool saving_{};
    bool recordCommandRunning_{};
    bool recordSelected_{};
    bool navigating_{};
    bool initialized_{};
    bool deliveringKeyReplacement_{};
    bool dataTruncated_{};
    int currentRecordIndex_{};
    QSet<QString> pendingControls_;
    QMap<QString,std::shared_ptr<linaccess::VbaValue>> editOriginals_;
};
