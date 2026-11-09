#pragma once

#include <QJsonObject>
#include "native_project_close_guard.hpp"
#include <QMainWindow>
#include <QSqlDatabase>
#include <QStringList>
#include <QPointer>
#include <QVariantMap>
#include <QSet>

class QTabWidget;
class QTreeWidget;
class QTreeWidgetItem;
class QKeyEvent;

class MainWindow final : public QMainWindow, public linaccess::NativeProjectCloseGuard {
    Q_OBJECT
public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;
    bool allowProjectClose(const QString& connection) override;
    Q_INVOKABLE void focusDatabaseWindow();
    Q_INVOKABLE void cycleDocument(bool backwards);
    Q_INVOKABLE int activateOpenTable(const QString& name);
    Q_INVOKABLE int activateOpenQuery(const QString& name);
    Q_INVOKABLE int openTableDocument(const QString& name,int view,int mode);
    Q_INVOKABLE QVariantMap openQueryDocument(const QString& name,int view,int mode);
    Q_INVOKABLE QVariantMap closeQueryDocument(int type,const QString& name,int save);
    Q_INVOKABLE QVariantMap dispatchTableCommand(int command);
    Q_INVOKABLE QVariantMap closeTableDocument(int type,const QString& name,int save);
    Q_INVOKABLE QVariantMap navigateQueryDocument(const QString& name,int action,int offset);
    Q_INVOKABLE QVariantMap navigateTableDocument(const QString& name,int action,int offset,bool explicitTable);
    Q_INVOKABLE QVariantMap findTableRecord(const QVariantMap& request);
    Q_INVOKABLE QVariantMap tableKeyboardContext();
    Q_INVOKABLE QVariantMap requeryTableDocument();
    Q_INVOKABLE QVariantMap outputDatasheet(bool pdf,const QString& path = {});
    Q_INVOKABLE QVariantMap previewDatasheet();
    Q_INVOKABLE QVariantMap setupDatasheetPage();
    Q_INVOKABLE int selectDatabaseObject(int objectType,const QString& name,bool activate);
    Q_INVOKABLE QVariantMap renameDatabaseObject(const QString& newName,int objectType,const QString& oldName);
    Q_INVOKABLE QVariantMap deleteDatabaseObject(int objectType,const QString& name);

private:
    bool eventFilter(QObject* watched,QEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;
    void closeEvent(QCloseEvent* event) override;
    QSqlDatabase db_;
    QString databasePath_;
    QTreeWidget* explorer_{};
    QTabWidget* tabs_{};
    QPointer<QWidget> activeDataPage_;
    int queryExecutionDepth_{};
    bool datasheetOutputActive_{};
    QSet<QString> executingQueries_;

    QVariantMap runProjectMacro(const QString& name);
    QVariantMap outputDatasheetInternal(int mode,const QString& path);
    bool finishActiveDataEditors();
    bool activateDataPage(QWidget* page);
    QWidget* findOpenTable(const QString& name) const;
    QWidget* findOpenQuery(const QString& name) const;
    QVariantMap loadQueryDocument(const QString& name,int view,int mode,QWidget* refresh);
    QVariantMap saveDataPage(QWidget* page,bool interactive,bool currentOnly=false);
    QVariantMap closeTableDataPage(QWidget* page);
    QVariantMap requeryTableDataPage(QWidget* page);

    bool canCloseDataPage(QWidget* page);
    bool hasProject();
    void openDatabase(const QString& path);
    Q_INVOKABLE void newDatabase();
    Q_INVOKABLE void chooseDatabase();
    void inspectAccess();
    void inspectBridgeExport();
    void showAccessArtifact(int id);
    void refreshExplorer();
    QStringList tableNames();
    QStringList objectNames(const QString& kind);
    QJsonObject objectDefinition(const QString& kind, const QString& name);
    bool saveObject(const QString& kind, const QString& name, const QJsonObject& definition);
    void newTable();
    int showTable(const QString& name,bool readOnly=false,bool interactive=true,const QString& queryName={},const QString& querySource={});
    void editQuery(const QString& name = {});
    void editMacro(const QString& name = {});
    void editModule(const QString& name = {});
    void editForm(const QString& name = {});
    void runForm(const QString& name);
    void runReport(const QString& name);
    void editReport(const QString& name = {});
    void openExplorerItem(QTreeWidgetItem* item, int column);
};
