#pragma once

#include <QJsonObject>
#include "native_project_close_guard.hpp"
#include <QMainWindow>
#include <QSqlDatabase>
#include <QStringList>
#include <QPointer>

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
    Q_INVOKABLE int selectDatabaseObject(int objectType,const QString& name,bool activate);

private:
    bool eventFilter(QObject* watched,QEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;
    void closeEvent(QCloseEvent* event) override;
    QSqlDatabase db_;
    QString databasePath_;
    QTreeWidget* explorer_{};
    QTabWidget* tabs_{};
    QPointer<QWidget> activeDataPage_;

    bool finishActiveDataEditors();
    bool activateDataPage(QWidget* page);
    QWidget* findOpenTable(const QString& name) const;

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
    void showTable(const QString& name);
    void editQuery(const QString& name = {});
    void editForm(const QString& name = {});
    void runForm(const QString& name);
    void editReport(const QString& name = {});
    void openExplorerItem(QTreeWidgetItem* item, int column);
};
