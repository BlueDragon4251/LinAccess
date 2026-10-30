#pragma once

#include <QJsonObject>
#include <QMainWindow>
#include <QSqlDatabase>
#include <QStringList>

class QTabWidget;
class QTreeWidget;
class QTreeWidgetItem;

class MainWindow final : public QMainWindow {
    Q_OBJECT
public:
    explicit MainWindow(QWidget* parent = nullptr);

private:
    QSqlDatabase db_;
    QString databasePath_;
    QTreeWidget* explorer_{};
    QTabWidget* tabs_{};

    bool hasProject();
    void openDatabase(const QString& path);
    void newDatabase();
    void chooseDatabase();
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
