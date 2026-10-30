#pragma once

#include <QDialog>
#include <QJsonArray>
#include <QJsonObject>
#include <QSqlDatabase>

class QLabel;
class QTableWidget;
class QTextBrowser;

class BridgeImporter final : public QDialog {
    Q_OBJECT
public:
    BridgeImporter(QString directory, QSqlDatabase project, QWidget* parent = nullptr);

private:
    QString directory_;
    QSqlDatabase project_;
    QJsonObject manifest_;
    QJsonArray objects_;
    QTableWidget* catalog_{};
    QTextBrowser* preview_{};
    QLabel* status_{};

    bool load();
    QString artifactPath(const QJsonObject& object) const;
    void previewSelected();
    void importAll();
};
