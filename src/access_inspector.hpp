#pragma once

#include "access_catalog.hpp"
#include "access_ole_package.hpp"
#include "csv.hpp"

#include <QDialog>
#include <QByteArray>
#include <QSqlDatabase>
#include <QStringList>

#include <string>
#include <vector>

class QLabel;
class QListWidget;
class QTextBrowser;
class QTableWidget;
class QPushButton;
class QPlainTextEdit;

class AccessInspector final : public QDialog {
    Q_OBJECT
public:
    AccessInspector(QString sourceFile, QSqlDatabase projectDb, QWidget* parent = nullptr);

private:
    QString sourceFile_;
    QSqlDatabase projectDb_;
    QListWidget* tables_{};
    QListWidget* objects_{};
    QPlainTextEdit* objectPreview_{};
    QTableWidget* preview_{};
    QTextBrowser* status_{};
    QPushButton* importButton_{};
    QPushButton* extractPackageButton_{};
    QPushButton* catalogImportButton_{};
    QString selectedTable_;
    std::vector<std::vector<linaccess::CsvField>> csvRows_;
    QVector<linaccess::AccessCatalogEntry> catalog_;
    std::optional<linaccess::AccessOlePackage> selectedPackage_;

    void loadTables();
    void previewTable();
    void updatePackageSelection();
    void extractPackage();
    void importTable();
    void importAllTables();
    void loadCatalog();
    void previewObject();
    void importCatalog();
    QByteArray runMdbTool(const QString& name, const QStringList& arguments);
};
