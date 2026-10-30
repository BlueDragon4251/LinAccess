#include "access_inspector.hpp"

#include <QApplication>
#include <QListWidget>
#include <QMessageBox>
#include <QPushButton>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QTemporaryDir>
#include <QTimer>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc, char** argv) {
    QApplication app(argc, argv);
    assert(argc == 2);
    QTemporaryDir temp;
    assert(temp.isValid());
    auto db = QSqlDatabase::addDatabase("QSQLITE", "mdb-gui-test");
    db.setDatabaseName(temp.filePath("project.linaccess"));
    assert(db.open());
    QSqlQuery init(db);
    assert(init.exec("CREATE TABLE _la_access_artifacts ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, "
        "source_sha256 TEXT NOT NULL, source_filename TEXT NOT NULL, "
        "kind TEXT NOT NULL, name TEXT NOT NULL, status TEXT NOT NULL, "
        "format TEXT NOT NULL, detail TEXT NOT NULL, source_bytes BLOB, "
        "UNIQUE(source_sha256,kind,name,format))"));
    assert(init.exec("CREATE TABLE _la_import_log ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, source_file TEXT, kind TEXT, "
        "name TEXT, status TEXT, details TEXT)"));

    AccessInspector dialog(QString::fromLocal8Bit(argv[1]), db);
    QPushButton* import = nullptr;
    for (auto* button : dialog.findChildren<QPushButton*>()) {
        if (button->text().contains("Objekt-Metadaten")) import = button;
    }
    assert(import && import->isEnabled());
    bool sawAllObjects = false;
    for (auto* list : dialog.findChildren<QListWidget*>()) {
        if (list->count() == 11) sawAllObjects = true;
    }
    assert(sawAllObjects);

    QTimer respond;
    QObject::connect(&respond, &QTimer::timeout, &dialog, [] {
        for (auto* widget : QApplication::topLevelWidgets()) {
            auto* msg = qobject_cast<QMessageBox*>(widget);
            if (!msg || !msg->isVisible()) continue;
            if (msg->standardButtons() & QMessageBox::Yes)
                msg->button(QMessageBox::Yes)->click();
            else if (msg->standardButtons() & QMessageBox::Ok)
                msg->button(QMessageBox::Ok)->click();
            else msg->accept();
        }
    });
    respond.start(20);
    QTimer::singleShot(0, import, &QPushButton::click);
    QTimer::singleShot(900, &dialog, &QDialog::reject);
    dialog.exec();

    QSqlQuery check(db);
    assert(check.exec("SELECT COUNT(*),COUNT(DISTINCT name) FROM _la_access_artifacts "
                      "WHERE status='metadata_only'"));
    assert(check.next() && check.value(0).toInt() == 11);
    assert(check.value(1).toInt() == 11);
    assert(check.exec("SELECT CAST(source_bytes AS TEXT) FROM _la_access_artifacts "
                      "WHERE name='3_qry_Rezepte'"));
    assert(check.next() && check.value(0).toString().contains("2_tbl_Nährwerte"));
    // Exercise the new 'all tables' path against a real Access 2003 MDB.
    AccessInspector bulk(QString::fromLocal8Bit(argv[1]), db);
    QPushButton* importAll = nullptr;
    for (auto* button : bulk.findChildren<QPushButton*>())
        if (button->text().contains("Alle Tabellen")) importAll = button;
    assert(importAll && importAll->isEnabled());
    QTimer::singleShot(0, importAll, &QPushButton::click);
    QTimer::singleShot(10000, &bulk, &QDialog::reject);
    bulk.exec();
    assert(check.exec("SELECT COUNT(*) FROM sqlite_master "
                      "WHERE type='table' AND name IN ('3_tbl_Rezepte','33_tbl_ZutatenRezept')"));
    assert(check.next() && check.value(0).toInt() == 2);
    std::cout << "Real MDB inspector GUI test passed (11 catalog objects)\n";
}
