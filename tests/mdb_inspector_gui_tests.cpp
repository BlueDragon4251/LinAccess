#include "access_inspector.hpp"

#include <QApplication>
#include <QListWidget>
#include <QMessageBox>
#include <QPushButton>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QTemporaryDir>
#include <QTimer>
#include <QTableWidget>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QLineEdit>
#include <QDir>
#include <QCryptographicHash>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

namespace {

QByteArray readFile(const QString& path) {
    QFile file(path);
    assert(file.open(QIODevice::ReadOnly));
    return file.readAll();
}

void verifyOlePackageExtraction(const QString& reference, const QString& outputDirectory,
                                QSqlDatabase db) {
    const auto sourceHash = QCryptographicHash::hash(readFile(reference), QCryptographicHash::Sha256);
    AccessInspector inspector(reference, db);
    inspector.show();
    QApplication::processEvents();

    auto* grid = inspector.findChild<QTableWidget*>();
    QPushButton* preview = nullptr;
    QPushButton* extract = nullptr;
    for (auto* button : inspector.findChildren<QPushButton*>()) {
        if (button->text() == "Tabelle ansehen") preview = button;
        if (button->text().startsWith("OLE-Paket speichern")) extract = button;
    }
    assert(grid && preview && extract && !extract->isEnabled());
    for (auto* list : inspector.findChildren<QListWidget*>()) {
        const auto matches = list->findItems("ole_package_probe", Qt::MatchExactly);
        if (!matches.isEmpty()) list->setCurrentItem(matches.front());
    }
    preview->click();
    assert(grid->rowCount() == 3 && grid->columnCount() == 2);

    const auto chooseDestination = [&](const QString& destination, bool expectFailure = false) {
        bool sawDialog = false;
        bool sawWarning = false;
        QTimer::singleShot(0, &inspector, [&] {
            auto* picker = inspector.findChild<QFileDialog*>();
            assert(picker && picker->isVisible());
            const auto* nameEdit = picker->findChild<QLineEdit*>("fileNameEdit");
            assert(nameEdit && nameEdit->text().isEmpty());
            sawDialog = true;
            if (destination.isEmpty()) {
                picker->reject();
                return;
            }
            const QFileInfo target(destination);
            picker->setDirectory(target.absolutePath());
            picker->selectFile(target.fileName());
            static_cast<QDialog*>(picker)->accept();
            if (expectFailure) {
                QTimer::singleShot(0, &inspector, [&] {
                    auto* warning = inspector.findChild<QMessageBox*>();
                    if (warning && warning->isVisible()) {
                        sawWarning = warning->icon() == QMessageBox::Warning;
                        warning->button(QMessageBox::Ok)->click();
                    }
                });
            }
        });
        extract->click();
        assert(sawDialog && sawWarning == expectFailure);
    };

    grid->setCurrentCell(0, 0);
    assert(!extract->isEnabled());
    const QDir fixtureDirectory = QFileInfo(reference).dir();
    const QDir output(outputDirectory);
    int supportedRow = -1;
    int linkedRow = -1;
    for (int row = 0; row < grid->rowCount(); ++row) {
        const int id = grid->item(row, 0)->text().toInt();
        assert(id >= 1 && id <= 3);
        grid->setCurrentCell(row, 1);
        assert(extract->isEnabled() == (id != 3));
        if (id == 3) {
            linkedRow = row;
            continue;
        }
        supportedRow = row;
        const auto destination = output.filePath(QString("chosen-package-%1.bin").arg(id));
        chooseDestination({});
        assert(!QFileInfo::exists(destination));
        chooseDestination(destination);
        assert(readFile(destination) == readFile(fixtureDirectory.filePath(
            QString("ole-content-%1.bin").arg(id))));
        grid->clearSelection();
        assert(!extract->isEnabled());
    }

    assert(supportedRow >= 0 && linkedRow >= 0);
    grid->setCurrentCell(supportedRow, 1);
    assert(extract->isEnabled());
    chooseDestination(reference, true);
    assert(QCryptographicHash::hash(readFile(reference), QCryptographicHash::Sha256) == sourceHash);
    grid->setCurrentCell(linkedRow, 1);
    assert(!extract->isEnabled());
    grid->setCurrentCell(supportedRow, 1);
    assert(extract->isEnabled());
    preview->click();
    assert(!extract->isEnabled());
    std::cout << "Jet4 OLE inspector verified selection, cancel, exact bytes and overwrite refusal\n";
}

}

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
    check.finish();
    const auto reference=qEnvironmentVariable("LINACCESS_JET_BINARY_REFERENCE");
    if(!reference.isEmpty()) {
        const auto digest=[&] { QFile source(reference); assert(source.open(QIODevice::ReadOnly)); return QCryptographicHash::hash(source.readAll(),QCryptographicHash::Sha256); };
        const auto before=digest(); AccessInspector binary(reference,db); binary.show(); QApplication::processEvents();
        for(auto* list:binary.findChildren<QListWidget*>()) {
            const auto matches=list->findItems("binary_probe",Qt::MatchExactly); if(!matches.isEmpty()) list->setCurrentItem(matches.front());
        }
        QPushButton *preview=nullptr,*copy=nullptr,*extract=nullptr;
        for(auto* button:binary.findChildren<QPushButton*>()) {
            if(button->text()=="Tabelle ansehen") preview=button;
            if(button->text().contains("Typisierte Kopie")) copy=button;
            if(button->text().startsWith("OLE-Paket speichern")) extract=button;
        }
        assert(preview && copy && extract); preview->click();
        auto* grid=binary.findChild<QTableWidget*>(); assert(grid && grid->rowCount()==4 && grid->columnCount()==2);
        const QMap<int,int> sizes{{1,7},{3,4097},{4,65539}};
        for(int row=0;row<grid->rowCount();++row) {
            const int id=grid->item(row,0)->text().toInt();
            assert(grid->item(row,1)->text()==(id==2 ? QString("<NULL>") : QString("<Binärdaten: %1 Bytes>").arg(sizes.value(id))));
            grid->setCurrentCell(row, 1);
            assert(!extract->isEnabled());
        }
        assert(copy->isEnabled()); copy->click();
        assert(check.exec("SELECT id,payload,typeof(payload) FROM binary_probe ORDER BY id")); int copied=0;
        while(check.next()) {
            const int id=check.value(0).toInt();
            if(id==2) assert(check.value(1).isNull() && check.value(2).toString()=="null");
            else {
                QFile expected(QFileInfo(reference).dir().filePath("payload-"+QString::number(id)+".bin")); assert(expected.open(QIODevice::ReadOnly));
                assert(check.value(2).toString()=="blob" && check.value(1).toByteArray()==expected.readAll());
            }
            ++copied;
        }
        check.finish(); assert(copied==4 && digest()==before); binary.close();
    }
    const auto oleReference = qEnvironmentVariable("LINACCESS_JET_OLE_PACKAGE_REFERENCE");
    if (!oleReference.isEmpty()) verifyOlePackageExtraction(oleReference, temp.path(), db);
    std::cout << "Real MDB inspector GUI test passed (11 catalog objects)\n";
}
