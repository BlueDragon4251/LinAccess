#include "mainwindow.hpp"
#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "vba_runtime.hpp"

#include <QApplication>
#include <QSqlQuery>
#include <sqlite3.h>
#include <iostream>

namespace {
int checkSqlBackend(int argc,char** argv) {
    QCoreApplication application(argc,argv);
    auto db=QSqlDatabase::addDatabase("QSQLITE"); db.setDatabaseName(":memory:");
    if (!db.open()) return 1;
    try {
        linaccess::installJetSqlFunctions(db);
        QSqlQuery query(db);
        for (const auto* sql : {"CREATE TABLE a(id INTEGER)","CREATE TABLE b(id INTEGER)","CREATE TABLE c(id INTEGER)","CREATE TABLE d(id INTEGER)",
             "INSERT INTO a VALUES(1),(2)","INSERT INTO b VALUES(1),(3)","INSERT INTO c VALUES(1),(4)","INSERT INTO d VALUES(1),(5)",
             "CREATE TEMP VIEW backend_nested AS SELECT a.id AS ai,b.id AS bi,c.id AS ci,d.id AS di FROM a RIGHT JOIN (b RIGHT JOIN (c LEFT JOIN d ON c.id=d.id) ON b.id=c.id) ON a.id=b.id"})
            if (!query.exec(QString::fromLatin1(sql))) return 1;
        linaccess::AccessQueryModel model;
        model.execute(db,"SELECT * FROM backend_nested ORDER BY ci");
        if (model.rowCount()!=2 || model.columnCount()!=4 || model.data(model.index(0,0)).toInt()!=1
            || model.data(model.index(1,2)).toInt()!=4 || !model.data(model.index(1,0)).isNull()
            || !model.data(model.index(1,1)).isNull() || !model.data(model.index(1,3)).isNull()) return 1;
        std::cout<<"SQLite backend verified: "<<SQLITE_VERSION<<"; nested JOIN/NULL regression passed\n";
        return 0;
    } catch (const linaccess::VbaRuntimeError& error) {
        std::cerr<<"SQLite backend check failed: error "<<error.number<<'\n'; return 1;
    } catch (const std::exception&) { return 1; }
}
}

int main(int argc, char* argv[]) {
    if (argc==2 && QByteArray(argv[1])=="--check-sql-backend") return checkSqlBackend(argc,argv);
    QApplication application(argc, argv);
    application.setApplicationName("LinAccess");
    application.setOrganizationName("LinAccess Community");
    MainWindow window;
    window.show();
    return application.exec();
}
