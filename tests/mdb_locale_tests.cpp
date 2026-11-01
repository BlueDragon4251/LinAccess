#include "access_catalog.hpp"
#include "access_table_import.hpp"
#include <QCoreApplication>
#include <QCryptographicHash>
#include <QFile>
#include <QTemporaryDir>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <algorithm>
#include <iostream>

int main(int argc,char** argv) {
    QCoreApplication app(argc,argv); assert(argc==2);
    assert(qgetenv("LC_ALL")=="C" && qgetenv("LANG")=="C");
    const QString original=QString::fromLocal8Bit(argv[1]);
    QFile input(original); assert(input.open(QIODevice::ReadOnly));
    const auto before=QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256); input.close();
    QTemporaryDir work; assert(work.isValid()); const auto copy=work.filePath(QString::fromUtf8("unicode-ä.mdb"));
    assert(QFile::copy(original,copy));
    linaccess::AccessCatalog source(copy);
    QStringList tables;
    for (const auto& object:source.objects()) if (object.kind=="table" && !object.internal) tables.append(object.name);
    assert(tables.size()==29);
    int checked=0;
    for (const auto& table:tables) {
        if (std::none_of(table.begin(),table.end(),[](QChar ch) { return ch.unicode()>127; })) continue;
        const auto schema=linaccess::parseMdbSchema(source.tableSchema(table));
        assert(!schema.columns.isEmpty()); ++checked;
    }
    assert(checked>0);
    assert(qgetenv("LC_ALL")=="C" && qgetenv("LANG")=="C");
    assert(input.open(QIODevice::ReadOnly));
    assert(QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256)==before);
    std::cout<<"MDB UTF-8 arguments under a C-locale parent: "<<checked<<" table schemas, unchanged original hash\n";
}
