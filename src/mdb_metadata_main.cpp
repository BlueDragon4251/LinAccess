// Read-only MDBTools process boundary. No VBA loading and no database writes.
#include <QCoreApplication>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QFile>
#include <QSet>
#include <memory>
#include <cstring>
#include <mdbtools.h>

namespace {
QString checkedName(const char* name,std::size_t bound) {
    const auto size=strnlen(name,bound);
    if(!size || size==bound) throw 1;
    const QByteArray bytes(name,static_cast<qsizetype>(size));
    const auto result=QString::fromUtf8(bytes);
    if(!result.isValidUtf16() || result.toUtf8()!=bytes || result.contains(QChar(0))) throw 1;
    return result;
}
}
int main(int argc,char** argv) {
    QCoreApplication app(argc,argv);
    if(argc==2 && app.arguments()[1]=="--version") { printf("LinAccess MDB metadata1; MDBTools%s\n",mdb_get_version()); return 0; }
    if(argc!=3) return 2;
    try {
        const auto args=app.arguments(); const auto file=QFile::encodeName(args[1]); auto requested=args[2].toUtf8();
        if(args[1].contains(QChar(0)) || args[2].contains(QChar(0))) throw 1;
        std::unique_ptr<MdbHandle,decltype(&mdb_close)> mdb(mdb_open(file.constData(),MDB_NOFLAGS),&mdb_close);
        if(!mdb) throw 1;
        std::unique_ptr<MdbTableDef,decltype(&mdb_free_tabledef)> table(mdb_read_table_by_name(mdb.get(),requested.data(),MDB_TABLE),&mdb_free_tabledef);
        if(!table || !mdb_read_columns(table.get()) || !table->columns || table->num_cols!=table->columns->len
            || table->num_cols>32767 || table->num_idxs>4096) throw 1;
        // MDBTools1.0.x returns NULL even after successful index reads. Validate
        // the actual populated array, never use that return value as success.
        mdb_read_indices(table.get());
        if(!table->indices || table->indices->len!=table->num_idxs) throw 1;
        QJsonArray columns,indexes; QSet<QString> columnNames,indexNames;
        for(unsigned i=0;i<table->columns->len;++i) {
            const auto* column=static_cast<MdbColumn*>(g_ptr_array_index(table->columns,i)); if(!column) throw 1;
            const auto name=checkedName(column->name,sizeof(column->name));
            if(columnNames.contains(name.toCaseFolded())) throw 1;
            columnNames.insert(name.toCaseFolded()); columns.append(name);
        }
        for(unsigned i=0;i<table->indices->len;++i) {
            const auto* index=static_cast<MdbIndex*>(g_ptr_array_index(table->indices,i)); if(!index) throw 1;
            const auto name=checkedName(index->name,sizeof(index->name));
            if(indexNames.contains(name.toCaseFolded()) || index->index_type>2 || index->num_keys>MDB_MAX_IDX_COLS
                || (!index->num_keys && index->index_type!=2)) throw 1;
            indexNames.insert(name.toCaseFolded()); QJsonArray fields; QSet<QString> names;
            for(unsigned j=0;j<index->num_keys;++j) {
                const auto ordinal=index->key_col_num[j];
                if(ordinal<1 || ordinal>static_cast<int>(table->columns->len) || index->key_col_order[j]>1) throw 1;
                const auto field=columns[ordinal-1].toString(); if(names.contains(field.toCaseFolded())) throw 1;
                names.insert(field.toCaseFolded()); fields.append(QJsonObject{{"name",field},{"descending",index->key_col_order[j]!=0}});
            }
            indexes.append(QJsonObject{{"name",name},{"kind",index->index_type},{"flags",index->flags},{"fields",fields}});
        }
        const QJsonObject root{{"version",1},{"source_table",args[2]},{"library",QString::fromUtf8(mdb_get_version())},{"columns",columns},{"indexes",indexes}};
        const auto output=QJsonDocument(root).toJson(QJsonDocument::Compact);
        if(output.size()>1024*1024) throw 1;
        QFile out; if(!out.open(stdout,QIODevice::WriteOnly) || out.write(output)!=output.size() || !out.flush()) throw 1;
        return 0;
    } catch(...) {
        // Do not copy original names, paths, values or private source into logs.
        fputs("MDB index metadata could not be verified\n",stderr); return 3;
    }
}
