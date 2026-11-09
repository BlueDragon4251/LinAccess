// Read-only MDBTools process boundary. No VBA loading and no database writes.
#include <QCoreApplication>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QFile>
#include <QSet>
#include <QCryptographicHash>
#include <bit>
#include <limits>
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
QByteArray digest(const QString& path) {
    QFile file(path); if(!file.open(QIODevice::ReadOnly)) throw 1;
    QCryptographicHash hash(QCryptographicHash::Sha256); if(!hash.addData(&file)) throw 1;
    return hash.result().toHex();
}
void output(const QJsonObject& root,qsizetype limit) {
    const auto bytes=QJsonDocument(root).toJson(QJsonDocument::Compact); if(bytes.size()>limit) throw 1;
    QFile out; if(!out.open(stdout,QIODevice::WriteOnly) || out.write(bytes)!=bytes.size() || !out.flush()) throw 1;
}
}
int main(int argc,char** argv) {
    QCoreApplication app(argc,argv);
    if(argc==2 && app.arguments()[1]=="--version") { printf("LinAccess MDB metadata1; MDBTools%s\n",mdb_get_version()); return 0; }
    const bool singles=argc==4 && app.arguments()[1]=="--single-values";
    const bool doubles=argc==4 && app.arguments()[1]=="--double-values";
    const bool floating=singles || doubles;
    if(argc!=3 && !floating) return 2;
    try {
        const auto args=app.arguments(); const int filePosition=floating ? 2 : 1,tablePosition=floating ? 3 : 2;
        const auto file=QFile::encodeName(args[filePosition]); auto requested=args[tablePosition].toUtf8();
        if(args[filePosition].contains(QChar(0)) || args[tablePosition].contains(QChar(0))) throw 1;
        const auto before=digest(args[filePosition]);
        std::unique_ptr<MdbHandle,decltype(&mdb_close)> mdb(mdb_open(file.constData(),MDB_NOFLAGS),&mdb_close);
        if(!mdb) throw 1;
        std::unique_ptr<MdbTableDef,decltype(&mdb_free_tabledef)> table(mdb_read_table_by_name(mdb.get(),requested.data(),MDB_TABLE),&mdb_free_tabledef);
        if(!table || !mdb_read_columns(table.get()) || !table->columns || table->num_cols!=table->columns->len
            || table->num_cols>32767 || table->num_idxs>4096) throw 1;
        // MDBTools1.0.x returns NULL even after successful index reads. Validate
        // the actual populated array, never use that return value as success.
        mdb_read_indices(table.get());
        if(!table->indices || table->indices->len!=table->num_idxs) throw 1;
        const bool displaySupported=!floating && mdb->f->jet_version==MDB_VER_JET4;
        QJsonArray columns,indexes,singleColumns,doubleColumns,counterColumns,fieldProperties; QVector<const MdbColumn*> singleFields,doubleFields; QSet<QString> columnNames,indexNames;
        for(unsigned i=0;i<table->columns->len;++i) {
            const auto* column=static_cast<MdbColumn*>(g_ptr_array_index(table->columns,i)); if(!column) throw 1;
            const auto name=checkedName(column->name,sizeof(column->name));
            if(columnNames.contains(name.toCaseFolded())) throw 1;
            columnNames.insert(name.toCaseFolded()); columns.append(name);
            if(displaySupported) {QJsonObject properties{{"name",name},{"format_proven",true}};for(const auto* key:{"Caption","Description","Format"}) if(const auto* value=mdb_col_get_prop(column,key)) {const auto size=strnlen(value,16385);if(size>16384) throw 1;const QByteArray bytes(value,size);const auto text=QString::fromUtf8(bytes);if(text.size()>4096 || !text.isValidUtf16() || text.toUtf8()!=bytes || text.contains(QChar(0))) throw 1;properties.insert(QString::fromLatin1(key).toLower(),text);}for(const auto* key:{"Required","AllowZeroLength"}) {
                if(QString::fromLatin1(key)=="AllowZeroLength" && column->col_type!=MDB_TEXT && column->col_type!=MDB_MEMO) continue;
                if(const auto* value=mdb_col_get_prop(column,key)) {const QByteArray bytes(value);if(bytes!="yes" && bytes!="no") throw 1;properties.insert(QString::fromLatin1(key)=="Required" ? "required" : "allow_zero_length",bytes=="yes");}
            }fieldProperties.append(properties);}
            if(column->col_type==MDB_FLOAT) { singleColumns.append(name); singleFields.append(column); }
            if(column->col_type==MDB_DOUBLE) { doubleColumns.append(name); doubleFields.append(column); }
        }
        if(floating) {
            const auto& selected=doubles ? doubleFields : singleFields;
            const int width=doubles ? 8 : 4;
            static_assert(sizeof(float)==4 && std::numeric_limits<float>::is_iec559);
            QJsonArray rows; qsizetype cells=0; mdb_rewind_table(table.get());
            if(!selected.isEmpty()) while(mdb_fetch_row(table.get())) {
                if(selected.size()>1000000-cells) throw 1; cells+=selected.size(); QJsonArray values;
                for(const auto* column:selected) {
                    if(!column->cur_value_len) { values.append(QJsonValue::Null); continue; }
                    if(column->cur_value_len!=width || column->cur_value_start<0 || column->cur_value_start>mdb->fmt->pg_size-width) throw 1;
                    if(doubles) {
                        const auto value=mdb_get_double(mdb->pg_buf,column->cur_value_start);
                        values.append(QString("%1").arg(std::bit_cast<quint64>(value),16,16,QChar('0')));
                    } else {
                        const auto value=mdb_get_single(mdb->pg_buf,column->cur_value_start);
                        values.append(QString("%1").arg(std::bit_cast<quint32>(value),8,16,QChar('0')));
                    }
                }
                rows.append(values);
            }
            if(digest(args[filePosition])!=before) throw 1;
            output({{"version",1},{"source_table",args[tablePosition]},{"source_sha256",QString::fromLatin1(before)},
                {"columns",doubles ? doubleColumns : singleColumns},{"rows",rows}},16*1024*1024);
            return 0;
        }
        // Jet3/4 table-definition state stores the last allocated Long,
        // including IDs no longer represented by rows. Never infer it from MAX.
        bool hasLongCounter=false;
        for(unsigned i=0;i<table->columns->len;++i) {
            const auto* column=static_cast<MdbColumn*>(g_ptr_array_index(table->columns,i));
            if(column->is_uuid_auto) throw 1; // GUID generation is not a Long sequence.
            if(!column->is_long_auto) continue;
            if(hasLongCounter || column->col_type!=MDB_LONGINT
                || (mdb->f->jet_version!=MDB_VER_JET3 && mdb->f->jet_version!=MDB_VER_JET4)
                || mdb_col_get_prop(const_cast<MdbColumn*>(column),"NewValues")
                || mdb_col_get_prop(const_cast<MdbColumn*>(column),"Increment")
                || mdb_col_get_prop(const_cast<MdbColumn*>(column),"Seed")) throw 1;
            hasLongCounter=true;
            if(mdb_read_pg(mdb.get(),table->entry->table_pg)<=0) throw 1;
            const auto last=static_cast<qint32>(mdb_get_int32(mdb->pg_buf,mdb->f->jet_version==MDB_VER_JET4 ? 20 : 16));
            if(last<0) throw 1; // Random/negative counter state remains unsupported.
            counterColumns.append(QJsonObject{{"name",checkedName(column->name,sizeof(column->name))},{"last_value",last}});
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
        QJsonObject root{{"version",1},{"source_table",args[tablePosition]},{"library",QString::fromUtf8(mdb_get_version())},{"columns",columns},{"indexes",indexes},{"single_columns",singleColumns},{"double_columns",doubleColumns},{"counter_columns",counterColumns},{"source_sha256",QString::fromLatin1(before)}};
        if(displaySupported) {root.insert("field_properties",fieldProperties);root.insert("field_properties_format","JET4");}
        if(digest(args[filePosition])!=before) throw 1;
        output(root,1024*1024);
        return 0;
    } catch(...) {
        // Do not copy original names, paths, values or private source into logs.
        fputs("MDB metadata could not be verified\n",stderr); return 3;
    }
}
