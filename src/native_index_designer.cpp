#include "native_index_designer.hpp"
#include "dao_database.hpp"
#include "dao_schema_change.hpp"
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QHeaderView>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QListWidget>
#include <QMessageBox>
#include <QPushButton>
#include <QTableWidget>
#include <QVBoxLayout>
namespace linaccess {
void editNativeIndexes(QWidget* parent,QSqlDatabase database,const QString& tableName) {
    QDialog dialog(parent);dialog.setObjectName("native-index-designer");dialog.setWindowTitle("Indizes: "+tableName);dialog.resize(680,580);
    auto* layout=new QVBoxLayout(&dialog);
    auto* list=new QListWidget(&dialog);list->setObjectName("native-index-list");layout->addWidget(list);
    auto* remove=new QPushButton("Ausgewählten Index löschen",&dialog);remove->setObjectName("native-index-delete");layout->addWidget(remove);
    auto* name=new QLineEdit(&dialog);name->setObjectName("native-index-name");name->setPlaceholderText("Name des neuen Index");name->setMaxLength(64);layout->addWidget(name);
    auto* options=new QHBoxLayout;
    auto* primary=new QCheckBox("Primärschlüssel",&dialog);primary->setObjectName("native-index-primary");options->addWidget(primary);
    auto* unique=new QCheckBox("Eindeutig",&dialog);unique->setObjectName("native-index-unique");options->addWidget(unique);QObject::connect(primary,&QCheckBox::toggled,&dialog,[unique](bool checked) {if(checked) unique->setChecked(true);unique->setEnabled(!checked);});
    auto* required=new QCheckBox("Keine Nullwerte",&dialog);required->setObjectName("native-index-required");options->addWidget(required);layout->addLayout(options);
    auto* grid=new QTableWidget(0,2,&dialog);grid->setObjectName("native-index-fields");grid->setHorizontalHeaderLabels({"Feld","Sortierung"});grid->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);layout->addWidget(grid);
    auto* rowButtons=new QHBoxLayout;
    auto* addField=new QPushButton("+ Indexfeld",&dialog);addField->setObjectName("native-index-add-field");rowButtons->addWidget(addField);
    auto* removeField=new QPushButton("Indexfeld entfernen",&dialog);rowButtons->addWidget(removeField);layout->addLayout(rowButtons);
    auto* append=new QPushButton("Index anlegen",&dialog);append->setObjectName("native-index-create");layout->addWidget(append);
    auto* status=new QLabel(&dialog);status->setWordWrap(true);layout->addWidget(status);
    auto* close=new QDialogButtonBox(QDialogButtonBox::Close,&dialog);layout->addWidget(close);QObject::connect(close,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    std::shared_ptr<VbaObject> table;
    QStringList fields;
    const auto failure=[&](const std::exception& error) {QMessageBox::warning(&dialog,"Index konnte nicht geändert werden",QString::fromUtf8(error.what()));};
    const auto reload=[&] {
        const auto indexes=table->member("Indexes",{}).object;
        struct Entry {QString name,label;bool primary;};QVector<Entry> entries;
        for(const auto& value:indexes->enumerate()) {
            auto index=value.object;const auto indexName=QString::fromStdString(index->member("Name",{}).asText());
            const auto primary=index->member("Primary",{}).truthy(),isUnique=index->member("Unique",{}).truthy();
            QStringList columns;
            for(const auto& field:index->member("Fields",{}).object->enumerate()) columns.append(QString::fromStdString(field.object->member("Name",{}).asText())+(field.object->member("Attributes",{}).asNumber()==1 ? " ↓" : " ↑"));
            entries.append({indexName,indexName+" — "+columns.join(", ")+(primary ? " (Primärschlüssel)" : isUnique ? " (eindeutig)" : ""),primary});
        }
        list->clear();for(const auto& entry:entries) {auto* item=new QListWidgetItem(entry.label,list);item->setData(Qt::UserRole,entry.name);item->setData(Qt::UserRole+1,entry.primary);}
        if(list->count()) list->setCurrentRow(0);
    };
    try {
        table=makeDaoDatabase(database)->member("TableDefs",{VbaValue::string(tableName.toStdString())}).object;
        // Never offer a partial catalog when physical identity cannot be verified.
        reload();
        for(const auto& field:table->member("Fields",{}).object->enumerate()) {
            try {const auto type=field.object->member("Type",{}).asNumber();if(type>0 && type!=11 && type!=12 && type!=15) fields.append(QString::fromStdString(field.object->member("Name",{}).asText()));}
            catch(const VbaRuntimeError& error) {if(error.number!=3251) throw;}
        }
    } catch(const std::exception& error) {failure(error);return;}
    const auto addRow=[&] {
        if(grid->rowCount()>=10 || fields.isEmpty()) return;
        const auto row=grid->rowCount();grid->insertRow(row);auto* field=new QComboBox(grid);field->addItems(fields);grid->setCellWidget(row,0,field);
        auto* direction=new QComboBox(grid);direction->addItems({"Aufsteigend","Absteigend"});grid->setCellWidget(row,1,direction);grid->setCurrentCell(row,0);
    };
    QObject::connect(addField,&QPushButton::clicked,&dialog,addRow);addRow();
    QObject::connect(removeField,&QPushButton::clicked,&dialog,[&] {if(grid->currentRow()>=0) grid->removeRow(grid->currentRow());});
    bool catalogValid=true;
    QObject::connect(list,&QListWidget::currentRowChanged,&dialog,[&] {const auto* item=list->currentItem();remove->setEnabled(catalogValid && item);});
    remove->setEnabled(list->currentItem());
    append->setEnabled(!fields.isEmpty());addField->setEnabled(!fields.isEmpty());
    const auto refreshAfterCommit=[&](const QString& message) {
        status->setText(message);
        try {reload();}
        catch(const std::exception& error) {catalogValid=false;append->setEnabled(false);remove->setEnabled(false);QMessageBox::warning(&dialog,"Änderung gespeichert",message+"\nDie Indexliste konnte nicht neu geladen werden: "+QString::fromUtf8(error.what()));}
    };
    const auto storedFailure=[&](const DaoStoredSchemaChange& error) {catalogValid=false;append->setEnabled(false);remove->setEnabled(false);status->setText("Indexänderung gespeichert. Bitte das Projekt neu öffnen.");QMessageBox::warning(&dialog,"Änderung gespeichert",status->text()+"\n"+QString::fromUtf8(error.what()));};
    QObject::connect(append,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid) return;
        try {
            auto index=table->member("CreateIndex",{VbaValue::string(name->text().toStdString())}).object;
            index->assign("Unique",VbaValue::logical(unique->isChecked()));index->assign("Required",VbaValue::logical(required->isChecked()));index->assign("Primary",VbaValue::logical(primary->isChecked()));
            auto columns=index->member("Fields",{}).object;
            for(int row=0;row<grid->rowCount();++row) {
                const auto fieldName=qobject_cast<QComboBox*>(grid->cellWidget(row,0))->currentText();
                auto field=index->member("CreateField",{VbaValue::string(fieldName.toStdString())}).object;
                field->assign("Attributes",VbaValue::longInteger(qobject_cast<QComboBox*>(grid->cellWidget(row,1))->currentIndex()==1 ? 1 : 0));columns->member("Append",{VbaValue::reference(field)});
            }
            table->member("Indexes",{}).object->member("Append",{VbaValue::reference(index)});
        } catch(const DaoStoredSchemaChange& error) {storedFailure(error);return;}catch(const std::exception& error) {failure(error);return;}
        refreshAfterCommit("Index angelegt: "+name->text());
    });
    QObject::connect(remove,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid) return;
        const auto* item=list->currentItem();if(!item) return;
        const auto indexName=item->data(Qt::UserRole).toString();
        if(QMessageBox::question(&dialog,"Index löschen","Index „"+indexName+"“ löschen? Gespeicherte Datensätze bleiben erhalten.",QMessageBox::Yes|QMessageBox::No,QMessageBox::No)!=QMessageBox::Yes) return;
        try {table->member("Indexes",{}).object->member("Delete",{VbaValue::string(indexName.toStdString())});}
        catch(const DaoStoredSchemaChange& error) {storedFailure(error);return;}catch(const std::exception& error) {failure(error);return;}
        refreshAfterCommit("Index gelöscht: "+indexName);
    });
    dialog.exec();
}
}
