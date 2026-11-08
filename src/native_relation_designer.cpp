#include "native_relation_designer.hpp"
#include "dao_database.hpp"
#include "dao_schema_change.hpp"
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QHeaderView>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QMessageBox>
#include <QPushButton>
#include <QTableWidget>
#include <QVBoxLayout>
namespace linaccess {
void editNativeRelations(QWidget* parent,QSqlDatabase database,const QString& suggestedParent) {
    QDialog dialog(parent);dialog.setObjectName("native-relation-designer");dialog.setWindowTitle("Beziehungen");dialog.resize(1000,700);
    auto* layout=new QVBoxLayout(&dialog);
    auto* list=new QTableWidget(0,7,&dialog);list->setObjectName("native-relation-list");list->setHorizontalHeaderLabels({"Name","Elterntabelle","Kindtabelle","Feldzuordnung","Integrität","Update-Kaskade","Löschkaskade"});list->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);list->setSelectionBehavior(QAbstractItemView::SelectRows);list->setSelectionMode(QAbstractItemView::SingleSelection);list->setEditTriggers(QAbstractItemView::NoEditTriggers);layout->addWidget(list);
    auto* remove=new QPushButton("Ausgewählte Beziehung entfernen",&dialog);remove->setObjectName("native-relation-delete");layout->addWidget(remove);
    auto* form=new QFormLayout;
    auto* name=new QLineEdit(&dialog);name->setObjectName("native-relation-name");name->setMaxLength(64);form->addRow("Neue Beziehung",name);
    auto* source=new QComboBox(&dialog);source->setObjectName("native-relation-parent");form->addRow("Elterntabelle",source);
    auto* target=new QComboBox(&dialog);target->setObjectName("native-relation-child");form->addRow("Kindtabelle",target);layout->addLayout(form);
    auto* integrity=new QCheckBox("Referenzielle Integrität (numerische Schlüssel)",&dialog);integrity->setObjectName("native-relation-integrity");integrity->setChecked(true);layout->addWidget(integrity);
    auto* options=new QHBoxLayout;
    auto* update=new QCheckBox("Aktualisierungen weitergeben",&dialog);update->setObjectName("native-relation-update-cascade");options->addWidget(update);
    auto* erase=new QCheckBox("Löschungen weitergeben",&dialog);erase->setObjectName("native-relation-delete-cascade");options->addWidget(erase);layout->addLayout(options);
    auto* fields=new QTableWidget(0,2,&dialog);fields->setObjectName("native-relation-fields");fields->setHorizontalHeaderLabels({"Elternfeld","Kindfeld"});fields->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);layout->addWidget(fields);
    auto* rowButtons=new QHBoxLayout;
    auto* addField=new QPushButton("+ Feldzuordnung",&dialog);addField->setObjectName("native-relation-add-field");rowButtons->addWidget(addField);
    auto* removeField=new QPushButton("Feldzuordnung entfernen",&dialog);rowButtons->addWidget(removeField);layout->addLayout(rowButtons);
    auto* append=new QPushButton("Beziehung anlegen",&dialog);append->setObjectName("native-relation-create");layout->addWidget(append);
    layout->addWidget(new QLabel("Integrität und Kaskaden benötigen numerische Schlüssel und eine in LinAccess angelegte Kindtabelle.",&dialog));
    auto* status=new QLabel(&dialog);status->setWordWrap(true);layout->addWidget(status);
    auto* close=new QDialogButtonBox(QDialogButtonBox::Close,&dialog);layout->addWidget(close);QObject::connect(close,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    const auto db=makeDaoDatabase(database);bool catalogValid=true;
    const auto failure=[&](const std::exception& problem) {QMessageBox::warning(&dialog,"Beziehung konnte nicht geändert werden",QString::fromUtf8(problem.what()));};
    const auto disable=[&] {catalogValid=false;append->setEnabled(false);remove->setEnabled(false);list->setEnabled(false);};
    const auto reload=[&] {
        QVector<QStringList> entries;
        for(const auto& value:db->member("Relations",{}).object->enumerate()) {
            if(entries.size()>=4096) throw VbaRuntimeError(3251,"Beziehungsansicht überschreitet die unterstützte Größe");
            const auto relation=value.object;const int attributes=static_cast<int>(relation->member("Attributes",{}).asNumber());QStringList mapping;
            for(const auto& field:relation->member("Fields",{}).object->enumerate()) mapping.append(QString::fromStdString(field.object->member("Name",{}).asText())+" → "+QString::fromStdString(field.object->member("ForeignName",{}).asText()));
            const bool enforced=!(attributes&2);
            entries.append({QString::fromStdString(relation->member("Name",{}).asText()),QString::fromStdString(relation->member("Table",{}).asText()),QString::fromStdString(relation->member("ForeignTable",{}).asText()),mapping.join(", "),enforced ? "ja" : "nein",enforced && (attributes&256) ? "ja" : "nein",enforced && (attributes&4096) ? "ja" : "nein"});
        }
        list->setRowCount(0);for(const auto& entry:entries) {const int row=list->rowCount();list->insertRow(row);for(int column=0;column<entry.size();++column) list->setItem(row,column,new QTableWidgetItem(entry[column]));}
        if(list->rowCount()) list->setCurrentCell(0,0);
    };
    try {
        QStringList tables;for(const auto& table:db->member("TableDefs",{}).object->enumerate()) tables.append(QString::fromStdString(table.object->member("Name",{}).asText()));source->addItems(tables);target->addItems(tables);
        if(source->findText(suggestedParent)>=0) source->setCurrentText(suggestedParent);for(int i=0;i<target->count();++i) if(target->itemText(i)!=source->currentText()) {target->setCurrentIndex(i);break;}
        reload();
    } catch(const std::exception& problem) {failure(problem);return;}
    if(source->count()==0) status->setText("Legen Sie zuerst eine Tabelle an.");
    const auto columns=[&](const QString& tableName) {
        QStringList result;if(tableName.isEmpty()) return result;
        const auto table=db->member("TableDefs",{VbaValue::string(tableName.toStdString())}).object;
        for(const auto& field:table->member("Fields",{}).object->enumerate()) {
            try {const auto type=field.object->member("Type",{}).asNumber();const bool numeric=type==1 || type==2 || type==3 || type==4 || type==7;
                if(numeric || (!integrity->isChecked() && (type==5 || type==6 || type==8 || type==10))) result.append(QString::fromStdString(field.object->member("Name",{}).asText()));}
            catch(const VbaRuntimeError& problem) {if(problem.number!=3251) throw;}
        }
        return result;
    };
    const auto updateDelete=[&] {remove->setEnabled(catalogValid && list->currentRow()>=0 && list->currentRow()<list->rowCount());};QObject::connect(list,&QTableWidget::itemSelectionChanged,&dialog,updateDelete);updateDelete();
    const auto refill=[&] {
        try {const QStringList names[]{columns(source->currentText()),columns(target->currentText())};
            for(int row=0;row<fields->rowCount();++row) for(int column=0;column<2;++column) {auto* combo=qobject_cast<QComboBox*>(fields->cellWidget(row,column));const auto previous=combo->currentText();combo->clear();combo->addItems(names[column]);if(combo->findText(previous)>=0) combo->setCurrentText(previous);}
            const bool available=!names[0].isEmpty() && !names[1].isEmpty();append->setEnabled(catalogValid && available);addField->setEnabled(available);
        } catch(const std::exception& problem) {disable();failure(problem);}
    };
    const auto addRow=[&] {if(fields->rowCount()>=10) return;const int row=fields->rowCount();fields->insertRow(row);for(int column=0;column<2;++column) fields->setCellWidget(row,column,new QComboBox(fields));fields->setCurrentCell(row,0);refill();};
    QObject::connect(addField,&QPushButton::clicked,&dialog,addRow);QObject::connect(removeField,&QPushButton::clicked,&dialog,[&] {if(fields->currentRow()>=0) fields->removeRow(fields->currentRow());});
    QObject::connect(source,&QComboBox::currentTextChanged,&dialog,[&] {refill();});QObject::connect(target,&QComboBox::currentTextChanged,&dialog,[&] {refill();});
    QObject::connect(integrity,&QCheckBox::toggled,&dialog,[&](bool enabled) {update->setEnabled(enabled);erase->setEnabled(enabled);if(!enabled) {update->setChecked(false);erase->setChecked(false);}refill();});addRow();
    const auto storedFailure=[&](const DaoStoredSchemaChange& problem) {disable();status->setText("Änderung gespeichert. Bitte das Projekt neu öffnen.");QMessageBox::warning(&dialog,"Änderung gespeichert",status->text()+"\n"+QString::fromUtf8(problem.what()));};
    const auto refreshAfterCommit=[&](const QString& message) {status->setText(message);try {reload();updateDelete();}catch(const std::exception& problem) {disable();QMessageBox::warning(&dialog,"Änderung gespeichert",message+"\nDie Beziehungsliste konnte nicht neu geladen werden: "+QString::fromUtf8(problem.what()));}};
    QObject::connect(append,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid) return;const auto relationName=name->text().trimmed();
        try {
            const int attributes=integrity->isChecked() ? ((update->isChecked() ? 256 : 0)|(erase->isChecked() ? 4096 : 0)) : 2;
            auto relation=db->member("CreateRelation",{VbaValue::string(relationName.toStdString()),VbaValue::string(source->currentText().toStdString()),VbaValue::string(target->currentText().toStdString()),VbaValue::longInteger(attributes)}).object;
            for(int row=0;row<fields->rowCount();++row) {auto field=relation->member("CreateField",{VbaValue::string(qobject_cast<QComboBox*>(fields->cellWidget(row,0))->currentText().toStdString())}).object;field->assign("ForeignName",VbaValue::string(qobject_cast<QComboBox*>(fields->cellWidget(row,1))->currentText().toStdString()));relation->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});}
            db->member("Relations",{}).object->member("Append",{VbaValue::reference(relation)});
        } catch(const DaoStoredSchemaChange& problem) {storedFailure(problem);return;}catch(const std::exception& problem) {failure(problem);return;}
        refreshAfterCommit("Beziehung angelegt: "+relationName);
    });
    QObject::connect(remove,&QPushButton::clicked,&dialog,[&] {
        const int row=list->currentRow();if(!catalogValid || row<0 || !list->item(row,0)) return;const auto relationName=list->item(row,0)->text();
        if(QMessageBox::question(&dialog,"Beziehung entfernen","Beziehung „"+relationName+"“ entfernen? Die gespeicherten Datensätze bleiben erhalten.",QMessageBox::Yes|QMessageBox::No,QMessageBox::No)!=QMessageBox::Yes) return;
        try {db->member("Relations",{}).object->member("Delete",{VbaValue::string(relationName.toStdString())});}
        catch(const DaoStoredSchemaChange& problem) {storedFailure(problem);return;}catch(const std::exception& problem) {failure(problem);return;}
        refreshAfterCommit("Beziehung entfernt: "+relationName);
    });
    dialog.exec();
}
}
