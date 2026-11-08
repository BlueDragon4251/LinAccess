#include "native_field_designer.hpp"
#include "dao_database.hpp"
#include "access_autonumber.hpp"
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QHeaderView>
#include <QLabel>
#include <QLineEdit>
#include <QMessageBox>
#include <QPushButton>
#include <QSpinBox>
#include <QSqlError>
#include <QSqlQuery>
#include <QTableWidget>
#include <QVBoxLayout>
namespace linaccess {
void editNativeFields(QWidget* parent,QSqlDatabase database,const QString& tableName) {
    QDialog dialog(parent);dialog.setObjectName("native-field-designer");dialog.setWindowTitle("Tabellenfelder: "+tableName);dialog.resize(780,600);
    auto* layout=new QVBoxLayout(&dialog);
    auto* list=new QTableWidget(0,5,&dialog);list->setObjectName("native-field-list");list->setHorizontalHeaderLabels({"Feld","Datentyp","Größe","Pflichtfeld","Standardwert"});list->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);list->setSelectionBehavior(QAbstractItemView::SelectRows);list->setSelectionMode(QAbstractItemView::SingleSelection);list->setEditTriggers(QAbstractItemView::NoEditTriggers);layout->addWidget(list);
    auto* remove=new QPushButton("Ausgewähltes Feld löschen",&dialog);remove->setObjectName("native-field-delete");layout->addWidget(remove);
    auto* form=new QFormLayout;
    auto* name=new QLineEdit(&dialog);name->setObjectName("native-field-name");name->setMaxLength(64);form->addRow("Neues Feld",name);
    auto* type=new QComboBox(&dialog);type->setObjectName("native-field-type");
    const QMap<int,QString> typeNames{{1,"Ja/Nein"},{2,"Byte"},{3,"Ganzzahl (16 Bit)"},{4,"Lange Ganzzahl (32 Bit)"},{7,"Gleitkommazahl (Double)"},{8,"Datum/Uhrzeit"},{10,"Text"},{11,"Binärdaten"},{12,"Langer Text (Memo)"}};
    for(const int code:{10,12,2,3,4,7,8,1,11}) type->addItem(typeNames.value(code),code);form->addRow("Datentyp",type);
    auto* size=new QSpinBox(&dialog);size->setObjectName("native-field-size");size->setRange(1,255);size->setValue(255);form->addRow("Textgröße",size);
    auto* required=new QCheckBox("Eingabe erforderlich",&dialog);required->setObjectName("native-field-required");form->addRow(QString{},required);
    auto* defaultValue=new QLineEdit(&dialog);defaultValue->setObjectName("native-field-default");defaultValue->setMaxLength(255);defaultValue->setPlaceholderText("Optionales Literal, z. B. 0 oder \"Text\"");form->addRow("Standardwert",defaultValue);layout->addLayout(form);
    auto* append=new QPushButton("Feld anlegen",&dialog);append->setObjectName("native-field-create");layout->addWidget(append);
    auto* status=new QLabel(&dialog);status->setWordWrap(true);layout->addWidget(status);
    auto* close=new QDialogButtonBox(QDialogButtonBox::Close,&dialog);layout->addWidget(close);QObject::connect(close,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    std::shared_ptr<VbaObject> table;
    const auto error=[&](const std::exception& problem) {QMessageBox::warning(&dialog,"Feld konnte nicht geändert werden",QString::fromUtf8(problem.what()));};
    const auto reload=[&] {
        auto quoted=tableName;quoted.replace('"',"\"\"");QSet<QString> protectedFields=generatedAccessColumns(database,tableName);
        QSqlQuery keys(database);if(!keys.exec("PRAGMA main.table_info(\""+quoted+"\")")) throw VbaRuntimeError(3146,"Feldeigenschaften konnten nicht gelesen werden");
        while(keys.next()) if(keys.value(5).toInt()>0) protectedFields.insert(keys.value(1).toString().toCaseFolded());keys.finish();
        struct Entry {QStringList values;bool protectedField;};QVector<Entry> entries;
        for(const auto& value:table->member("Fields",{}).object->enumerate()) {
            const auto field=value.object;const auto fieldName=QString::fromStdString(field->member("Name",{}).asText());
            const auto property=[&](const char* member) {try {return field->member(member,{});}catch(const VbaRuntimeError& problem) {if(problem.number!=3251) throw;return VbaValue::empty();}};
            const auto kind=property("Type"),capacity=property("Size"),mandatory=property("Required"),initial=property("DefaultValue");
            const auto code=kind.kind==VbaValue::Kind::Number ? static_cast<int>(kind.number) : 0;
            QString displayedSize="Nicht verfügbar";
            if(code==11 || code==12) displayedSize="Variabel";
            else if(capacity.kind==VbaValue::Kind::Number) displayedSize=QString::number(static_cast<int>(capacity.number))+(code==10 ? " Zeichen" : " Byte");
            entries.append({{fieldName,typeNames.value(code,"Nicht verfügbar"),displayedSize,mandatory.kind==VbaValue::Kind::Boolean ? (mandatory.boolean ? "ja" : "nein") : "Nicht verfügbar",initial.kind==VbaValue::Kind::Text ? QString::fromStdString(initial.text) : "Nicht verfügbar"},protectedFields.contains(fieldName.toCaseFolded())});
        }
        list->setRowCount(0);for(const auto& entry:entries) {const int row=list->rowCount();list->insertRow(row);for(int column=0;column<entry.values.size();++column) list->setItem(row,column,new QTableWidgetItem(entry.values[column]));list->item(row,0)->setData(Qt::UserRole,entry.protectedField);}
        if(list->rowCount()) list->setCurrentCell(0,0);
    };
    try {table=makeDaoDatabase(database)->member("TableDefs",{VbaValue::string(tableName.toStdString())}).object;reload();}
    catch(const std::exception& problem) {error(problem);return;}
    bool catalogValid=true;
    const auto updateDelete=[&] {const int row=list->currentRow();remove->setEnabled(catalogValid && row>=0 && row<list->rowCount() && list->rowCount()>1 && list->item(row,0) && !list->item(row,0)->data(Qt::UserRole).toBool());};
    QObject::connect(list,&QTableWidget::itemSelectionChanged,&dialog,updateDelete);updateDelete();
    QObject::connect(type,&QComboBox::currentIndexChanged,&dialog,[&] {size->setEnabled(type->currentData().toInt()==10);});
    const auto refreshAfterCommit=[&](const QString& message) {
        status->setText(message);
        try {reload();updateDelete();}
        catch(const std::exception& problem) {catalogValid=false;append->setEnabled(false);remove->setEnabled(false);QMessageBox::warning(&dialog,"Änderung gespeichert",message+"\nDie Feldliste konnte nicht neu geladen werden: "+QString::fromUtf8(problem.what()));}
    };
    QObject::connect(append,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid) return;
        const auto fieldName=name->text().trimmed();
        try {
            const auto code=type->currentData().toInt();auto field=table->member("CreateField",{VbaValue::string(fieldName.toStdString()),VbaValue::integer(static_cast<std::int16_t>(code)),VbaValue::longInteger(code==10 ? size->value() : 0)}).object;
            field->assign("Required",VbaValue::logical(required->isChecked()));if(!defaultValue->text().trimmed().isEmpty()) field->assign("DefaultValue",VbaValue::string(defaultValue->text().trimmed().toStdString()));
            table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});
        } catch(const std::exception& problem) {error(problem);return;}
        refreshAfterCommit("Feld angelegt: "+fieldName);
    });
    QObject::connect(remove,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid) return;
        const int row=list->currentRow();if(row<0 || list->item(row,0)->data(Qt::UserRole).toBool()) return;
        const auto fieldName=list->item(row,0)->text();
        if(QMessageBox::question(&dialog,"Feld löschen","Feld „"+fieldName+"“ und alle darin gespeicherten Werte löschen? Verweise in Abfragen und Formularen werden nicht automatisch angepasst.",QMessageBox::Yes|QMessageBox::No,QMessageBox::No)!=QMessageBox::Yes) return;
        try {table->member("Fields",{}).object->member("Delete",{VbaValue::string(fieldName.toStdString())});}
        catch(const std::exception& problem) {error(problem);return;}
        refreshAfterCommit("Feld gelöscht: "+fieldName);
    });
    dialog.exec();
}
}
