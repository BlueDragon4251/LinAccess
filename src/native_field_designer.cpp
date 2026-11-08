#include "native_field_designer.hpp"
#include "dao_database.hpp"
#include "dao_schema_change.hpp"
#include "native_schema_editor.hpp"
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
#include <QSqlQuery>
#include <QTabWidget>
#include <QTableWidget>
#include <QVBoxLayout>
namespace linaccess {
void editNativeFields(QWidget* parent,QSqlDatabase database,const QString& tableName) {
    QDialog dialog(parent);dialog.setObjectName("native-field-designer");dialog.setWindowTitle("Tabellenfelder: "+tableName);dialog.resize(800,650);
    auto* layout=new QVBoxLayout(&dialog);
    auto* list=new QTableWidget(0,5,&dialog);list->setObjectName("native-field-list");list->setHorizontalHeaderLabels({"Feld","Datentyp","Größe","Pflichtfeld","Standardwert"});list->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);list->setSelectionBehavior(QAbstractItemView::SelectRows);list->setSelectionMode(QAbstractItemView::SingleSelection);list->setEditTriggers(QAbstractItemView::NoEditTriggers);layout->addWidget(list);
    auto* remove=new QPushButton("Ausgewähltes Feld löschen",&dialog);remove->setObjectName("native-field-delete");layout->addWidget(remove);
    auto* tabs=new QTabWidget(&dialog);tabs->setObjectName("native-field-editor-tabs");layout->addWidget(tabs);
    auto* createPage=new QWidget(tabs);auto* createLayout=new QVBoxLayout(createPage);auto* form=new QFormLayout;
    auto* name=new QLineEdit(createPage);name->setObjectName("native-field-name");name->setMaxLength(64);form->addRow("Neues Feld",name);
    auto* type=new QComboBox(createPage);type->setObjectName("native-field-type");
    const QMap<int,QString> typeNames{{1,"Ja/Nein"},{2,"Byte"},{3,"Ganzzahl (16 Bit)"},{4,"Lange Ganzzahl (32 Bit)"},{6,"Gleitkommazahl (Single)"},{7,"Gleitkommazahl (Double)"},{8,"Datum/Uhrzeit"},{10,"Text"},{11,"Binärdaten"},{12,"Langer Text (Memo)"}};
    for(const int code:{10,12,2,3,4,6,7,8,1,11}) type->addItem(typeNames.value(code),code);form->addRow("Datentyp",type);
    auto* size=new QSpinBox(createPage);size->setObjectName("native-field-size");size->setRange(1,255);size->setValue(255);form->addRow("Textgröße",size);
    auto* required=new QCheckBox("Eingabe erforderlich",createPage);required->setObjectName("native-field-required");form->addRow(QString{},required);
    auto* allowZero=new QCheckBox("Leere Zeichenfolge zulassen",createPage);allowZero->setObjectName("native-field-allow-zero");allowZero->setChecked(true);form->addRow(QString{},allowZero);
    auto* defaultValue=new QLineEdit(createPage);defaultValue->setObjectName("native-field-default");defaultValue->setMaxLength(255);defaultValue->setPlaceholderText("Literal; Datum auch Date(), Now() oder Time()");form->addRow("Standardwert",defaultValue);createLayout->addLayout(form);
    auto* append=new QPushButton("Feld anlegen",createPage);append->setObjectName("native-field-create");createLayout->addWidget(append);tabs->addTab(createPage,"Neues Feld");
    auto* propertyPage=new QWidget(tabs);auto* propertyLayout=new QVBoxLayout(propertyPage);auto* propertyForm=new QFormLayout;
    auto* selectedName=new QLabel(propertyPage);propertyForm->addRow("Ausgewähltes Feld",selectedName);
    auto* propertySize=new QSpinBox(propertyPage);propertySize->setObjectName("native-field-property-size");propertySize->setRange(1,255);propertyForm->addRow("Textgröße",propertySize);
    auto* propertyRequired=new QCheckBox("Eingabe erforderlich",propertyPage);propertyRequired->setObjectName("native-field-property-required");propertyForm->addRow(QString{},propertyRequired);
    auto* propertyZero=new QCheckBox("Leere Zeichenfolge zulassen",propertyPage);propertyZero->setObjectName("native-field-property-allow-zero");propertyForm->addRow(QString{},propertyZero);
    auto* propertyDefault=new QLineEdit(propertyPage);propertyDefault->setObjectName("native-field-property-default");propertyDefault->setMaxLength(255);propertyForm->addRow("Standardwert",propertyDefault);propertyLayout->addLayout(propertyForm);
    auto* renameName=new QLineEdit(propertyPage);renameName->setObjectName("native-field-rename-name");renameName->setMaxLength(64);propertyForm->addRow("Neuer Feldname",renameName);
    auto* renameButton=new QPushButton("Feld umbenennen",propertyPage);renameButton->setObjectName("native-field-rename");propertyLayout->addWidget(renameButton);
    propertyLayout->addWidget(new QLabel("Datum/Uhrzeit: auch Date(), Now() oder Time() als Standardwert.",propertyPage));
    propertyLayout->addWidget(new QLabel("Neue Standardwerte gelten für neue Datensätze; vorhandene Werte werden nicht ersetzt.",propertyPage));
    auto* apply=new QPushButton("Eigenschaften speichern",propertyPage);apply->setObjectName("native-field-property-save");propertyLayout->addWidget(apply);tabs->addTab(propertyPage,"Eigenschaften");
    auto* status=new QLabel(&dialog);status->setWordWrap(true);layout->addWidget(status);
    auto* close=new QDialogButtonBox(QDialogButtonBox::Close,&dialog);layout->addWidget(close);QObject::connect(close,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    std::shared_ptr<VbaObject> table,selectedField;bool catalogValid=true,allNative=false;int selectedType=0;
    const auto error=[&](const std::exception& problem) {QMessageBox::warning(&dialog,"Feld konnte nicht geändert werden",QString::fromUtf8(problem.what()));};
    const auto disable=[&] {catalogValid=false;append->setEnabled(false);remove->setEnabled(false);apply->setEnabled(false);renameButton->setEnabled(false);list->setEnabled(false);};
    const auto property=[](const std::shared_ptr<VbaObject>& field,const char* member) {try {return field->member(member,{});}catch(const VbaRuntimeError& problem) {if(problem.number!=3251) throw;return VbaValue::empty();}};
    const auto reload=[&] {
        const auto previous=list->currentRow()>=0 && list->item(list->currentRow(),0) ? list->item(list->currentRow(),0)->text() : QString{};
        auto quoted=tableName;quoted.replace('"',"\"\"");QSet<QString> protectedFields=generatedAccessColumns(database,tableName);
        QSqlQuery keys(database);if(!keys.exec("PRAGMA main.table_info(\""+quoted+"\")")) throw VbaRuntimeError(3146,"Feldeigenschaften konnten nicht gelesen werden");
        while(keys.next()) if(keys.value(5).toInt()>0) protectedFields.insert(keys.value(1).toString().toCaseFolded());keys.finish();
        struct Entry {QStringList values;bool protectedField;};QVector<Entry> entries;bool native=true;
        for(const auto& value:table->member("Fields",{}).object->enumerate()) {
            const auto field=value.object;const auto fieldName=QString::fromStdString(field->member("Name",{}).asText());
            const auto kind=property(field,"Type"),capacity=property(field,"Size"),mandatory=property(field,"Required"),initial=property(field,"DefaultValue"),attributes=property(field,"Attributes");
            native=native && mandatory.kind==VbaValue::Kind::Boolean && attributes.kind==VbaValue::Kind::Number;
            const auto code=kind.kind==VbaValue::Kind::Number ? static_cast<int>(kind.number) : 0;QString displayedSize="Nicht verfügbar";
            if(code==11 || code==12) displayedSize="Variabel";else if(capacity.kind==VbaValue::Kind::Number) displayedSize=QString::number(static_cast<int>(capacity.number))+(code==10 ? " Zeichen" : " Byte");
            entries.append({{fieldName,typeNames.value(code,"Nicht verfügbar"),displayedSize,mandatory.kind==VbaValue::Kind::Boolean ? (mandatory.boolean ? "ja" : "nein") : "Nicht verfügbar",initial.kind==VbaValue::Kind::Text ? QString::fromStdString(initial.text) : "Nicht verfügbar"},protectedFields.contains(fieldName.toCaseFolded())});
        }
        allNative=native;list->setRowCount(0);int selection=0;
        for(const auto& entry:entries) {const int row=list->rowCount();list->insertRow(row);for(int column=0;column<entry.values.size();++column) list->setItem(row,column,new QTableWidgetItem(entry.values[column]));list->item(row,0)->setData(Qt::UserRole,entry.protectedField);if(entry.values[0]==previous) selection=row;}
        if(list->rowCount()) list->setCurrentCell(selection,0);
    };
    const auto updateDelete=[&] {const int row=list->currentRow();remove->setEnabled(catalogValid && row>=0 && row<list->rowCount() && list->rowCount()>1 && list->item(row,0) && !list->item(row,0)->data(Qt::UserRole).toBool());};
    const auto loadProperties=[&] {
        selectedField.reset();apply->setEnabled(false);renameButton->setEnabled(false);propertySize->setEnabled(false);propertyRequired->setEnabled(false);propertyZero->setEnabled(false);propertyDefault->setEnabled(false);const int row=list->currentRow();
        if(!catalogValid || row<0 || !list->item(row,0)) {selectedName->clear();return;}
        selectedName->setText(list->item(row,0)->text());renameName->setText(selectedName->text());
        try {
            auto field=table->member("Fields",{VbaValue::string(selectedName->text().toStdString())}).object;
            const auto kind=property(field,"Type"),capacity=property(field,"Size"),mandatory=property(field,"Required"),initial=property(field,"DefaultValue"),attributes=property(field,"Attributes");
            selectedType=kind.kind==VbaValue::Kind::Number ? static_cast<int>(kind.number) : 0;
            const bool eligible=allNative && mandatory.kind==VbaValue::Kind::Boolean && attributes.kind==VbaValue::Kind::Number && !(static_cast<int>(attributes.number)&16) && (selectedType!=10 || capacity.kind==VbaValue::Kind::Number);
            propertyRequired->setChecked(mandatory.kind==VbaValue::Kind::Boolean && mandatory.boolean);propertyRequired->setEnabled(eligible);
            if(selectedType==10 && capacity.kind==VbaValue::Kind::Number) propertySize->setValue(static_cast<int>(capacity.number));propertySize->setEnabled(eligible && selectedType==10);
            propertyDefault->setText(initial.kind==VbaValue::Kind::Text ? QString::fromStdString(initial.text) : QString{});propertyDefault->setEnabled(eligible && initial.kind==VbaValue::Kind::Text);
            const auto zero=property(field,"AllowZeroLength");propertyZero->setChecked(zero.kind==VbaValue::Kind::Boolean ? zero.boolean : true);propertyZero->setEnabled(eligible && (selectedType==10 || selectedType==12));
            if(eligible) {selectedField=std::move(field);apply->setEnabled(true);renameButton->setEnabled(true);}
        } catch(const std::exception& problem) {disable();error(problem);}
    };
    try {table=makeDaoDatabase(database)->member("TableDefs",{VbaValue::string(tableName.toStdString())}).object;reload();}
    catch(const std::exception& problem) {error(problem);return;}
    QObject::connect(list,&QTableWidget::itemSelectionChanged,&dialog,[&] {updateDelete();loadProperties();});updateDelete();loadProperties();
    QObject::connect(type,&QComboBox::currentIndexChanged,&dialog,[&] {size->setEnabled(type->currentData().toInt()==10);allowZero->setEnabled(type->currentData().toInt()==10 || type->currentData().toInt()==12);});
    const auto refreshAfterCommit=[&](const QString& message) {status->setText(message);try {reload();updateDelete();}catch(const std::exception& problem) {disable();QMessageBox::warning(&dialog,"Änderung gespeichert",message+"\nDie Feldliste konnte nicht neu geladen werden: "+QString::fromUtf8(problem.what()));}};
    const auto storedFailure=[&](const DaoStoredSchemaChange& problem) {disable();status->setText("Änderung gespeichert. Bitte das Projekt neu öffnen.");QMessageBox::warning(&dialog,"Änderung gespeichert",status->text()+"\n"+QString::fromUtf8(problem.what()));};
    QObject::connect(append,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid) return;const auto fieldName=name->text().trimmed();
        try {const auto code=type->currentData().toInt();auto field=table->member("CreateField",{VbaValue::string(fieldName.toStdString()),VbaValue::integer(static_cast<std::int16_t>(code)),VbaValue::longInteger(code==10 ? size->value() : 0)}).object;
            field->assign("Required",VbaValue::logical(required->isChecked()));if(code==10 || code==12) field->assign("AllowZeroLength",VbaValue::logical(allowZero->isChecked()));if(!defaultValue->text().trimmed().isEmpty()) field->assign("DefaultValue",VbaValue::string(defaultValue->text().trimmed().toStdString()));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});
        } catch(const std::exception& problem) {error(problem);return;}refreshAfterCommit("Feld angelegt: "+fieldName);
    });
    QObject::connect(apply,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid || !selectedField) return;QString fieldName;
        try {fieldName=QString::fromStdString(selectedField->member("Name",{}).asText());QMap<QString,VbaValue> changes{{"Required",VbaValue::logical(propertyRequired->isChecked())}};
            if(propertyZero->isEnabled()) changes.insert("AllowZeroLength",VbaValue::logical(propertyZero->isChecked()));if(selectedType==10) changes.insert("Size",VbaValue::longInteger(propertySize->value()));if(propertyDefault->isEnabled()) changes.insert("DefaultValue",VbaValue::string(propertyDefault->text().trimmed().toStdString()));updateNativeFieldProperties(database,tableName,fieldName,changes);
        } catch(const DaoStoredSchemaChange& problem) {storedFailure(problem);return;}catch(const std::exception& problem) {error(problem);return;}refreshAfterCommit("Feldeigenschaften gespeichert: "+fieldName);
    });
    QObject::connect(renameButton,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid || !selectedField) return;const auto newName=renameName->text().trimmed();
        try {const auto oldName=QString::fromStdString(selectedField->member("Name",{}).asText());updateNativeFieldProperties(database,tableName,oldName,{{"Name",VbaValue::string(newName.toStdString())}});}
        catch(const DaoStoredSchemaChange& problem) {storedFailure(problem);return;}catch(const std::exception& problem) {error(problem);return;}
        refreshAfterCommit("Feld umbenannt: "+newName+". Gespeicherte Access-Abfragen und Formularbindungen bitte auf den neuen Feldnamen anpassen.");
        for(int row=0;row<list->rowCount();++row) if(list->item(row,0)->text()==newName) list->setCurrentCell(row,0);
    });
    QObject::connect(remove,&QPushButton::clicked,&dialog,[&] {
        if(!catalogValid) return;const int row=list->currentRow();if(row<0 || list->item(row,0)->data(Qt::UserRole).toBool()) return;const auto fieldName=list->item(row,0)->text();
        if(QMessageBox::question(&dialog,"Feld löschen","Feld „"+fieldName+"“ und alle darin gespeicherten Werte löschen? Verweise in Abfragen und Formularen werden nicht automatisch angepasst.",QMessageBox::Yes|QMessageBox::No,QMessageBox::No)!=QMessageBox::Yes) return;
        try {table->member("Fields",{}).object->member("Delete",{VbaValue::string(fieldName.toStdString())});}catch(const std::exception& problem) {error(problem);return;}refreshAfterCommit("Feld gelöscht: "+fieldName);
    });
    dialog.exec();
}
}
