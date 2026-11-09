#include "native_query_action.hpp"
#include "access_query.hpp"
#include "vba_runtime.hpp"
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QHeaderView>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QSet>
#include <QStandardItemModel>
#include <QTableWidget>
#include <QVBoxLayout>
namespace linaccess {
namespace {
QString quote(QString name) {
    if(name.isEmpty() || name.size()>64 || !name.isValidUtf16() || name.contains(QChar(0)))throw VbaRuntimeError(5,"Ungültiger Zielname.");
    name.replace(']',"]]");return '['+name+']';
}
bool internal(const QString& name){return name.startsWith("_la_",Qt::CaseInsensitive) || name.startsWith("sqlite_",Qt::CaseInsensitive);}
QStringList targetFields(QSqlDatabase database,const QString& target) {
    if(internal(target) || !database.tables(QSql::Tables).contains(target))throw VbaRuntimeError(3078,"Anfügeziel ist keine verfügbare Projekttabelle.");
    const auto fields=accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quote(target)));if(fields.size()>255)throw VbaRuntimeError(3251,"Anfügeziel hat mehr als255 Felder.");return fields;
}
}
QString nativeQueryActionSql(QSqlDatabase database,const QJsonObject& action,const QString& header,const QString& selection,const QString& intoSelection,const QString& physicalSource,const QString& predicate) {
    if(action.contains("type") && !action.value("type").isString())throw VbaRuntimeError(5,"Ungültige Abfrageart.");const auto type=action.value("type").toString("select");
    if(type=="select")return header+selection;
    if(type=="update" || type=="delete") {
        if(physicalSource.isEmpty() || internal(physicalSource))throw VbaRuntimeError(3251,"Aktualisieren/Löschen benötigt eine einzelne Projekttabelle ohne Verknüpfung, Summen, DISTINCT oder Sortierung.");const auto fields=targetFields(database,physicalSource);(void)accessWhereClause(predicate);const auto where=predicate.isEmpty() ? QString{} : " WHERE "+predicate;
        if(type=="delete")return header+"DELETE FROM "+quote(physicalSource)+where;
        if(!action.value("updates").isObject() || action.value("updates").toObject().isEmpty())throw VbaRuntimeError(5,"Bitte mindestens ein Aktualisierungsfeld mit Ausdruck wählen.");QStringList values;QSet<QString> used;
        const auto updates=action.value("updates").toObject();for(auto it=updates.constBegin();it!=updates.constEnd();++it){if(!fields.contains(it.key()) || used.contains(it.key().toCaseFolded()) || !it.value().isString())throw VbaRuntimeError(5,"Ungültiges Aktualisierungsfeld.");const auto expression=it.value().toString().trimmed();if(expression.size()>4096)throw VbaRuntimeError(5,"Aktualisierungsausdruck ist zu lang.");values.append(quote(it.key())+" = "+accessScalarDomainExpression(expression));used.insert(it.key().toCaseFolded());}
        return header+"UPDATE "+quote(physicalSource)+" SET "+values.join(", ")+where;
    }
    if(!action.value("target").isString())throw VbaRuntimeError(5,"Aktionsabfrage benötigt einen Zielnamen.");const auto target=action.value("target").toString().trimmed();(void)quote(target);if(internal(target))throw VbaRuntimeError(3251,"Interne Projektobjekte sind kein Ziel des visuellen Entwurfs.");
    if(type=="make_table") {
        if(database.tables(static_cast<QSql::TableType>(QSql::Tables|QSql::Views)).contains(target,Qt::CaseInsensitive))throw VbaRuntimeError(3010,"Zieltabelle oder Abfrage existiert bereits.");
        return header+intoSelection;
    }
    if(type!="append")throw VbaRuntimeError(3251,"Diese Abfrageart ist im visuellen Entwurf nicht unterstützt.");
    const auto outputs=accessQueryColumns(database,compileAccessQuery(header+selection));const auto available=targetFields(database,target);if(!action.value("mapping").isObject())throw VbaRuntimeError(5,"Anfügeabfrage benötigt Zielspaltenzuordnungen.");const auto mapping=action.value("mapping").toObject();QStringList fields;QSet<QString> used;
    for(const auto& output:outputs){if(!mapping.value(output).isString())throw VbaRuntimeError(5,"Für jede Ausgabespalte bitte ein Anfügefeld zuordnen.");const auto field=mapping.value(output).toString();if(!available.contains(field) || used.contains(field.toCaseFolded()))throw VbaRuntimeError(5,"Anfügefelder müssen vorhanden und eindeutig zugeordnet sein.");used.insert(field.toCaseFolded());fields.append(quote(field));}
    return header+"INSERT INTO "+quote(target)+" ("+fields.join(", ")+") "+selection;
}
std::optional<QJsonObject> editNativeQueryAction(QWidget* parent,QSqlDatabase database,const QStringList& outputs,const QString& physicalSource,const QJsonObject& initial,const std::function<void()>& checkProject) {
    checkProject();if(outputs.isEmpty() || outputs.size()>255)throw VbaRuntimeError(3251,"Aktionsentwurf benötigt1..255 Ausgabespalten.");QDialog dialog(parent);dialog.setObjectName("native-query-action");dialog.setWindowTitle("Abfrageart und Ausgabeziel");auto* layout=new QVBoxLayout(&dialog);auto* form=new QFormLayout;layout->addLayout(form);
    auto* type=new QComboBox(&dialog);type->setObjectName("native-query-action-type");type->addItem("Auswahl","select");type->addItem("Tabellenerstellung","make_table");type->addItem("Anfügen","append");type->addItem("Aktualisieren","update");type->addItem("Löschen","delete");if(physicalSource.isEmpty())if(auto* choices=qobject_cast<QStandardItemModel*>(type->model()))for(int index:{3,4})choices->item(index)->setEnabled(false);const auto mode=type->findData(initial.value("type").toString("select"));if(mode<0)throw VbaRuntimeError(3251,"Unbekannte gespeicherte Abfrageart.");type->setCurrentIndex(mode);form->addRow("Abfrageart",type);
    auto* name=new QLineEdit(&dialog);name->setObjectName("native-query-action-name");name->setMaxLength(64);name->setText(initial.value("target").toString());form->addRow("Neue Tabelle",name);
    auto* target=new QComboBox(&dialog);target->setObjectName("native-query-action-target");auto tables=database.tables(QSql::Tables);tables.removeDuplicates();tables.sort(Qt::CaseInsensitive);for(const auto& table:tables)if(!internal(table))target->addItem(table);target->setCurrentIndex(target->findText(initial.value("target").toString()));form->addRow("Vorhandene Tabelle",target);
    auto* mapping=new QTableWidget(outputs.size(),2,&dialog);mapping->setObjectName("native-query-action-mapping");mapping->setHorizontalHeaderLabels({"Ausgabespalte","Anfügen an Feld"});mapping->horizontalHeader()->setStretchLastSection(true);layout->addWidget(mapping,1);
    for(int row=0;row<outputs.size();++row){auto* item=new QTableWidgetItem(outputs[row]);item->setFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable);mapping->setItem(row,0,item);auto* choice=new QComboBox(mapping);mapping->setCellWidget(row,1,choice);}
    auto* modificationSource=new QLabel("Änderungsziel: "+physicalSource,&dialog);modificationSource->setTextFormat(Qt::PlainText);modificationSource->setWordWrap(true);layout->addWidget(modificationSource);
    auto* updates=new QTableWidget(0,2,&dialog);updates->setObjectName("native-query-action-updates");updates->setHorizontalHeaderLabels({"Ändern / Feld","Aktualisieren auf Ausdruck"});updates->horizontalHeader()->setStretchLastSection(true);layout->addWidget(updates,1);
    if(!physicalSource.isEmpty())for(const auto& field:targetFields(database,physicalSource)){const int row=updates->rowCount();updates->insertRow(row);auto* selected=new QTableWidgetItem(field);selected->setFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable|Qt::ItemIsUserCheckable);selected->setCheckState(initial.value("updates").toObject().contains(field) ? Qt::Checked : Qt::Unchecked);updates->setItem(row,0,selected);auto* expression=new QLineEdit(updates);expression->setMaxLength(4096);expression->setText(initial.value("updates").toObject().value(field).toString());updates->setCellWidget(row,1,expression);}
    auto* explanation=new QLabel("Entwurf und Übernehmen führen keine Aktionsabfrage aus. Erst Ausführen im SQL-Editor erstellt die Tabelle, fügt Datensätze an, aktualisiert oder löscht passende Datensätze. Nicht zugeordnete Zielspalten verwenden die bestehenden nativen Standardwert-/Pflichtfeldregeln.",&dialog);explanation->setWordWrap(true);layout->addWidget(explanation);auto* status=new QLabel(&dialog);status->setTextFormat(Qt::PlainText);status->setWordWrap(true);layout->addWidget(status);auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);buttons->button(QDialogButtonBox::Ok)->setObjectName("native-query-action-apply");layout->addWidget(buttons);QObject connections;std::optional<QJsonObject> result;
    const auto populate=[&]{QStringList fields;try{checkProject();if(target->currentIndex()>=0)fields=targetFields(database,target->currentText());status->clear();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}for(int row=0;row<outputs.size();++row){auto* choice=qobject_cast<QComboBox*>(mapping->cellWidget(row,1));const auto previous=choice->currentText();choice->clear();choice->addItems(fields);const auto saved=initial.value("mapping").toObject().value(outputs[row]).toString();const auto wanted=!previous.isEmpty() ? previous : saved;choice->setCurrentIndex(choice->findText(wanted));}};
    const auto modeChanged=[&]{const auto value=type->currentData().toString();name->setEnabled(value=="make_table");target->setEnabled(value=="append");mapping->setVisible(value=="append");updates->setVisible(value=="update");modificationSource->setVisible(value=="update" || value=="delete");};
    QObject::connect(type,&QComboBox::currentIndexChanged,&connections,modeChanged);QObject::connect(target,&QComboBox::currentIndexChanged,&connections,populate);
    QObject::connect(buttons,&QDialogButtonBox::rejected,&connections,[&]{dialog.reject();});QObject::connect(buttons,&QDialogButtonBox::accepted,&connections,[&]{try{checkProject();QJsonObject action{{"type",type->currentData().toString()}};if(action.value("type").toString()=="make_table"){const auto value=name->text().trimmed();(void)quote(value);if(internal(value) || database.tables(static_cast<QSql::TableType>(QSql::Tables|QSql::Views)).contains(value,Qt::CaseInsensitive))throw VbaRuntimeError(3010,"Neue Zieltabelle benötigt einen freien Projektnamen.");action.insert("target",value);}else if(action.value("type").toString()=="append"){if(target->currentIndex()<0)throw VbaRuntimeError(5,"Bitte eine Zieltabelle wählen.");const auto fields=targetFields(database,target->currentText());QJsonObject assignments;QSet<QString> used;for(int row=0;row<outputs.size();++row){const auto value=qobject_cast<QComboBox*>(mapping->cellWidget(row,1))->currentText();if(!fields.contains(value) || used.contains(value.toCaseFolded()))throw VbaRuntimeError(5,"Bitte alle Ausgabespalten auf eindeutige vorhandene Zielspalten zuordnen.");assignments.insert(outputs[row],value);used.insert(value.toCaseFolded());}action.insert("target",target->currentText());action.insert("mapping",assignments);}else if(action.value("type").toString()=="update" || action.value("type").toString()=="delete"){if(physicalSource.isEmpty())throw VbaRuntimeError(3251,"Aktualisieren/Löschen benötigt eine einzelne Projekttabelle ohne Verknüpfung, Summen, DISTINCT oder Sortierung.");if(action.value("type").toString()=="update"){QJsonObject values;for(int row=0;row<updates->rowCount();++row)if(updates->item(row,0)->checkState()==Qt::Checked){const auto expression=qobject_cast<QLineEdit*>(updates->cellWidget(row,1))->text().trimmed();(void)accessScalarDomainExpression(expression);values.insert(updates->item(row,0)->text(),expression);}if(values.isEmpty())throw VbaRuntimeError(5,"Bitte mindestens ein Feld zur Aktualisierung wählen.");action.insert("updates",values);}}result=action;dialog.accept();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    populate();modeChanged();dialog.resize(620,460);if(dialog.exec()!=QDialog::Accepted)return {};return result;
}
}
