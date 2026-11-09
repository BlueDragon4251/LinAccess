#include "native_query_builder.hpp"
#include "native_query_sources.hpp"
#include <algorithm>
#include "access_query.hpp"
#include "vba_runtime.hpp"
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QHeaderView>
#include <QJsonArray>
#include <QLabel>
#include <QLineEdit>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QRegularExpression>
#include <QSet>
#include <QSignalBlocker>
#include <QStyledItemDelegate>
#include <QTableWidget>
#include <QVBoxLayout>
namespace linaccess {
namespace {
QString quoted(QString name) {
    if(name.isEmpty() || name.size()>64 || !name.isValidUtf16() || name.contains(QChar(0))) throw VbaRuntimeError(5,"Ungültiger Tabellen-, Feld- oder Aliasname.");
    name.replace(']',"]]");return '['+name+']';
}
const QStringList parameterTypes{"Text","Long","Integer","Byte","Single","Double","Currency","DateTime","YesNo"};
QString parameterHeader(const QJsonArray& parameters) {
    if(parameters.size()>64) throw VbaRuntimeError(3251,"Abfrageentwurf unterstützt höchstens64 deklarierte Parameter.");
    QStringList declarations;QSet<QString> names;
    for(const auto& value:parameters) {
        if(!value.isObject())throw VbaRuntimeError(5,"Ungültiger Parameterentwurf.");const auto entry=value.toObject();
        const auto name=entry.value("name").toString().trimmed(),type=entry.value("type").toString();
        if(!parameterTypes.contains(type) || names.contains(name.toCaseFolded()))throw VbaRuntimeError(5,"Parameter benötigen eindeutige Namen und unterstützte Typen.");
        declarations.append(quoted(name)+' '+type);names.insert(name.toCaseFolded());
    }
    return declarations.isEmpty() ? QString{} : "PARAMETERS "+declarations.join(", ")+"; ";
}
std::optional<QJsonArray> editParameters(QWidget* parent,const QJsonArray& initial,const std::function<void()>& checkProject) {
    (void)parameterHeader(initial);QDialog dialog(parent);dialog.setObjectName("native-query-parameters");dialog.setWindowTitle("Abfrageparameter");auto* layout=new QVBoxLayout(&dialog);
    auto* grid=new QTableWidget(0,2,&dialog);grid->setObjectName("native-query-parameter-fields");grid->setHorizontalHeaderLabels({"Name","Typ"});grid->horizontalHeader()->setStretchLastSection(true);layout->addWidget(grid);
    auto* row=new QHBoxLayout;auto* add=new QPushButton("Hinzufügen",&dialog);auto* remove=new QPushButton("Entfernen",&dialog);add->setObjectName("native-query-parameter-add");remove->setObjectName("native-query-parameter-remove");row->addWidget(add);row->addWidget(remove);layout->addLayout(row);
    auto* status=new QLabel(&dialog);status->setTextFormat(Qt::PlainText);layout->addWidget(status);auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);buttons->button(QDialogButtonBox::Ok)->setObjectName("native-query-parameter-apply");layout->addWidget(buttons);QObject connections;std::optional<QJsonArray> result;
    const auto capture=[&]{QJsonArray values;for(int i=0;i<grid->rowCount();++i)values.append(QJsonObject{{"name",grid->item(i,0)->text().trimmed()},{"type",qobject_cast<QComboBox*>(grid->cellWidget(i,1))->currentText()}});return values;};
    const auto append=[&](const QJsonObject& entry){const int i=grid->rowCount();if(i>=64)return;grid->insertRow(i);auto* name=new QTableWidgetItem(entry.value("name").toString());grid->setItem(i,0,name);auto* type=new QComboBox(grid);type->addItems(parameterTypes);type->setCurrentText(entry.value("type").toString("Text"));grid->setCellWidget(i,1,type);grid->setCurrentCell(i,0);};
    for(const auto& value:initial)append(value.toObject());
    QObject::connect(add,&QPushButton::clicked,&connections,[&]{append({});});QObject::connect(remove,&QPushButton::clicked,&connections,[&]{if(grid->currentRow()>=0)grid->removeRow(grid->currentRow());});
    QObject::connect(buttons,&QDialogButtonBox::rejected,&connections,[&]{dialog.reject();});QObject::connect(buttons,&QDialogButtonBox::accepted,&connections,[&]{try{checkProject();auto values=capture();(void)compileAccessQuery(parameterHeader(values)+"SELECT 1");result=values;dialog.accept();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    dialog.resize(480,360);if(dialog.exec()!=QDialog::Accepted)return {};return result;
}
const QStringList totalModes{"Group By","Where","Count","Sum","Avg","Min","Max","Var","StDev"};
struct TotalsDesign {bool enabled;QStringList modes;};
std::optional<TotalsDesign> editTotals(QWidget* parent,const QJsonObject& initial,const std::function<void()>& checkProject) {
    QDialog dialog(parent);dialog.setObjectName("native-query-totals");dialog.setWindowTitle("Gruppierung und Aggregate");auto* layout=new QVBoxLayout(&dialog);
    auto* enabled=new QCheckBox("Gruppierung / Aggregate verwenden",&dialog);enabled->setObjectName("native-query-totals-enabled");enabled->setChecked(initial.value("totals").toBool());layout->addWidget(enabled);
    auto* grid=new QTableWidget(0,2,&dialog);grid->setObjectName("native-query-total-fields");grid->setHorizontalHeaderLabels({"Feld / Ausdruck","Funktion"});grid->horizontalHeader()->setStretchLastSection(true);layout->addWidget(grid);
    for(const auto& value:initial.value("fields").toArray()){const auto field=value.toObject();const auto mode=field.value("total").toString("Group By");if(!totalModes.contains(mode))throw VbaRuntimeError(5,"Ungültige gespeicherte Aggregatfunktion.");const int row=grid->rowCount();grid->insertRow(row);auto display=field.value(field.contains("expression") ? "expression" : "name").toString();if(!field.contains("expression") && !initial.value("joins").toArray().isEmpty())display=field.value("table").toString(initial.value("table").toString())+'.'+display;auto* label=new QTableWidgetItem(display);label->setFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable);grid->setItem(row,0,label);auto* choice=new QComboBox(grid);choice->addItems(totalModes);choice->setCurrentText(mode);grid->setCellWidget(row,1,choice);}
    auto* status=new QLabel(&dialog);status->setTextFormat(Qt::PlainText);layout->addWidget(status);auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);buttons->button(QDialogButtonBox::Ok)->setObjectName("native-query-totals-apply");layout->addWidget(buttons);std::optional<TotalsDesign> result;QObject connections;
    QObject::connect(buttons,&QDialogButtonBox::rejected,&connections,[&]{dialog.reject();});QObject::connect(buttons,&QDialogButtonBox::accepted,&connections,[&]{try{checkProject();TotalsDesign edited{enabled->isChecked(),{}};for(int row=0;row<grid->rowCount();++row)edited.modes.append(qobject_cast<QComboBox*>(grid->cellWidget(row,1))->currentText());result=edited;dialog.accept();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    dialog.resize(600,420);if(dialog.exec()!=QDialog::Accepted)return {};return result;
}
class BuilderDelegate final:public QStyledItemDelegate {
public:using QStyledItemDelegate::QStyledItemDelegate;
    QWidget* createEditor(QWidget* parent,const QStyleOptionViewItem& option,const QModelIndex& index) const override {
        auto* editor=QStyledItemDelegate::createEditor(parent,option,index);
        if(auto* text=qobject_cast<QLineEdit*>(editor)) text->setMaxLength(index.column()==2 ? 64 : 4096);
        return editor;
    }
};
QString designSql(QSqlDatabase db,const QJsonObject& design) {
    const auto table=design.value("table").toString();
    if(!db.tables(static_cast<QSql::TableType>(QSql::Tables|QSql::Views)).contains(table)) throw VbaRuntimeError(3078,"Tabelle oder gespeicherte Abfrage ist nicht verfügbar.");
    const auto sources=nativeQuerySources(db,table,design.value("joins").toArray());
    const auto rows=design.value("fields").toArray();if(rows.isEmpty() || rows.size()>255) throw VbaRuntimeError(3251,"Abfrageentwurf benötigt1..255 Felder.");
    const auto count=design.value("groups").toInt(1);if(count<1 || count>8)throw VbaRuntimeError(3251,"Abfrageentwurf unterstützt1..8 Kriteriengruppen.");
    QStringList projection,sort,grouping;QVector<QStringList> predicates(count),having(count);const bool totals=design.value("totals").toBool();QSet<QString> used,labels;
    const bool distinct=design.value("distinct").toBool();
    static const QRegularExpression operatorPrefix("^(?:[=<>]|(?:Not\\s+)?Like\\b|Is\\s+(?:Not\\s+)?Null\\b|Between\\b|In\\s*\\()",QRegularExpression::CaseInsensitiveOption);
    static const QRegularExpression literalPrefix("^(?:['\"#]|[+-]?(?:[0-9]|\\.[0-9])|True$|False$)",QRegularExpression::CaseInsensitiveOption);
    for(const auto& value:rows) {
        if(!value.isObject()) throw VbaRuntimeError(5,"Ungültiger Feldentwurf.");const auto row=value.toObject();
        const auto field=row.value("name").toString();const bool calculated=row.contains("expression");
        QString reference;
        if(calculated) {const auto expression=row.value("expression").toString().trimmed();if(expression.size()>4096)throw VbaRuntimeError(5,"Berechneter Ausdruck ist zu lang.");reference='('+accessScalarDomainExpression(expression,true)+')';}
        else {const auto source=row.value("table").toString(table);const auto key=source.toCaseFolded()+QChar(0)+field.toCaseFolded();if(!sources.fields.value(source).contains(field) || used.contains(key))throw VbaRuntimeError(5,"Ein Feld fehlt oder wurde mehrfach aufgenommen.");used.insert(key);reference=quoted(source)+'.'+quoted(field);}
        const bool visible=row.value("visible").toBool();const auto alias=row.value("alias").toString().trimmed();
        if(calculated && alias.isEmpty())throw VbaRuntimeError(5,"Berechnetes Feld benötigt einen Alias.");
        const auto total=row.value("total").toString("Group By");if(!totalModes.contains(total))throw VbaRuntimeError(5,"Ungültige Aggregatfunktion.");
        const bool where=totals && total=="Where";
        if(where && (visible || row.value("sort").toInt()!=0))throw VbaRuntimeError(5,"Where-Felder müssen ausgeblendet und unsortiert sein.");
        if(totals && total=="Group By")grouping.append(reference);
        if(totals && total!="Group By" && !where)reference=total+'('+reference+')';
        if(visible) {
            const auto label=alias.isEmpty() ? field : alias;if(labels.contains(label.toCaseFolded())) throw VbaRuntimeError(5,"Ausgabefelder benötigen eindeutige Namen.");labels.insert(label.toCaseFolded());
            projection.append(reference+(alias.isEmpty() ? QString{} : " AS "+quoted(alias)));
        }
        const auto direction=row.value("sort").toInt();if(direction<0 || direction>2) throw VbaRuntimeError(5,"Ungültige Sortierung.");
        if(direction) {if(distinct && !visible) throw VbaRuntimeError(3251,"Bei eindeutigen Zeilen muss ein Sortierfeld auch ausgegeben werden.");sort.append(reference+(direction==1 ? " ASC" : " DESC"));}
        QStringList criteria{row.value("criteria").toString()};
        const auto alternatives=row.value("or_criteria").toArray();if(alternatives.size()>count-1)throw VbaRuntimeError(5,"Zu viele alternative Feldkriterien.");
        for(const auto& alternative:alternatives) {if(!alternative.isString())throw VbaRuntimeError(5,"Ungültiges ODER-Kriterium.");criteria.append(alternative.toString());}
        for(int group=0;group<criteria.size();++group) {
            auto criterion=criteria[group].trimmed();if(criterion.size()>4096)throw VbaRuntimeError(5,"Kriterium ist zu lang.");
            if(!criterion.isEmpty()) {
                if(!operatorPrefix.match(criterion).hasMatch()){if(!literalPrefix.match(criterion).hasMatch())throw VbaRuntimeError(5,"Kriterium benötigt einen Vergleich oder ein Literal.");criterion="= "+criterion;}
                const auto predicate=reference+' '+criterion;(void)accessWhereClause(predicate);(totals && !where ? having : predicates)[group].append('('+predicate+')');
            }
        }
    }
    if(projection.isEmpty()) throw VbaRuntimeError(5,"Mindestens ein Feld zur Ausgabe auswählen.");
    auto sql=parameterHeader(design.value("parameters").toArray())+QString(distinct ? "SELECT DISTINCT " : "SELECT ")+projection.join(", ")+" FROM "+sources.sql;
    const auto combine=[](const QVector<QStringList>& alternatives){QStringList groups;for(const auto& group:alternatives)if(!group.isEmpty())groups.append('('+group.join(" AND ")+')');return groups.join(" OR ");};
    const auto predicate=combine(predicates);if(!predicate.isEmpty()){(void)accessWhereClause(predicate);sql+=" WHERE "+predicate;}
    if(!grouping.isEmpty())sql+=" GROUP BY "+grouping.join(", ");
    const auto aggregatePredicate=combine(having);if(!aggregatePredicate.isEmpty()){(void)accessWhereClause(aggregatePredicate);sql+=" HAVING "+aggregatePredicate;}
    if(!sort.isEmpty()) sql+=" ORDER BY "+sort.join(", ");
    const auto plan=compileAccessQuery(sql);if(plan.daoQueryType!=0 || !plan.makeTable.isEmpty()) throw VbaRuntimeError(3251,"Entwurf muss eine Auswahlabfrage sein.");
    (void)accessQueryColumns(db,plan);return sql;
}
}
std::optional<NativeQueryDesign> editNativeQueryDesign(QWidget* parent,QSqlDatabase database,const QJsonObject& initial,const std::function<void()>& checkProject,bool allowParameters) {
    checkProject();if(installAccessQueryViews(database)<0)throw VbaRuntimeError(3251,"Gespeicherte Abfragequellen sind momentan nicht verfügbar.");const auto initialParameters=initial.value("parameters").toArray();(void)parameterHeader(initialParameters);if(!allowParameters && !initialParameters.isEmpty())throw VbaRuntimeError(3251,"Diese Datenquelle unterstützt keine Abfrageparameter.");QDialog dialog(parent);dialog.setObjectName("native-query-builder");dialog.setWindowTitle("Abfrageentwurf");dialog.resize(900,650);auto* layout=new QVBoxLayout(&dialog);
    auto* source=new QComboBox(&dialog);source->setObjectName("native-query-builder-source");source->setPlaceholderText("Tabelle oder gespeicherte Abfrage auswählen");source->setAccessibleName("Tabelle oder gespeicherte Abfrage");
    auto tables=database.tables(static_cast<QSql::TableType>(QSql::Tables|QSql::Views));tables.removeDuplicates();std::sort(tables.begin(),tables.end(),[](const auto& a,const auto& b){return a.compare(b,Qt::CaseInsensitive)<0;});
    for(const auto& table:tables) if(!table.startsWith("_la_",Qt::CaseInsensitive) && !table.startsWith("sqlite_",Qt::CaseInsensitive)) source->addItem(table);
    source->setCurrentIndex(-1);layout->addWidget(source);
    auto* distinct=new QCheckBox("Nur eindeutige Zeilen",&dialog);distinct->setObjectName("native-query-builder-distinct");layout->addWidget(distinct);
    auto* grid=new QTableWidget(0,5,&dialog);grid->setObjectName("native-query-builder-fields");grid->setHorizontalHeaderLabels({"Ausgabe","Feld / Ausdruck","Alias","Sortierung","Kriterium"});grid->setItemDelegate(new BuilderDelegate(grid));grid->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);grid->horizontalHeader()->setStretchLastSection(true);layout->addWidget(grid,1);
    auto* rows=new QHBoxLayout;auto* up=new QPushButton("Feld nach oben",&dialog);auto* down=new QPushButton("Feld nach unten",&dialog);auto* addOr=new QPushButton("ODER-Gruppe hinzufügen",&dialog);auto* removeOr=new QPushButton("Letzte ODER-Gruppe entfernen",&dialog);auto* params=new QPushButton("Parameter…",&dialog);auto* reload=new QPushButton("Quelle neu einlesen",&dialog);
    up->setObjectName("native-query-builder-up");down->setObjectName("native-query-builder-down");addOr->setObjectName("native-query-builder-or-add");removeOr->setObjectName("native-query-builder-or-remove");params->setObjectName("native-query-builder-parameters");params->setVisible(allowParameters);reload->setObjectName("native-query-builder-reload");rows->addWidget(up);rows->addWidget(down);rows->addWidget(addOr);rows->addWidget(removeOr);rows->addWidget(params);rows->addWidget(reload);layout->addLayout(rows);
    auto* expressions=new QHBoxLayout;auto* addExpression=new QPushButton("Ausdruck hinzufügen",&dialog);auto* removeField=new QPushButton("Feld entfernen",&dialog);addExpression->setObjectName("native-query-builder-expression-add");removeField->setObjectName("native-query-builder-field-remove");expressions->addWidget(addExpression);expressions->addWidget(removeField);auto* totalsButton=new QPushButton("Gruppierung…",&dialog);totalsButton->setObjectName("native-query-builder-totals");expressions->addWidget(totalsButton);auto* joinsButton=new QPushButton("Quellen verknüpfen…",&dialog);joinsButton->setObjectName("native-query-builder-joins");expressions->addWidget(joinsButton);expressions->addStretch();layout->addLayout(expressions);
    auto* preview=new QPlainTextEdit(&dialog);preview->setObjectName("native-query-builder-sql");preview->setReadOnly(true);preview->setMaximumHeight(130);layout->addWidget(preview);
    auto* status=new QLabel(&dialog);status->setObjectName("native-query-builder-status");status->setTextFormat(Qt::PlainText);status->setWordWrap(true);layout->addWidget(status);
    auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);auto* apply=buttons->button(QDialogButtonBox::Ok);apply->setText("Übernehmen");apply->setObjectName("native-query-builder-apply");layout->addWidget(buttons);
    bool loading=false,totals=initial.value("totals").toBool();QJsonArray parameters=initialParameters,joins;std::optional<NativeQueryDesign> result;QObject connections;
    const auto capture=[&] {QJsonArray fields;for(int row=0;row<grid->rowCount();++row){QJsonArray alternatives;for(int column=5;column<grid->columnCount();++column)alternatives.append(grid->item(row,column)->text());const bool calculated=grid->item(row,1)->data(Qt::UserRole).toBool();QJsonObject field{{"name",calculated ? QString{} : grid->item(row,1)->data(Qt::UserRole+3).toString()},{"table",grid->item(row,1)->data(Qt::UserRole+2).toString()},{"visible",grid->item(row,0)->checkState()==Qt::Checked},{"alias",grid->item(row,2)->text()},{"sort",qobject_cast<QComboBox*>(grid->cellWidget(row,3))->currentIndex()},{"criteria",grid->item(row,4)->text()},{"or_criteria",alternatives}};field.insert("total",grid->item(row,1)->data(Qt::UserRole+1).toString());if(calculated)field.insert("expression",grid->item(row,1)->text());fields.append(field);}return QJsonObject{{"version",2},{"table",source->currentText()},{"distinct",distinct->isChecked()},{"totals",totals},{"joins",joins},{"groups",grid->columnCount()-4},{"parameters",parameters},{"fields",fields}};};
    const auto refresh=[&] {if(loading)return;totalsButton->setText(totals ? "Gruppierung aktiv…" : "Gruppierung…");joinsButton->setEnabled(source->currentIndex()>=0);addExpression->setEnabled(source->currentIndex()>=0 && grid->rowCount()<255);removeField->setEnabled(grid->rowCount()>0);addOr->setEnabled(grid->columnCount()<12);removeOr->setEnabled(grid->columnCount()>5);if(source->currentIndex()<0){preview->clear();status->setText("Bitte eine Tabelle oder gespeicherte Abfrage auswählen.");apply->setEnabled(false);return;}try{checkProject();const auto sql=designSql(database,capture());checkProject();preview->setPlainText(sql);status->setText("Entwurf gültig. Übernehmen aktualisiert den SQL-Editor.");apply->setEnabled(true);}catch(const std::exception& error){preview->clear();status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}};
    const auto append=[&](const QJsonObject& field) {const int row=grid->rowCount();if(row>=255)throw VbaRuntimeError(3251,"Abfrageentwurf unterstützt höchstens255 Feldzeilen.");grid->insertRow(row);auto* output=new QTableWidgetItem;output->setFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable|Qt::ItemIsUserCheckable);output->setCheckState(field.value("visible").toBool(true) ? Qt::Checked : Qt::Unchecked);grid->setItem(row,0,output);const bool calculated=field.contains("expression");auto* name=new QTableWidgetItem(field.value(calculated ? "expression" : "name").toString());name->setData(Qt::UserRole,calculated);const auto fieldSource=field.value("table").toString(source->currentText());name->setData(Qt::UserRole+2,fieldSource);name->setData(Qt::UserRole+3,field.value("name").toString());if(!calculated && !joins.isEmpty())name->setText(fieldSource+"."+field.value("name").toString());const auto total=field.value("total").toString("Group By");if((field.contains("total") && !field.value("total").isString()) || !totalModes.contains(total))throw VbaRuntimeError(5,"Ungültige gespeicherte Aggregatfunktion.");name->setData(Qt::UserRole+1,total);name->setToolTip(total);name->setFlags(calculated ? Qt::ItemFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable|Qt::ItemIsEditable) : Qt::ItemFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable));grid->setItem(row,1,name);grid->setItem(row,2,new QTableWidgetItem(field.value("alias").toString()));auto* order=new QComboBox(grid);order->addItems({"(keine)","Aufsteigend","Absteigend"});order->setCurrentIndex(field.value("sort").toInt());grid->setCellWidget(row,3,order);grid->setItem(row,4,new QTableWidgetItem(field.value("criteria").toString()));const auto alternatives=field.value("or_criteria").toArray();for(int column=5;column<grid->columnCount();++column)grid->setItem(row,column,new QTableWidgetItem((column-5<alternatives.size() ? alternatives.at(column-5).toString() : QString{})));QObject::connect(order,&QComboBox::currentIndexChanged,&connections,refresh);};
    const auto populate=[&](const QJsonObject& state,bool restore) {
        loading=true;joins=restore ? state.value("joins").toArray() : QJsonArray{};totals=restore && state.value("totals").toBool();grid->setRowCount(0);
        const auto groups=restore ? state.value("groups").toInt(1) : 1;
        if(groups<1 || groups>8){loading=false;apply->setEnabled(false);status->setText("Ungültige Zahl gespeicherter Kriteriengruppen.");return;}
        grid->setColumnCount(groups+4);QStringList headers{"Ausgabe","Feld / Ausdruck","Alias","Sortierung","Kriterium"};for(int i=1;i<groups;++i)headers.append("oder "+QString::number(i));grid->setHorizontalHeaderLabels(headers);
        try {
            checkProject();if(restore && ((state.contains("joins") && !state.value("joins").isArray()) || (state.contains("totals") && !state.value("totals").isBool())))throw VbaRuntimeError(5,"Ungültige gespeicherte Verknüpfungs- oder Gruppierungsmetadaten.");const auto table=source->currentText();const auto resolved=nativeQuerySources(database,table,joins);int fieldCount=0;for(const auto& fields:resolved.fields)fieldCount+=fields.size();
            if(fieldCount>255)throw VbaRuntimeError(3251,"Quelle hat mehr als255 Felder.");QSet<QString> added;
            if(restore && (!state.value("fields").isArray() || state.value("fields").toArray().size()>255))throw VbaRuntimeError(3251,"Gespeicherter Feldentwurf überschreitet den unterstützten Umfang.");
            if(restore) for(const auto& value:state.value("fields").toArray()) {if(!value.isObject())throw VbaRuntimeError(5,"Ungültiger gespeicherter Feldentwurf.");const auto field=value.toObject();if((field.contains("table") && !field.value("table").isString()) || !field.value("name").isString() || !field.value("visible").isBool() || !field.value("alias").isString() || !field.value("sort").isDouble() || field.value("sort").toDouble()!=field.value("sort").toInt() || !field.value("criteria").isString())throw VbaRuntimeError(5,"Ungültige gespeicherte Feldeigenschaften.");if(field.contains("or_criteria") && (!field.value("or_criteria").isArray() || field.value("or_criteria").toArray().size()>groups-1))throw VbaRuntimeError(5,"Ungültige gespeicherte ODER-Kriterien.");for(const auto& criterion:field.value("or_criteria").toArray())if(!criterion.isString() || criterion.toString().size()>4096)throw VbaRuntimeError(5,"Ungültiges gespeichertes ODER-Kriterium.");if(field.value("alias").toString().size()>64 || field.value("criteria").toString().size()>4096)throw VbaRuntimeError(5,"Gespeicherter Alias oder Vergleich ist zu lang.");const auto name=field.value("name").toString(),fieldSource=field.value("table").toString(table),key=fieldSource.toCaseFolded()+QChar(0)+name.toCaseFolded();if(field.contains("expression")){if(!field.value("expression").isString() || field.value("expression").toString().size()>4096)throw VbaRuntimeError(5,"Ungültiger gespeicherter Ausdruck.");append(field);}else{if(!resolved.fields.value(fieldSource).contains(name) || added.contains(key))throw VbaRuntimeError(5,"Gespeichertes Feld fehlt; Quelle neu einlesen.");append(field);added.insert(key);}}
            for(const auto& fieldSource:resolved.order)for(const auto& field:resolved.fields.value(fieldSource))if(!added.contains(fieldSource.toCaseFolded()+QChar(0)+field.toCaseFolded()) && grid->rowCount()<255)append({{"name",field},{"table",fieldSource},{"visible",!restore},{"total",restore ? "Where" : "Group By"}});
            distinct->setChecked(restore && state.value("distinct").toBool());loading=false;refresh();
        } catch(const std::exception& error) {grid->setRowCount(0);loading=false;preview->clear();status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}
    };
    QObject::connect(source,&QComboBox::currentIndexChanged,&connections,[&]{populate({},false);});QObject::connect(grid,&QTableWidget::itemChanged,&connections,refresh);QObject::connect(distinct,&QCheckBox::toggled,&connections,refresh);QObject::connect(reload,&QPushButton::clicked,&connections,[&]{populate({},false);});
    QObject::connect(addExpression,&QPushButton::clicked,&connections,[&]{if(source->currentIndex()<0 || grid->rowCount()>=255)return;loading=true;append({{"name",QString{}},{"expression",QString{}},{"visible",true}});grid->setCurrentCell(grid->rowCount()-1,1);loading=false;refresh();grid->setFocus();grid->editItem(grid->item(grid->rowCount()-1,1));});
    QObject::connect(removeField,&QPushButton::clicked,&connections,[&]{const int row=grid->currentRow();if(row<0)return;loading=true;grid->removeRow(row);loading=false;refresh();});
    QObject::connect(addOr,&QPushButton::clicked,&connections,[&]{if(grid->columnCount()>=12)return;loading=true;const int column=grid->columnCount();grid->insertColumn(column);grid->setHorizontalHeaderItem(column,new QTableWidgetItem("oder "+QString::number(column-4)));for(int row=0;row<grid->rowCount();++row)grid->setItem(row,column,new QTableWidgetItem);loading=false;refresh();});
    QObject::connect(removeOr,&QPushButton::clicked,&connections,[&]{if(grid->columnCount()>5){loading=true;grid->removeColumn(grid->columnCount()-1);loading=false;refresh();}});
    QObject::connect(params,&QPushButton::clicked,&connections,[&]{try{checkProject();const auto edited=editParameters(&dialog,parameters,checkProject);if(edited)parameters=*edited;refresh();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}});
    QObject::connect(totalsButton,&QPushButton::clicked,&connections,[&]{try{checkProject();const auto edited=editTotals(&dialog,capture(),checkProject);if(edited){loading=true;totals=edited->enabled;for(int row=0;row<grid->rowCount();++row){auto* name=grid->item(row,1);name->setData(Qt::UserRole+1,edited->modes[row]);name->setToolTip(edited->modes[row]);}}loading=false;refresh();}catch(const std::exception& error){loading=false;status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}});
    QObject::connect(joinsButton,&QPushButton::clicked,&connections,[&]{try{checkProject();auto state=capture();const auto edited=editNativeQueryJoins(&dialog,database,source->currentText(),joins,checkProject);if(!edited)return;state.insert("joins",*edited);const auto resolved=nativeQuerySources(database,source->currentText(),*edited);QJsonArray retained;for(const auto& value:state.value("fields").toArray()){const auto field=value.toObject();if(field.contains("expression") || resolved.fields.contains(field.value("table").toString(source->currentText())))retained.append(field);}state.insert("fields",retained);populate(state,true);}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}});
    const auto move=[&](int delta){const int row=grid->currentRow(),target=row+delta;if(row<0 || target<0 || target>=grid->rowCount())return;auto fields=capture().value("fields").toArray();const QJsonValue first=fields[row];fields[row]=fields[target];fields[target]=first;loading=true;grid->setRowCount(0);for(const auto& field:fields)append(field.toObject());grid->setCurrentCell(target,2);loading=false;refresh();};QObject::connect(up,&QPushButton::clicked,&connections,[&]{move(-1);});QObject::connect(down,&QPushButton::clicked,&connections,[&]{move(1);});
    QObject::connect(buttons,&QDialogButtonBox::rejected,&connections,[&]{dialog.reject();});QObject::connect(buttons,&QDialogButtonBox::accepted,&connections,[&]{apply->setFocus();try{checkProject();auto state=capture();auto sql=designSql(database,state);checkProject();result=NativeQueryDesign{sql,state};dialog.accept();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}});
    if((initial.value("version").toInt()==1 || initial.value("version").toInt()==2) && source->findText(initial.value("table").toString())>=0) {const QSignalBlocker blocked(source);source->setCurrentIndex(source->findText(initial.value("table").toString()));populate(initial,true);}else refresh();
    if(dialog.exec()!=QDialog::Accepted)return {};return result;
}
}
