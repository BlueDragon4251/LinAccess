#include "native_query_sources.hpp"
#include "access_query.hpp"
#include "vba_runtime.hpp"
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QJsonObject>
#include <QLabel>
#include <QLineEdit>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QSet>
#include <QSignalBlocker>
#include <QTableWidget>
#include <QHeaderView>
#include <QVBoxLayout>
namespace linaccess {
namespace {
QString quote(QString name) {
    if(name.isEmpty() || name.size()>64 || !name.isValidUtf16() || name.contains(QChar(0)))throw VbaRuntimeError(5,"Ungültiger Quellen- oder Feldname.");
    name.replace(']',"]]");return '['+name+']';
}
QStringList sourceNames(QSqlDatabase database) {
    auto names=database.tables(static_cast<QSql::TableType>(QSql::Tables|QSql::Views));
    names.removeDuplicates();names.removeIf([](const auto& name){return name.startsWith("_la_",Qt::CaseInsensitive) || name.startsWith("__la_top_",Qt::CaseInsensitive) || name.startsWith("sqlite_",Qt::CaseInsensitive);});names.sort(Qt::CaseInsensitive);return names;
}
}
NativeQuerySources nativeQuerySources(QSqlDatabase database,const QString& main,const QJsonArray& joins) {
    if(joins.size()>8)throw VbaRuntimeError(3251,"Entwurf unterstützt höchstens acht Verknüpfungen.");
    NativeQuerySources result;const auto available=sourceNames(database);QSet<QString> used;
    const auto add=[&](const QString& name,const QString& identity){(void)quote(identity);if(!available.contains(name))throw VbaRuntimeError(3078,"Verknüpfte Datenquelle fehlt.");if(used.contains(identity.toCaseFolded()))throw VbaRuntimeError(3251,"Quellen benötigen eindeutige Namen; bei Wiederverwendung bitte einen neuen Alias angeben.");used.insert(identity.toCaseFolded());result.order.append(identity);result.origins.insert(identity,name);result.fields.insert(identity,accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quote(name))));};
    add(main,main);result.sql=quote(main);
    const QStringList operations{"=","<>","<","<=",">",">="};
    for(const auto& value:joins) {
        if(!value.isObject())throw VbaRuntimeError(5,"Ungültiger Verknüpfungsentwurf.");const auto join=value.toObject();
        for(const auto* key:{"table","kind"})if(!join.value(key).isString())throw VbaRuntimeError(5,"Ungültige Verknüpfungseigenschaft.");
        const auto table=join.value("table").toString(),kind=join.value("kind").toString();if(kind!="INNER" && kind!="LEFT" && kind!="RIGHT")throw VbaRuntimeError(5,"Ungültige Verknüpfungsart.");
        QJsonArray conditions;
        if(join.contains("conditions")){if(!join.value("conditions").isArray())throw VbaRuntimeError(5,"Ungültige Verknüpfungsvergleiche.");conditions=join.value("conditions").toArray();}
        else{for(const auto* key:{"left_table","left_field","right_field"})if(!join.value(key).isString())throw VbaRuntimeError(5,"Ungültiger gespeicherter Feldvergleich.");conditions.append(QJsonObject{{"left_table",join.value("left_table")},{"left_field",join.value("left_field")},{"right_field",join.value("right_field")},{"operator","="}});}
        if(conditions.isEmpty() || conditions.size()>16)throw VbaRuntimeError(3251,"Eine Verknüpfung benötigt1..16 UND-verknüpfte Feldvergleiche.");
        if(join.contains("alias") && !join.value("alias").isString())throw VbaRuntimeError(5,"Ungültiger Quellenalias.");const auto alias=join.value("alias").toString().trimmed(),identity=alias.isEmpty() ? table : alias;
        const auto previous=result.order;add(table,identity);QStringList predicates;
        for(const auto& entry:conditions){if(!entry.isObject())throw VbaRuntimeError(5,"Ungültiger Feldvergleich.");const auto condition=entry.toObject();for(const auto* key:{"left_table","left_field","right_field","operator"})if(!condition.value(key).isString())throw VbaRuntimeError(5,"Ungültige Feldvergleichseigenschaft.");
            const auto left=condition.value("left_table").toString(),leftField=condition.value("left_field").toString(),rightField=condition.value("right_field").toString(),operation=condition.value("operator").toString();
            if(!previous.contains(left) || !result.fields.value(left).contains(leftField) || !result.fields.value(identity).contains(rightField) || !operations.contains(operation))throw VbaRuntimeError(5,"Ungültige Quelle, Feld oder Vergleichsoperator.");
            predicates.append('('+quote(left)+'.'+quote(leftField)+' '+operation+' '+quote(identity)+'.'+quote(rightField)+')');}
        result.sql='('+result.sql+' '+kind+" JOIN "+quote(table)+(alias.isEmpty() ? QString{} : " AS "+quote(alias))+" ON "+predicates.join(" AND ")+')';
    }
    return result;
}
std::optional<QJsonArray> editNativeQueryJoins(QWidget* parent,QSqlDatabase database,const QString& main,const QJsonArray& initial,const std::function<void()>& checkProject) {
    checkProject();(void)nativeQuerySources(database,main,initial);QJsonArray joins=initial,staged;QDialog dialog(parent);dialog.setObjectName("native-query-joins");dialog.setWindowTitle("Datenquellen verknüpfen");auto* layout=new QVBoxLayout(&dialog);auto* form=new QFormLayout;layout->addLayout(form);
    auto* left=new QComboBox(&dialog);left->setObjectName("native-query-join-left-table");auto* leftField=new QComboBox(&dialog);leftField->setObjectName("native-query-join-left-field");auto* right=new QComboBox(&dialog);right->setObjectName("native-query-join-right-table");auto* rightField=new QComboBox(&dialog);rightField->setObjectName("native-query-join-right-field");auto* kind=new QComboBox(&dialog);kind->setObjectName("native-query-join-kind");kind->addItems({"INNER","LEFT","RIGHT"});auto* operation=new QComboBox(&dialog);operation->setObjectName("native-query-join-operator");operation->addItems({"=","<>","<","<=",">",">="});
    auto* alias=new QLineEdit(&dialog);alias->setObjectName("native-query-join-alias");alias->setMaxLength(64);alias->setPlaceholderText("Bei Wiederverwendung erforderlich");
    form->addRow("Vorhandene Quelle",left);form->addRow("Feld",leftField);form->addRow("Neue Quelle",right);form->addRow("Quellenalias (optional)",alias);form->addRow("Feld",rightField);form->addRow("Vergleich",operation);form->addRow("Verknüpfungsart",kind);auto* semantics=new QLabel("INNER: passende Zeilen. LEFT: alle Zeilen der bisherigen Quellen, rechts bei fehlender Zuordnung Null. RIGHT: alle Zeilen der neuen Quelle.",&dialog);semantics->setWordWrap(true);layout->addWidget(semantics);
    auto* comparisons=new QTableWidget(0,4,&dialog);comparisons->setObjectName("native-query-join-comparisons");comparisons->setHorizontalHeaderLabels({"Vorhandene Quelle","Feld","Vergleich","Neues Feld"});comparisons->horizontalHeader()->setStretchLastSection(true);comparisons->setEditTriggers(QAbstractItemView::NoEditTriggers);comparisons->setMinimumHeight(80);comparisons->setMaximumHeight(120);layout->addWidget(comparisons);
    auto* comparisonButtons=new QHBoxLayout;auto* stage=new QPushButton("Vergleich vormerken",&dialog);stage->setObjectName("native-query-join-condition-add");auto* unstage=new QPushButton("Letzten Vergleich entfernen",&dialog);unstage->setObjectName("native-query-join-condition-remove");comparisonButtons->addWidget(stage);comparisonButtons->addWidget(unstage);layout->addLayout(comparisonButtons);auto* conditionHint=new QLabel("Vorgemerkte Vergleiche und aktuelles vollständiges Feldpaar bilden die ON-Bedingung (UND).",&dialog);conditionHint->setWordWrap(true);layout->addWidget(conditionHint);
    auto* add=new QPushButton("Verknüpfung hinzufügen",&dialog);add->setObjectName("native-query-join-add");auto* remove=new QPushButton("Letzte Verknüpfung entfernen",&dialog);remove->setObjectName("native-query-join-remove");auto* joinButtons=new QHBoxLayout;joinButtons->addWidget(add);joinButtons->addWidget(remove);auto* discard=new QPushButton("Neue Verknüpfung verwerfen",&dialog);discard->setObjectName("native-query-join-discard");joinButtons->addWidget(discard);layout->addLayout(joinButtons);
    auto* preview=new QPlainTextEdit(&dialog);preview->setObjectName("native-query-join-sql");preview->setReadOnly(true);preview->setMinimumHeight(80);preview->setMaximumHeight(150);layout->addWidget(preview,1);auto* status=new QLabel(&dialog);status->setWordWrap(true);status->setTextFormat(Qt::PlainText);layout->addWidget(status);auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);buttons->button(QDialogButtonBox::Ok)->setObjectName("native-query-joins-apply");layout->addWidget(buttons);QObject connections;std::optional<QJsonArray> result;
    const auto render=[&]{comparisons->setRowCount(staged.size());for(int row=0;row<staged.size();++row){const auto entry=staged.at(row).toObject();int column=0;for(const auto* key:{"left_table","left_field","operator","right_field"})comparisons->setItem(row,column++,new QTableWidgetItem(entry.value(key).toString()));}stage->setEnabled(staged.size()<16);unstage->setEnabled(!staged.isEmpty());};
    const auto current=[&]{if(left->currentIndex()<0 || right->currentIndex()<0 || leftField->currentIndex()<0 || rightField->currentIndex()<0)throw VbaRuntimeError(5,"Bitte beide Quellen und Felder auswählen.");return QJsonObject{{"left_table",left->currentText()},{"left_field",leftField->currentText()},{"right_field",rightField->currentText()},{"operator",operation->currentText()}};};
    const auto fields=[&](QComboBox* source,QComboBox* target){target->clear();if(source->currentIndex()<0)return;try{checkProject();if(source==left)target->addItems(nativeQuerySources(database,main,joins).fields.value(source->currentText()));else target->addItems(accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quote(source->currentText()))));}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}};
    const auto refresh=[&]{checkProject();const auto resolved=nativeQuerySources(database,main,joins);staged=QJsonArray{};render();operation->setCurrentIndex(0);const QSignalBlocker blockLeft(left),blockRight(right);left->clear();right->clear();left->addItems(resolved.order);for(int row=0;row<left->count();++row)left->setItemData(row,resolved.origins.value(left->itemText(row)),Qt::ToolTipRole);right->addItems(sourceNames(database));right->setCurrentIndex(-1);alias->clear();fields(left,leftField);fields(right,rightField);preview->setPlainText(resolved.sql);remove->setEnabled(!joins.isEmpty());add->setEnabled(joins.size()<8 && right->count()>0);};
    QObject::connect(left,&QComboBox::currentIndexChanged,&connections,[&]{fields(left,leftField);});QObject::connect(right,&QComboBox::currentIndexChanged,&connections,[&]{staged=QJsonArray{};render();fields(right,rightField);});
    QObject::connect(stage,&QPushButton::clicked,&connections,[&]{try{checkProject();if(staged.size()>=16)throw VbaRuntimeError(3251,"Höchstens16 Feldvergleiche.");const auto entry=current();if(staged.contains(entry))throw VbaRuntimeError(5,"Dieser Feldvergleich ist bereits vorgemerkt.");staged.append(entry);leftField->setCurrentIndex(-1);rightField->setCurrentIndex(-1);status->clear();render();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    QObject::connect(unstage,&QPushButton::clicked,&connections,[&]{if(!staged.isEmpty())staged.removeLast();render();});
    QObject::connect(add,&QPushButton::clicked,&connections,[&]{try{checkProject();if(joins.size()>=8)throw VbaRuntimeError(3251,"Höchstens acht Verknüpfungen.");if(right->currentIndex()<0)throw VbaRuntimeError(5,"Bitte eine neue Quelle auswählen.");auto conditions=staged;if(leftField->currentIndex()>=0 || rightField->currentIndex()>=0){const auto entry=current();if(!conditions.contains(entry))conditions.append(entry);}if(conditions.isEmpty() || conditions.size()>16)throw VbaRuntimeError(3251,"Eine Verknüpfung benötigt1..16 Feldvergleiche.");auto edited=joins;const auto first=conditions.first().toObject();edited.append(QJsonObject{{"table",right->currentText()},{"left_table",first.value("left_table")},{"left_field",first.value("left_field")},{"right_field",first.value("right_field")},{"kind",kind->currentText()},{"alias",alias->text().trimmed()},{"conditions",conditions}});(void)nativeQuerySources(database,main,edited);joins=edited;status->clear();refresh();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    QObject::connect(discard,&QPushButton::clicked,&connections,[&]{try{refresh();status->clear();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    QObject::connect(remove,&QPushButton::clicked,&connections,[&]{try{checkProject();if(!joins.isEmpty())joins.removeLast();refresh();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    QObject::connect(buttons,&QDialogButtonBox::rejected,&connections,[&]{dialog.reject();});QObject::connect(buttons,&QDialogButtonBox::accepted,&connections,[&]{try{checkProject();if(right->currentIndex()>=0 || !staged.isEmpty())throw VbaRuntimeError(5,"Die neue Verknüpfung zuerst hinzufügen oder verwerfen.");(void)nativeQuerySources(database,main,joins);result=joins;dialog.accept();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    refresh();dialog.resize(800,700);if(dialog.exec()!=QDialog::Accepted)return {};return result;
}
}
