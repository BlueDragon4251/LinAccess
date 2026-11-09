#include "native_query_sources.hpp"
#include "access_query.hpp"
#include "vba_runtime.hpp"
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QJsonObject>
#include <QLabel>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QSet>
#include <QSignalBlocker>
#include <QVBoxLayout>
namespace linaccess {
namespace {
QString quote(QString name) {
    if(name.isEmpty() || name.size()>64 || !name.isValidUtf16() || name.contains(QChar(0)))throw VbaRuntimeError(5,"Ungültiger Quellen- oder Feldname.");
    name.replace(']',"]]");return '['+name+']';
}
QStringList sourceNames(QSqlDatabase database) {
    auto names=database.tables(static_cast<QSql::TableType>(QSql::Tables|QSql::Views));
    names.removeDuplicates();names.removeIf([](const auto& name){return name.startsWith("_la_",Qt::CaseInsensitive) || name.startsWith("sqlite_",Qt::CaseInsensitive);});names.sort(Qt::CaseInsensitive);return names;
}
}
NativeQuerySources nativeQuerySources(QSqlDatabase database,const QString& main,const QJsonArray& joins) {
    if(joins.size()>8)throw VbaRuntimeError(3251,"Entwurf unterstützt höchstens acht Verknüpfungen.");
    NativeQuerySources result;const auto available=sourceNames(database);QSet<QString> used;
    const auto add=[&](const QString& name){if(!available.contains(name))throw VbaRuntimeError(3078,"Verknüpfte Datenquelle fehlt.");if(used.contains(name.toCaseFolded()))throw VbaRuntimeError(3251,"Jede Quelle darf einmal vorkommen; Selbstverknüpfungen benötigen SQL-Aliase im SQL-Editor.");used.insert(name.toCaseFolded());result.order.append(name);result.fields.insert(name,accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quote(name))));};
    add(main);result.sql=quote(main);
    for(const auto& value:joins) {
        if(!value.isObject())throw VbaRuntimeError(5,"Ungültiger Verknüpfungsentwurf.");const auto join=value.toObject();
        for(const auto* key:{"table","left_table","left_field","right_field","kind"})if(!join.value(key).isString())throw VbaRuntimeError(5,"Ungültige Verknüpfungseigenschaft.");
        const auto table=join.value("table").toString(),left=join.value("left_table").toString(),leftField=join.value("left_field").toString(),rightField=join.value("right_field").toString(),kind=join.value("kind").toString();
        if(!result.fields.contains(left) || !result.fields.value(left).contains(leftField) || (kind!="INNER" && kind!="LEFT"))throw VbaRuntimeError(5,"Ungültige Verknüpfungsquelle, Feld oder Verknüpfungsart.");
        add(table);if(!result.fields.value(table).contains(rightField))throw VbaRuntimeError(3265,"Verknüpfungsfeld fehlt.");
        result.sql='('+result.sql+' '+kind+" JOIN "+quote(table)+" ON "+quote(left)+'.'+quote(leftField)+" = "+quote(table)+'.'+quote(rightField)+')';
    }
    return result;
}
std::optional<QJsonArray> editNativeQueryJoins(QWidget* parent,QSqlDatabase database,const QString& main,const QJsonArray& initial,const std::function<void()>& checkProject) {
    checkProject();(void)nativeQuerySources(database,main,initial);QJsonArray joins=initial;QDialog dialog(parent);dialog.setObjectName("native-query-joins");dialog.setWindowTitle("Datenquellen verknüpfen");auto* layout=new QVBoxLayout(&dialog);auto* form=new QFormLayout;layout->addLayout(form);
    auto* left=new QComboBox(&dialog);left->setObjectName("native-query-join-left-table");auto* leftField=new QComboBox(&dialog);leftField->setObjectName("native-query-join-left-field");auto* right=new QComboBox(&dialog);right->setObjectName("native-query-join-right-table");auto* rightField=new QComboBox(&dialog);rightField->setObjectName("native-query-join-right-field");auto* kind=new QComboBox(&dialog);kind->setObjectName("native-query-join-kind");kind->addItems({"INNER","LEFT"});
    form->addRow("Vorhandene Quelle",left);form->addRow("Feld",leftField);form->addRow("Neue Quelle",right);form->addRow("Feld",rightField);form->addRow("Verknüpfungsart",kind);layout->addWidget(new QLabel("INNER: passende Zeilen. LEFT: alle Zeilen der bisherigen Quellen, rechts bei fehlender Zuordnung Null. Vergleich: Feld = Feld.",&dialog));
    auto* add=new QPushButton("Verknüpfung hinzufügen",&dialog);add->setObjectName("native-query-join-add");auto* remove=new QPushButton("Letzte Verknüpfung entfernen",&dialog);remove->setObjectName("native-query-join-remove");layout->addWidget(add);layout->addWidget(remove);
    auto* preview=new QPlainTextEdit(&dialog);preview->setObjectName("native-query-join-sql");preview->setReadOnly(true);layout->addWidget(preview);auto* status=new QLabel(&dialog);status->setWordWrap(true);status->setTextFormat(Qt::PlainText);layout->addWidget(status);auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);buttons->button(QDialogButtonBox::Ok)->setObjectName("native-query-joins-apply");layout->addWidget(buttons);QObject connections;std::optional<QJsonArray> result;
    const auto fields=[&](QComboBox* source,QComboBox* target){target->clear();if(source->currentIndex()<0)return;try{checkProject();target->addItems(accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quote(source->currentText()))));}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}};
    const auto refresh=[&]{checkProject();const auto resolved=nativeQuerySources(database,main,joins);const QSignalBlocker blockLeft(left),blockRight(right);left->clear();right->clear();left->addItems(resolved.order);for(const auto& name:sourceNames(database))if(!resolved.order.contains(name,Qt::CaseInsensitive))right->addItem(name);right->setCurrentIndex(-1);fields(left,leftField);fields(right,rightField);preview->setPlainText(resolved.sql);remove->setEnabled(!joins.isEmpty());add->setEnabled(joins.size()<8 && right->count()>0);};
    QObject::connect(left,&QComboBox::currentIndexChanged,&connections,[&]{fields(left,leftField);});QObject::connect(right,&QComboBox::currentIndexChanged,&connections,[&]{fields(right,rightField);});
    QObject::connect(add,&QPushButton::clicked,&connections,[&]{try{checkProject();if(joins.size()>=8)throw VbaRuntimeError(3251,"Höchstens acht Verknüpfungen.");if(left->currentIndex()<0 || right->currentIndex()<0 || leftField->currentIndex()<0 || rightField->currentIndex()<0)throw VbaRuntimeError(5,"Bitte beide Quellen und Felder auswählen.");auto edited=joins;edited.append(QJsonObject{{"table",right->currentText()},{"left_table",left->currentText()},{"left_field",leftField->currentText()},{"right_field",rightField->currentText()},{"kind",kind->currentText()}});(void)nativeQuerySources(database,main,edited);joins=edited;status->clear();refresh();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    QObject::connect(remove,&QPushButton::clicked,&connections,[&]{try{checkProject();if(!joins.isEmpty())joins.removeLast();refresh();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    QObject::connect(buttons,&QDialogButtonBox::rejected,&connections,[&]{dialog.reject();});QObject::connect(buttons,&QDialogButtonBox::accepted,&connections,[&]{try{checkProject();(void)nativeQuerySources(database,main,joins);result=joins;dialog.accept();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));}});
    refresh();dialog.resize(740,620);if(dialog.exec()!=QDialog::Accepted)return {};return result;
}
}
