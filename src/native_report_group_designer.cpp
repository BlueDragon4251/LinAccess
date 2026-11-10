#include "native_report_group_designer.hpp"
#include "native_report_groups.hpp"
#include "access_query.hpp"
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QHeaderView>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QTableWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
namespace linaccess {
std::optional<QJsonArray> editNativeReportGroups(QWidget* owner,QSqlDatabase database,const QString& source,
    const QJsonArray& initial,const std::function<void()>& guard) {
    guard();const auto plan=compileAccessQuery(source);
    if(!plan.makeTable.isEmpty() || !plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive)) throw VbaRuntimeError(3251,"Report grouping requires a read-only SELECT source");
    if(installAccessQueryViews(database)<0) throw VbaRuntimeError(3251,"Report source views are busy");
    QVector<QString> types;const auto sourceColumns=accessQueryColumns(database,plan,&types);QStringList columns;
    for(int column=0;column<sourceColumns.size();++column) if(types.value(column).compare("BLOB",Qt::CaseInsensitive)!=0 && !types.value(column).contains("BINARY",Qt::CaseInsensitive)) columns.append(sourceColumns[column]);guard();
    if(columns.isEmpty() || columns.size()>32 || initial.size()>10) throw VbaRuntimeError(3251,"Report grouping source/level bound exceeded");
    QDialog dialog(owner);dialog.setObjectName("native-report-group-designer");dialog.setWindowTitle("Bericht gruppieren und sortieren");dialog.resize(900,420);
    auto* layout=new QVBoxLayout(&dialog);layout->addWidget(new QLabel("Die Reihenfolge der Zeilen bestimmt die Gruppenstufen. Berechnungen im Gruppenfuß gelten jeweils für diese Gruppe.",&dialog));
    auto* table=new QTableWidget(0,5,&dialog);table->setObjectName("native-report-groups-list");table->setHorizontalHeaderLabels({"Feld","Sortierung","Gruppenkopf","Berechnung im Gruppenfuß (optional)","Kopf auf Folgeseiten"});table->horizontalHeader()->setSectionResizeMode(3,QHeaderView::Stretch);table->horizontalHeader()->setSectionResizeMode(4,QHeaderView::ResizeToContents);table->setSelectionBehavior(QAbstractItemView::SelectRows);layout->addWidget(table,1);
    auto* actions=new QHBoxLayout;
    auto button=[&](const QString& title,const QString& name) {auto* item=new QPushButton(title,&dialog);item->setObjectName(name);actions->addWidget(item);return item;};
    auto* add=button("Hinzufügen","native-report-group-add");auto* remove=button("Entfernen","native-report-group-remove");auto* up=button("Nach oben","native-report-group-up");auto* down=button("Nach unten","native-report-group-down");actions->addStretch();layout->addLayout(actions);
    auto* status=new QLabel(&dialog);status->setWordWrap(true);layout->addWidget(status);
    const auto rowData=[&](int row) {return QJsonObject{{"field",qobject_cast<QComboBox*>(table->cellWidget(row,0))->currentText()},
        {"descending",qobject_cast<QComboBox*>(table->cellWidget(row,1))->currentIndex()==1},
        {"header",qobject_cast<QCheckBox*>(table->cellWidget(row,2))->isChecked()},
        {"footer_expression",qobject_cast<QLineEdit*>(table->cellWidget(row,3))->text().trimmed()},
        {"repeat_header",qobject_cast<QCheckBox*>(table->cellWidget(row,4))->isChecked()}};};
    const auto fill=[&](int row,const QJsonObject& item) {
        auto* field=new QComboBox(table);field->addItems(columns);const auto wanted=item.value("field").toString();int selected=-1;for(int index=0;index<columns.size();++index) if(columns[index].compare(wanted,Qt::CaseInsensitive)==0) selected=index;field->setCurrentIndex(wanted.isEmpty() ? 0 : selected);table->setCellWidget(row,0,field);
        auto* sort=new QComboBox(table);sort->addItems({"Aufsteigend","Absteigend"});sort->setCurrentIndex(item.value("descending").toBool() ? 1 : 0);table->setCellWidget(row,1,sort);
        auto* header=new QCheckBox(table);header->setAccessibleName("Gruppenkopf");header->setChecked(item.value("header").toBool(true));table->setCellWidget(row,2,header);
        auto* footer=new QLineEdit(item.value("footer_expression").toString(),table);footer->setPlaceholderText("=Count(*) oder =Sum([Feld])");table->setCellWidget(row,3,footer);
        auto* repeat=new QCheckBox(table);repeat->setAccessibleName("Gruppenkopf auf Folgeseiten wiederholen");repeat->setToolTip("Bei mehrseitigen Gruppen den Gruppenkopf auf jeder Folgeseite erneut anzeigen.");repeat->setChecked(item.value("repeat_header").toBool());table->setCellWidget(row,4,repeat);

    };
    for(const auto& value:initial) {if(!value.isObject()) throw VbaRuntimeError(5,"Invalid stored report group");const int row=table->rowCount();table->insertRow(row);fill(row,value.toObject());}
    QObject::connect(add,&QPushButton::clicked,&dialog,[&] {if(table->rowCount()>=10) {status->setText("Höchstens zehn Gruppenstufen sind möglich.");return;}const int row=table->rowCount();table->insertRow(row);fill(row,{});table->setCurrentCell(row,0);});
    QObject::connect(remove,&QPushButton::clicked,&dialog,[&] {if(table->currentRow()>=0) table->removeRow(table->currentRow());});
    const auto move=[&](int direction) {const int row=table->currentRow(),target=row+direction;if(row<0 || target<0 || target>=table->rowCount()) return;const auto a=rowData(row),b=rowData(target);fill(row,b);fill(target,a);table->setCurrentCell(target,0);};
    QObject::connect(up,&QPushButton::clicked,&dialog,[&] {move(-1);});QObject::connect(down,&QPushButton::clicked,&dialog,[&] {move(1);});
    auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);buttons->setObjectName("native-report-group-buttons");layout->addWidget(buttons);std::optional<QJsonArray> result;
    QObject::connect(buttons,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    QObject::connect(buttons,&QDialogButtonBox::accepted,&dialog,[&] {try {guard();QJsonArray rows;for(int row=0;row<table->rowCount();++row) rows.append(rowData(row));(void)nativeReportGroups(rows,columns);guard();result=rows;dialog.accept();}catch(const std::exception& error) {status->setText(QString::fromUtf8(error.what()));}});
    dialog.exec();guard();return result;
}
}
