#include "native_query_builder.hpp"
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
    if(!db.tables(QSql::Tables).contains(table)) throw VbaRuntimeError(3078,"Quelltabelle ist nicht verfügbar.");
    const auto fields=accessQueryColumns(db,compileAccessQuery("SELECT * FROM "+quoted(table)));
    const auto rows=design.value("fields").toArray();if(rows.isEmpty() || rows.size()>255) throw VbaRuntimeError(3251,"Abfrageentwurf benötigt1..255 Felder.");
    QStringList projection,predicates,sort;QSet<QString> used,labels;
    const bool distinct=design.value("distinct").toBool();
    static const QRegularExpression operatorPrefix("^(?:[=<>]|(?:Not\\s+)?Like\\b|Is\\s+(?:Not\\s+)?Null\\b|Between\\b|In\\s*\\()",QRegularExpression::CaseInsensitiveOption);
    static const QRegularExpression literalPrefix("^(?:['\"#]|[+-]?(?:[0-9]|\\.[0-9])|True$|False$)",QRegularExpression::CaseInsensitiveOption);
    for(const auto& value:rows) {
        if(!value.isObject()) throw VbaRuntimeError(5,"Ungültiger Feldentwurf.");const auto row=value.toObject();
        const auto field=row.value("name").toString();if(!fields.contains(field) || used.contains(field.toCaseFolded())) throw VbaRuntimeError(5,"Ein Feld fehlt oder wurde mehrfach aufgenommen.");used.insert(field.toCaseFolded());
        const auto reference=quoted(table)+'.'+quoted(field);const bool visible=row.value("visible").toBool();const auto alias=row.value("alias").toString().trimmed();
        if(visible) {
            const auto label=alias.isEmpty() ? field : alias;if(labels.contains(label.toCaseFolded())) throw VbaRuntimeError(5,"Ausgabefelder benötigen eindeutige Namen.");labels.insert(label.toCaseFolded());
            projection.append(reference+(alias.isEmpty() ? QString{} : " AS "+quoted(alias)));
        }
        const auto direction=row.value("sort").toInt();if(direction<0 || direction>2) throw VbaRuntimeError(5,"Ungültige Sortierung.");
        if(direction) {if(distinct && !visible) throw VbaRuntimeError(3251,"Bei eindeutigen Zeilen muss ein Sortierfeld auch ausgegeben werden.");sort.append(reference+(direction==1 ? " ASC" : " DESC"));}
        auto criterion=row.value("criteria").toString().trimmed();
        if(criterion.size()>4096) throw VbaRuntimeError(5,"Kriterium ist zu lang.");
        if(!criterion.isEmpty()) {
            if(!operatorPrefix.match(criterion).hasMatch()) {if(!literalPrefix.match(criterion).hasMatch()) throw VbaRuntimeError(5,"Kriterium benötigt einen Vergleich oder ein Literal.");criterion="= "+criterion;}
            const auto predicate=reference+' '+criterion;(void)accessWhereClause(predicate);predicates.append('('+predicate+')');
        }
    }
    if(projection.isEmpty()) throw VbaRuntimeError(5,"Mindestens ein Feld zur Ausgabe auswählen.");
    auto sql=QString(distinct ? "SELECT DISTINCT " : "SELECT ")+projection.join(", ")+" FROM "+quoted(table);
    if(!predicates.isEmpty()) {const auto predicate=predicates.join(" AND ");(void)accessWhereClause(predicate);sql+=" WHERE "+predicate;}
    if(!sort.isEmpty()) sql+=" ORDER BY "+sort.join(", ");
    const auto plan=compileAccessQuery(sql);if(plan.daoQueryType!=0 || !plan.makeTable.isEmpty() || !plan.parameters.isEmpty()) throw VbaRuntimeError(3251,"Entwurf muss eine parameterfreie Auswahlabfrage sein.");
    (void)accessQueryColumns(db,plan);return sql;
}
}
std::optional<NativeQueryDesign> editNativeQueryDesign(QWidget* parent,QSqlDatabase database,const QJsonObject& initial,const std::function<void()>& checkProject) {
    checkProject();QDialog dialog(parent);dialog.setObjectName("native-query-builder");dialog.setWindowTitle("Abfrageentwurf");dialog.resize(900,650);auto* layout=new QVBoxLayout(&dialog);
    auto* source=new QComboBox(&dialog);source->setObjectName("native-query-builder-source");source->setPlaceholderText("Quelltabelle auswählen");source->setAccessibleName("Quelltabelle");
    auto tables=database.tables(QSql::Tables);std::sort(tables.begin(),tables.end(),[](const auto& a,const auto& b){return a.compare(b,Qt::CaseInsensitive)<0;});
    for(const auto& table:tables) if(!table.startsWith("_la_",Qt::CaseInsensitive) && !table.startsWith("sqlite_",Qt::CaseInsensitive)) source->addItem(table);
    source->setCurrentIndex(-1);layout->addWidget(source);
    auto* distinct=new QCheckBox("Nur eindeutige Zeilen",&dialog);distinct->setObjectName("native-query-builder-distinct");layout->addWidget(distinct);
    auto* grid=new QTableWidget(0,5,&dialog);grid->setObjectName("native-query-builder-fields");grid->setHorizontalHeaderLabels({"Ausgabe","Feld","Alias","Sortierung","Kriterium"});grid->setItemDelegate(new BuilderDelegate(grid));grid->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);grid->horizontalHeader()->setStretchLastSection(true);layout->addWidget(grid,1);
    auto* rows=new QHBoxLayout;auto* up=new QPushButton("Feld nach oben",&dialog);auto* down=new QPushButton("Feld nach unten",&dialog);auto* reload=new QPushButton("Quelle neu einlesen",&dialog);up->setObjectName("native-query-builder-up");down->setObjectName("native-query-builder-down");reload->setObjectName("native-query-builder-reload");rows->addWidget(up);rows->addWidget(down);rows->addStretch();rows->addWidget(reload);layout->addLayout(rows);
    auto* preview=new QPlainTextEdit(&dialog);preview->setObjectName("native-query-builder-sql");preview->setReadOnly(true);preview->setMaximumHeight(130);layout->addWidget(preview);
    auto* status=new QLabel(&dialog);status->setObjectName("native-query-builder-status");status->setTextFormat(Qt::PlainText);status->setWordWrap(true);layout->addWidget(status);
    auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);auto* apply=buttons->button(QDialogButtonBox::Ok);apply->setText("Übernehmen");apply->setObjectName("native-query-builder-apply");layout->addWidget(buttons);
    bool loading=false;std::optional<NativeQueryDesign> result;QObject connections;
    const auto capture=[&] {QJsonArray fields;for(int row=0;row<grid->rowCount();++row) fields.append(QJsonObject{{"name",grid->item(row,1)->text()},{"visible",grid->item(row,0)->checkState()==Qt::Checked},{"alias",grid->item(row,2)->text()},{"sort",qobject_cast<QComboBox*>(grid->cellWidget(row,3))->currentIndex()},{"criteria",grid->item(row,4)->text()}});return QJsonObject{{"version",1},{"table",source->currentText()},{"distinct",distinct->isChecked()},{"fields",fields}};};
    const auto refresh=[&] {if(loading)return;if(source->currentIndex()<0){preview->clear();status->setText("Bitte eine Quelltabelle auswählen.");apply->setEnabled(false);return;}try{checkProject();const auto sql=designSql(database,capture());checkProject();preview->setPlainText(sql);status->setText("Entwurf gültig. Übernehmen aktualisiert den SQL-Editor.");apply->setEnabled(true);}catch(const std::exception& error){preview->clear();status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}};
    const auto append=[&](const QJsonObject& field) {const int row=grid->rowCount();grid->insertRow(row);auto* output=new QTableWidgetItem;output->setFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable|Qt::ItemIsUserCheckable);output->setCheckState(field.value("visible").toBool(true) ? Qt::Checked : Qt::Unchecked);grid->setItem(row,0,output);auto* name=new QTableWidgetItem(field.value("name").toString());name->setFlags(Qt::ItemIsEnabled|Qt::ItemIsSelectable);grid->setItem(row,1,name);grid->setItem(row,2,new QTableWidgetItem(field.value("alias").toString()));auto* order=new QComboBox(grid);order->addItems({"(keine)","Aufsteigend","Absteigend"});order->setCurrentIndex(field.value("sort").toInt());grid->setCellWidget(row,3,order);grid->setItem(row,4,new QTableWidgetItem(field.value("criteria").toString()));QObject::connect(order,&QComboBox::currentIndexChanged,&connections,refresh);};
    const auto populate=[&](bool restore) {
        loading=true;grid->setRowCount(0);
        try {
            checkProject();const auto table=source->currentText();const auto fields=accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quoted(table)));
            if(fields.size()>255)throw VbaRuntimeError(3251,"Quelle hat mehr als255 Felder.");QSet<QString> added;
            if(restore && (!initial.value("fields").isArray() || initial.value("fields").toArray().size()>255))throw VbaRuntimeError(3251,"Gespeicherter Feldentwurf überschreitet den unterstützten Umfang.");
            if(restore) for(const auto& value:initial.value("fields").toArray()) {if(!value.isObject())throw VbaRuntimeError(5,"Ungültiger gespeicherter Feldentwurf.");const auto field=value.toObject();if(!field.value("name").isString() || !field.value("visible").isBool() || !field.value("alias").isString() || !field.value("sort").isDouble() || field.value("sort").toDouble()!=field.value("sort").toInt() || !field.value("criteria").isString())throw VbaRuntimeError(5,"Ungültige gespeicherte Feldeigenschaften.");if(field.value("alias").toString().size()>64 || field.value("criteria").toString().size()>4096)throw VbaRuntimeError(5,"Gespeicherter Alias oder Vergleich ist zu lang.");const auto name=field.value("name").toString();if(!fields.contains(name) || added.contains(name))throw VbaRuntimeError(5,"Gespeichertes Feld fehlt; Quelle neu einlesen.");append(field);added.insert(name);}
            for(const auto& field:fields)if(!added.contains(field))append({{"name",field},{"visible",!restore}});
            distinct->setChecked(restore && initial.value("distinct").toBool());loading=false;refresh();
        } catch(const std::exception& error) {loading=false;preview->clear();status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}
    };
    QObject::connect(source,&QComboBox::currentIndexChanged,&connections,[&]{populate(false);});QObject::connect(grid,&QTableWidget::itemChanged,&connections,refresh);QObject::connect(distinct,&QCheckBox::toggled,&connections,refresh);QObject::connect(reload,&QPushButton::clicked,&connections,[&]{populate(false);});
    const auto move=[&](int delta){const int row=grid->currentRow(),target=row+delta;if(row<0 || target<0 || target>=grid->rowCount())return;auto fields=capture().value("fields").toArray();const QJsonValue first=fields[row];fields[row]=fields[target];fields[target]=first;loading=true;grid->setRowCount(0);for(const auto& field:fields)append(field.toObject());grid->setCurrentCell(target,2);loading=false;refresh();};QObject::connect(up,&QPushButton::clicked,&connections,[&]{move(-1);});QObject::connect(down,&QPushButton::clicked,&connections,[&]{move(1);});
    QObject::connect(buttons,&QDialogButtonBox::rejected,&connections,[&]{dialog.reject();});QObject::connect(buttons,&QDialogButtonBox::accepted,&connections,[&]{apply->setFocus();try{checkProject();auto state=capture();auto sql=designSql(database,state);checkProject();result=NativeQueryDesign{sql,state};dialog.accept();}catch(const std::exception& error){status->setText(QString::fromUtf8(error.what()));apply->setEnabled(false);}});
    if(initial.value("version").toInt()==1 && source->findText(initial.value("table").toString())>=0) {const QSignalBlocker blocked(source);source->setCurrentIndex(source->findText(initial.value("table").toString()));populate(true);}else refresh();
    if(dialog.exec()!=QDialog::Accepted)return {};return result;
}
}
