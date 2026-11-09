#include "mainwindow.hpp"
#include "native_query_builder.hpp"
#include "access_query.hpp"
#include <QAction>
#include <QApplication>
#include <QComboBox>
#include <QDialog>
#include <QJsonDocument>
#include <QJsonArray>
#include <QInputDialog>
#include <QLabel>
#include <QLineEdit>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QSqlDriver>
#include <QSqlQuery>
#include <QTableWidget>
#include <QTableView>
#include <QTimer>
#include <sqlite3.h>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
namespace {
QAction* createQuery(MainWindow& window) {for(auto* action:window.findChildren<QAction*>())if(action->text()=="+ Abfrage")return action;assert(false);return nullptr;}
QDialog* builder() {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-query-builder");return dialog;}
int fieldRow(QTableWidget* fields,const QString& name) {for(int row=0;row<fields->rowCount();++row)if(fields->item(row,1)->text()==name)return row;assert(false);return -1;}
}
void exerciseVisualQueryDesigner(MainWindow& window,QSqlDatabase database) {
    QSqlQuery sql(database);assert(sql.exec("CREATE TABLE builder_input(id INTEGER4,label TEXT,score REAL)"));
    assert(sql.exec("INSERT INTO builder_input VALUES(1,'rowA',2),(2,'rowB',4),(3,'rowC',6)"));sql.finish();
    const auto handle=database.driver()->handle();assert(handle.isValid() && QByteArray(handle.typeName())=="sqlite3*");auto* native=*static_cast<sqlite3* const*>(handle.constData());int reads=0;
    assert(sqlite3_create_function_v2(native,"BuilderProbe",0,SQLITE_UTF8,&reads,[](sqlite3_context* context,int,sqlite3_value**){++*static_cast<int*>(sqlite3_user_data(context));sqlite3_result_int(context,4);},nullptr,nullptr,nullptr)==SQLITE_OK);
    QString generated;
    QTimer::singleShot(0,&window,[&] {
        auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-query-designer");auto* source=editor->findChild<QPlainTextEdit*>("native-query-sql");auto* design=editor->findChild<QPushButton*>("native-query-design");assert(source && design);
        source->setPlainText("SELECT id FROM builder_input");const auto original=source->toPlainText();
        QTimer::singleShot(0,editor,[&]{auto* dialog=builder();dialog->findChild<QComboBox*>("native-query-builder-source")->setCurrentText("builder_input");assert(reads==0);dialog->reject();});design->click();assert(source->toPlainText()==original);
        QTimer::singleShot(0,editor,[&]{
            auto* dialog=builder();auto* tables=dialog->findChild<QComboBox*>("native-query-builder-source");auto* fields=dialog->findChild<QTableWidget*>("native-query-builder-fields");auto* apply=dialog->findChild<QPushButton*>("native-query-builder-apply");assert(tables && fields && apply && !apply->isEnabled());
            tables->setCurrentText("builder_input");assert(fields->rowCount()==3 && apply->isEnabled());
            const auto label=fieldRow(fields,"label");fields->item(label,2)->setText("shown");fields->setCurrentCell(label,2);dialog->findChild<QPushButton*>("native-query-builder-up")->click();assert(fields->item(0,1)->text()=="label");
            const auto id=fieldRow(fields,"id"),score=fieldRow(fields,"score");qobject_cast<QComboBox*>(fields->cellWidget(id,3))->setCurrentIndex(2);fields->item(score,0)->setCheckState(Qt::Unchecked);
            fields->item(score,4)->setText(">= BuilderProbe()");assert(apply->isEnabled() && reads==0);
            fields->item(id,2)->setText("shown");assert(!apply->isEnabled());fields->item(id,2)->setText("");assert(apply->isEnabled());
            fields->item(score,4)->setText(">= 1; DELETE FROM builder_input");assert(!apply->isEnabled() && reads==0);fields->item(score,4)->setText(">= BuilderProbe()");assert(apply->isEnabled());
            apply->click();
        });design->click();generated=source->toPlainText();assert(generated.contains("BuilderProbe()") && generated.contains(" AS [shown]") && reads==0);
        auto* rows=editor->findChild<QTableView*>();assert(rows && rows->model()->rowCount()==0);editor->findChild<QPushButton*>("native-query-run")->click();assert(rows->model()->rowCount()==2 && reads==3 && rows->model()->headerData(0,Qt::Horizontal).toString()=="shown");
        QSqlQuery reference(database);assert(reference.exec("SELECT label,id FROM builder_input WHERE score>=4 ORDER BY id DESC"));for(int row=0;row<2;++row){assert(reference.next());for(int column=0;column<2;++column)assert(rows->model()->index(row,column).data().toString()==reference.value(column).toString());}assert(!reference.next());reference.finish();
        editor->findChild<QLineEdit*>("native-query-name")->setText("visual_probe");editor->findChild<QPushButton*>("native-query-save")->click();assert(editor->result()==QDialog::Accepted);
    });createQuery(window)->trigger();
    sql.prepare("SELECT definition FROM _la_objects WHERE kind='query' AND name='visual_probe'");assert(sql.exec() && sql.next());const auto storedBytes=sql.value(0).toByteArray();const auto stored=QJsonDocument::fromJson(storedBytes).object();sql.finish();assert(stored.value("sql").toString()==generated && stored.value("visual_query").isObject());
    QTimer::singleShot(0,&window,[&]{auto* dialog=builder();assert(dialog->findChild<QComboBox*>("native-query-builder-source")->currentText()=="builder_input");auto* fields=dialog->findChild<QTableWidget*>("native-query-builder-fields");assert(fields->item(0,1)->text()=="label" && fields->item(0,2)->text()=="shown" && fields->item(fieldRow(fields,"score"),0)->checkState()==Qt::Unchecked);dialog->reject();});
    assert(!linaccess::editNativeQueryDesign(&window,database,stored.value("visual_query").toObject().value("design").toObject(),[]{}));assert(reads==3);
    QTimer::singleShot(0,&window,[&]{
        auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor);
        QTimer::singleShot(0,editor,[&]{
            auto* dialog=builder();dialog->findChild<QComboBox*>("native-query-builder-source")->setCurrentText("builder_input");
            QTimer::singleShot(0,dialog,[]{
                auto* params=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(params && params->objectName()=="native-query-parameters");auto* fields=params->findChild<QTableWidget*>("native-query-parameter-fields");auto* add=params->findChild<QPushButton*>("native-query-parameter-add");assert(fields && add);
                add->click();add->click();fields->item(0,0)->setText("floor");fields->item(1,0)->setText("ceiling");
                qobject_cast<QComboBox*>(fields->cellWidget(0,1))->setCurrentText("Long");qobject_cast<QComboBox*>(fields->cellWidget(1,1))->setCurrentText("Long");params->findChild<QPushButton*>("native-query-parameter-apply")->click();
            });dialog->findChild<QPushButton*>("native-query-builder-parameters")->click();
            dialog->findChild<QPushButton*>("native-query-builder-or-add")->click();auto* fields=dialog->findChild<QTableWidget*>("native-query-builder-fields");assert(fields->columnCount()==6);
            fields->item(fieldRow(fields,"score"),4)->setText(">= [floor]");fields->item(fieldRow(fields,"score"),5)->setText("<= [ceiling]");qobject_cast<QComboBox*>(fields->cellWidget(fieldRow(fields,"id"),3))->setCurrentIndex(2);
            auto* apply=dialog->findChild<QPushButton*>("native-query-builder-apply");assert(apply->isEnabled() && reads==3);apply->click();
        });editor->findChild<QPushButton*>("native-query-design")->click();
        auto* text=editor->findChild<QPlainTextEdit*>("native-query-sql");assert(text->toPlainText().startsWith("PARAMETERS ") && text->toPlainText().contains(" OR "));
        QTimer::singleShot(0,editor,[&]{
            auto* dialog=builder();auto* fields=dialog->findChild<QTableWidget*>("native-query-builder-fields");assert(fields->columnCount()==6 && fields->item(fieldRow(fields,"score"),5)->text()=="<= [ceiling]");
            QTimer::singleShot(0,dialog,[]{auto* params=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(params && params->objectName()=="native-query-parameters");auto* fields=params->findChild<QTableWidget*>("native-query-parameter-fields");assert(fields->rowCount()==2 && fields->item(0,0)->text()=="floor" && fields->item(1,0)->text()=="ceiling");params->reject();});dialog->findChild<QPushButton*>("native-query-builder-parameters")->click();dialog->reject();
        });editor->findChild<QPushButton*>("native-query-design")->click();assert(reads==3);
        QTimer::singleShot(0,editor,[&]{auto* prompt=qobject_cast<QInputDialog*>(QApplication::activeModalWidget());assert(prompt);prompt->setTextValue("6");QTimer::singleShot(0,editor,[]{auto* second=qobject_cast<QInputDialog*>(QApplication::activeModalWidget());assert(second);second->setTextValue("2");second->accept();});prompt->accept();});
        editor->findChild<QPushButton*>("native-query-run")->click();auto* result=editor->findChild<QTableView*>();assert(result && result->model()->rowCount()==2);
        QSqlQuery reference(database);assert(reference.exec("SELECT id,label,score FROM builder_input WHERE score>=6 OR score<=2 ORDER BY id DESC"));for(int row=0;row<2;++row){assert(reference.next());for(int column=0;column<3;++column)assert(reference.value(column).toString()==result->model()->index(row,column).data().toString());}assert(!reference.next());reference.finish();
        text->setPlainText("SELECT id FROM builder_input WHERE id=1");editor->findChild<QLineEdit*>("native-query-name")->setText("visual_manual");editor->findChild<QPushButton*>("native-query-save")->click();
    });createQuery(window)->trigger();
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='query' AND name='visual_manual'") && sql.next() && !QJsonDocument::fromJson(sql.value(0).toByteArray()).object().contains("visual_query"));sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM builder_input") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='query' AND name IN ('visual_probe','visual_manual')"));assert(linaccess::installAccessQueryViews(database)>=0);assert(sql.exec("DROP TABLE builder_input"));sql.finish();
    assert(sqlite3_create_function_v2(native,"BuilderProbe",0,SQLITE_UTF8,nullptr,nullptr,nullptr,nullptr,nullptr)==SQLITE_OK);
}
