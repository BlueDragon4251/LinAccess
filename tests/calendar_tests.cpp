#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include "vba_source_codec.hpp"
#include "vba_analyzer.hpp"
#include "vba_qt_value.hpp"
#include <QApplication>
#include <QCalendarWidget>
#include <QTableView>
#include <QLineEdit>
#include <QPushButton>
#include <QMessageBox>
#include <QTimer>
#include <QPointer>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QSqlDriver>
#include <QSqlQuery>
#include <QDateTime>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

void testOriginalCalendar(const QDir& directory,const QJsonArray& objects,QSqlDatabase db) {
    QByteArray filterSource,filterCode,calendarSource,calendarCode;
    QString filterName;
    const auto read=[&](const QJsonObject& object) {
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower()); return bytes;
    };
    for (const auto& entry:objects) {
        const auto object=entry.toObject(); const auto hash=object["sha256"].toString();
        if (hash=="b7a13cd0fb0bf009966c418ac5eaecaba184f4117c40777a8d3149b9f9d5f907") { filterSource=read(object); filterName=object["name"].toString(); }
        if (hash=="fd6af9e988bfaa03d9751c765621e3e63693bc7d13642852afae0f98399ba281") filterCode=read(object);
        if (hash=="c41791f3730a70a0b537b26054893d3561d7a12c764ff291d0d96051ba15fd9b") calendarSource=read(object);
        if (hash=="19aace6611cb21892249ec065d5f305eb862b3fb2669d3e374968a6f2aa0a897") calendarCode=read(object);
    }
    assert(!filterSource.isEmpty() && !filterCode.isEmpty() && !calendarSource.isEmpty() && !calendarCode.isEmpty());
    const auto filterLayout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(filterSource));
    const auto calendarLayout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(calendarSource));
    QString target,calendarControl; QStringList targets;
    for (const auto& control:filterLayout.controls) if (control.type=="TextBox" && control.onDblClick.startsWith('=')) targets.append(control.name);
    assert(targets.size()==2); target=targets.front();
    for (const auto& control:calendarLayout.controls) if (control.type=="CustomControl" && control.className.startsWith("MSCAL.",Qt::CaseInsensitive)) { assert(control.onClick.isEmpty()); calendarControl=control.name; }
    assert(!target.isEmpty() && !calendarControl.isEmpty());
    // Select a genuine midnight value without printing or persisting private data.
    QString dateTable,dateColumn; QDate wanted; int expectedCount=0;
    for (const auto& table:db.tables(QSql::Tables)) {
        if (table.startsWith("_la_") || table.startsWith("sqlite_")) continue;
        const auto types=linaccess::sqliteColumnTypes(db,table);
        for (auto it=types.cbegin();it!=types.cend();++it) {
            if (!it.value().contains("DATE")) continue;
            QSqlQuery rows(db);
            const auto sql="SELECT "+db.driver()->escapeIdentifier(it.key(),QSqlDriver::FieldName)+" FROM "+db.driver()->escapeIdentifier(table,QSqlDriver::TableName);
            assert(rows.exec(sql)); QStringList values;
            while (rows.next()) if (!rows.value(0).isNull()) values.append(rows.value(0).toString());
            for (const auto& value:values) {
                const auto parsed=QDateTime::fromString(value,"yyyy-MM-dd HH:mm:ss");
                if (parsed.isValid() && parsed.time()==QTime(0,0) && parsed.date().year()>=100) {
                    wanted=parsed.date(); dateTable=table; dateColumn=it.key(); expectedCount=values.count(value); break;
                }
            }
            if (wanted.isValid()) break;
        }
        if (wanted.isValid()) break;
    }
    assert(wanted.isValid() && expectedCount>0);
    AccessFormPreview filter(filterName,filterSource,filterCode,db);
    filter.show(); QApplication::setActiveWindow(&filter); QApplication::processEvents();
    auto* input=filter.findChild<QLineEdit*>("access-control-"+target); assert(input && input->isVisible()); input->setFocus(); QApplication::processEvents();
    assert(QApplication::focusWidget()==input);
    int unexpectedDialogs=0; QTimer guard;
    QObject::connect(&guard,&QTimer::timeout,[&] {
        if (auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) { ++unexpectedDialogs; box->accept(); }
    });
    guard.start(10); bool visited=false,timedOut=false; QPointer<AccessFormPreview> popupGuard;
    QTimer deadline; deadline.setSingleShot(true);
    std::shared_ptr<linaccess::VbaObject> closedCalendar;
    QObject::connect(&deadline,&QTimer::timeout,[&] {
        timedOut=true;
        if (popupGuard) {
            const auto click=popupGuard->eventResult("control:"+calendarControl+":click");
            std::cerr<<"Calendar click deadline: bound="<<static_cast<bool>(click)<<", status="<<(click ? static_cast<int>(click->status) : -1)
                <<", error="<<(click ? click->errorNumber : 0)<<", line="<<(click ? click->line : 0)<<"\n";
            popupGuard->reject();
        }
    });
    QTimer choose;
    QObject::connect(&choose,&QTimer::timeout,[&] {
        AccessFormPreview* popup=nullptr;
        for (auto* widget:QApplication::topLevelWidgets())
            if (auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&filter && candidate->isVisible()
                && candidate->findChild<QCalendarWidget*>("access-control-"+calendarControl)) { assert(!popup); popup=candidate; }
        if (!popup) return;
        choose.stop(); popupGuard=popup; deadline.start(3000);
        auto* calendar=popup->findChild<QCalendarWidget*>("access-control-"+calendarControl); assert(calendar);
        closedCalendar=makePreviewFormObject(*popup)->member("Controls",{linaccess::VbaValue::string(calendarControl.toStdString())}).object;
        assert(closedCalendar->member("ControlType",{}).asNumber()==119);
        // SaveAsText exports this original control disabled. Prepare an enabled
        // UI context explicitly; do not discard its imported NotDefault state.
        assert(!closedCalendar->member("Enabled",{}).truthy() && !calendar->isEnabled());
        closedCalendar->assign("Enabled",linaccess::VbaValue::logical(true));
        assert(calendar->isEnabled());
        const auto module=linaccess::decodeVbaSource(calendarCode);
        const auto procedures=linaccess::analyzeVba(module.toStdString()).procedures; const auto lines=module.split('\n');
        QString next,previous;
        for (std::size_t i=0;i<procedures.size();++i) {
            const auto& procedure=procedures[i]; const auto end=i+1<procedures.size() ? procedures[i+1].line-1 : lines.size();
            const auto body=lines.mid(procedure.line-1,end-procedure.line+1).join('\n');
            for (const auto& control:calendarLayout.controls) if (control.name+"_Click"==QString::fromStdString(procedure.name)) {
                if (body.contains(".NextMonth",Qt::CaseInsensitive)) next=control.name;
                if (body.contains(".PreviousMonth",Qt::CaseInsensitive)) previous=control.name;
            }
        }
        assert(!next.isEmpty() && !previous.isEmpty());
        calendar->setCurrentPage(2003,12);
        auto* forward=popup->findChild<QPushButton*>("access-control-"+next); auto* back=popup->findChild<QPushButton*>("access-control-"+previous); assert(forward && back);
        forward->click(); assert(calendar->yearShown()==2004 && calendar->monthShown()==1);
        back->click(); assert(calendar->yearShown()==2003 && calendar->monthShown()==12);
        calendar->setSelectedDate(wanted); calendar->showSelectedDate(); QApplication::processEvents();
        auto* table=calendar->findChild<QTableView*>(); assert(table && table->currentIndex().isValid());
        const auto cell=table->visualRect(table->currentIndex()); assert(!cell.isEmpty()); visited=true;
        QTest::mouseClick(table->viewport(),Qt::LeftButton,Qt::NoModifier,cell.center());
        const auto clicked=popup->eventResult("control:"+calendarControl+":click");
        assert(clicked && clicked->status==linaccess::VbaExecutionStatus::Completed);
    });
    choose.start(1);
    QTest::mouseDClick(input,Qt::LeftButton);
    assert(QTest::qWaitFor([&] { return visited && popupGuard.isNull(); },4000));
    choose.stop(); guard.stop(); deadline.stop();
    const auto debugEvent=filter.eventResult("control:"+target+":dblclick");
    if (!visited || unexpectedDialogs || !popupGuard.isNull())
        std::cerr<<"Calendar path: visited="<<visited<<", dialogs="<<unexpectedDialogs<<", destroyed="<<popupGuard.isNull()
            <<", event="<<static_cast<bool>(debugEvent)<<", status="<<(debugEvent ? static_cast<int>(debugEvent->status) : -1)
            <<", error="<<(debugEvent ? debugEvent->errorNumber : 0)<<", line="<<(debugEvent ? debugEvent->line : 0)<<"\n";
    if (debugEvent && debugEvent->status!=linaccess::VbaExecutionStatus::Completed) {
        const auto& detail=debugEvent->explanation;
        std::cerr<<"Calendar error categories: identifier="<<(detail.find("Variable nicht deklariert")!=std::string::npos)
            <<", numeric="<<(detail.find("Zahl erwartet")!=std::string::npos)<<", syntax="<<(detail.find("Ausdruck")!=std::string::npos)
            <<", object="<<(detail.find("Object")!=std::string::npos || detail.find("Objekt")!=std::string::npos)
            <<", scope="<<(detail.find("scope")!=std::string::npos)<<"\n";
    }
    assert(visited && !timedOut && unexpectedDialogs==0 && popupGuard.isNull());
    const auto event=filter.eventResult("control:"+target+":dblclick");
    assert(event && event->status==linaccess::VbaExecutionStatus::Completed);
    const auto actual=makePreviewFormObject(filter)->item(target.toStdString());
    assert(actual.kind==linaccess::VbaValue::Kind::Text && actual.asText()==(wanted.toString(Qt::ISODate)+" 00:00:00").toStdString());
    // Repeat against the other original event binding: Screen and the shared
    // global target must follow the newly focused control, not the first one.
    auto* other=filter.findChild<QLineEdit*>("access-control-"+targets.back()); assert(other);
    QApplication::setActiveWindow(&filter); other->setFocus(); QApplication::processEvents();
    assert(QApplication::focusWidget()==other);
    visited=false; timedOut=false; guard.start(10); choose.start(1);
    QTest::mouseDClick(other,Qt::LeftButton);
    assert(QTest::qWaitFor([&] { return visited && popupGuard.isNull(); },4000));
    choose.stop(); guard.stop(); deadline.stop();
    assert(visited && !timedOut && unexpectedDialogs==0 && popupGuard.isNull());
    const auto secondEvent=filter.eventResult("control:"+targets.back()+":dblclick");
    assert(secondEvent && secondEvent->status==linaccess::VbaExecutionStatus::Completed);
    assert(makePreviewFormObject(filter)->item(targets.back().toStdString()).asText()==actual.asText());
    bool invalidated=false;
    try { (void)closedCalendar->member("Value",{}); } catch(const linaccess::VbaRuntimeError& error) { invalidated=error.number==2450; }
    assert(invalidated);
    const auto bracket=[](QString name) { return "["+name.replace(']',"]]")+"]"; };
    const auto query="SELECT Count(*) FROM "+bracket(dateTable)+" WHERE "+bracket(dateColumn)+"=Forms!"+bracket(filterName)+"!"+bracket(target)+" AND "+bracket(dateColumn)+"=Forms!"+bracket(filterName)+"!"+bracket(targets.back());
    const auto result=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(query),{});
    assert(result.rows.size()==1 && result.rows.front().front().toInt()==expectedCount);
    assert(filter.close());
}

void testCalendarValues() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","calendar-values"); db.setDatabaseName(":memory:"); assert(db.open());
    const QByteArray source=R"FORM(Begin Form
Begin Section
Begin CustomControl
Name ="calendar"
Class ="MSCAL.Calendar.7"
End
Begin CustomControl
Name ="unsupported"
Class ="Unknown.Control.1"
End
End
End
)FORM";
    AccessFormPreview form("calendar values",source,{},db);
    auto object=makePreviewFormObject(form)->member("Controls",{linaccess::VbaValue::string("CALENDAR")}).object;
    assert(object->member("Name",{}).asText()=="calendar");
    assert(form.findChild<QCalendarWidget*>("access-control-calendar") && !form.findChild<QCalendarWidget*>("access-control-unsupported"));
    object->assign("Value",linaccess::VbaValue::string("2003-02-03"));
    assert(object->member("Value",{}).asText()=="2003-02-03 00:00:00");
    for (const auto& [value,error]:std::vector<std::pair<linaccess::VbaValue,int>>{
        {linaccess::VbaValue{},94},{linaccess::VbaValue::numeric(1),13},
        {linaccess::VbaValue::string("2003-02-30"),13},{linaccess::VbaValue::string("2003-02-03 12:30:00"),3251}}) {
        try { object->assign("Value",value); assert(false); } catch(const linaccess::VbaRuntimeError& failure) { assert(failure.number==error); }
        assert(object->member("Value",{}).asText()=="2003-02-03 00:00:00");
    }
    QSqlQuery query(db); assert(query.exec("CREATE TABLE stamps(id INTEGER PRIMARY KEY,stamp DATETIME)"));
    assert(query.exec("INSERT INTO stamps VALUES(1,'2003-01-01 12:30:00')"));
    auto bound=source; bound.replace("Begin Form\n","Begin Form\nRecordSource =\"stamps\"\n"); bound.replace("Name =\"calendar\"","Name =\"calendar\"\nControlSource =\"stamp\"");
    AccessFormPreview boundForm("bound calendar",bound,{},db);
    auto boundObject=makePreviewFormObject(boundForm)->member("Controls",{linaccess::VbaValue::string("calendar")}).object;
    assert(!boundForm.findChild<QCalendarWidget*>("access-control-calendar")->isEnabled());
    try { (void)boundObject->member("Value",{}); assert(false); } catch(const linaccess::VbaRuntimeError& error) { assert(error.number==3251); }
    assert(query.exec("SELECT stamp FROM stamps") && query.next() && query.value(0).toString()=="2003-01-01 12:30:00"); query.finish();
    assert(boundForm.close() && form.close());
}
