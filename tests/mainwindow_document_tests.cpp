#include "mainwindow.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QAction>
#include <QApplication>
#include <QCryptographicHash>
#include <QDialog>
#include <QFileDialog>
#include <QFile>
#include <QFileInfo>
#include <QLineEdit>
#include <QMessageBox>
#include <QPointer>
#include <QPushButton>
#include <QSqlQuery>
#include <QSplitter>
#include <QStandardPaths>
#include <QTemporaryDir>
#include <QTest>
#include <QTimer>
#include <QTreeWidget>
#include <QTreeWidgetItemIterator>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <cstdlib>
#include <iostream>

namespace {
QAction* action(MainWindow& window,const QString& text) {
    for(auto* candidate:window.findChildren<QAction*>()) if(candidate->text()==text) return candidate;
    assert(false); return nullptr;
}
void openProject(MainWindow& window,const QString& path,QWidget* shortcutTarget=nullptr) {
    bool selected=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->selectFile(path); selected=true;
        static_cast<QDialog*>(dialog)->accept();
    });
    if(shortcutTarget) { QTest::keyPress(shortcutTarget,Qt::Key_O,Qt::ControlModifier); QTest::keyRelease(&window,Qt::Key_O,Qt::ControlModifier); }
    else action(window,"Öffnen")->trigger();
    assert(selected);
}
void cancelProjectShortcut(QWidget* target,Qt::Key key) {
    bool rejected=false;
    QTimer::singleShot(0,target,[&] { auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog); rejected=true; dialog->reject(); });
    QTest::keyClick(target,key,Qt::ControlModifier); assert(rejected);
}
int artifact(QSqlDatabase db,const QString& kind,const QString& name,const QByteArray& bytes) {
    QSqlQuery sql(db);
    sql.prepare("INSERT INTO _la_access_artifacts(source_sha256,source_filename,kind,name,status,format,detail,source_bytes) VALUES(?,'synthetic',?,?,'source_only','synthetic','',?)");
    sql.addBindValue(QString::fromLatin1(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()));
    sql.addBindValue(kind); sql.addBindValue(name); sql.addBindValue(bytes); assert(sql.exec()); return sql.lastInsertId().toInt();
}
void openArtifact(MainWindow& window,int id) {
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer"); assert(explorer);
    QTreeWidgetItem* selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="access_artifact" && (*item)->data(0,Qt::UserRole+1).toInt()==id) { selected=*item; break; }
    assert(selected); bool opened=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());
        assert(dialog && dialog->objectName()=="access-artifact-details");
        for(auto* button:dialog->findChildren<QPushButton*>())
            if(button->text().startsWith("Access-Formular öffnen") || button->text().startsWith("Access-Bericht öffnen")) { opened=true; button->click(); return; }
        assert(false);
    });
    assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    assert(opened);
}
AccessFormPreview* document(const std::string& name) {
    for(auto* widget:QApplication::topLevelWidgets()) if(auto* form=dynamic_cast<AccessFormPreview*>(widget);form && form->isVisible())
        if(makePreviewFormObject(*form)->member("Name",{}).asText()==name) return form;
    return nullptr;
}
void dispose() { QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); QApplication::processEvents(); }
}

int main(int argc,char** argv) {
    QCoreApplication::setAttribute(Qt::AA_DontUseNativeDialogs);
    QApplication app(argc,argv); app.setQuitOnLastWindowClosed(false); app.setApplicationName("LinAccessSyntheticGuiTest");
    QStandardPaths::setTestModeEnabled(true);
    QTimer watchdog; watchdog.setSingleShot(true);
    QObject::connect(&watchdog,&QTimer::timeout,[] { std::cerr<<"Synthetic main-window document test timed out\n"; std::abort(); }); watchdog.start(20000);
    QTemporaryDir temporary; assert(temporary.isValid());
    const auto first=temporary.filePath("first.linaccess"),second=temporary.filePath("second.linaccess");
    auto seed=QSqlDatabase::addDatabase("QSQLITE","mainwindow-seed"); seed.setDatabaseName(first); assert(seed.open());
    QSqlQuery initial(seed); assert(initial.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(initial.exec("INSERT INTO records VALUES(1,'old')")); assert(initial.exec("CREATE TABLE events(id INTEGER PRIMARY KEY,kind TEXT)")); initial.finish(); seed.close();
    seed.setDatabaseName(second); assert(seed.open()); seed.close();
    MainWindow window; window.show(); QApplication::setActiveWindow(&window); QApplication::processEvents();
    openProject(window,first); auto db=QSqlDatabase::database("linaccess-main"); assert(db.isOpen() && db.databaseName()==first);
    const QByteArray form="Begin Form\nRecordSource =\"records\"\nKeyPreview =NotDefault\nOnKeyDown =\"[Event Procedure]\"\nOnActivate =\"[Event Procedure]\"\nOnDeactivate =\"[Event Procedure]\"\nOnLoad =\"[Event Procedure]\"\nOnUnload =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"editor\"\nControlSource =\"value\"\nBeforeUpdate =\"[Event Procedure]\"\nEnd\nEnd\nEnd";
    const QByteArray code="Sub Form_Load()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('load')\"\nEnd Sub\nSub Form_Unload(Cancel As Integer)\nIf Me.Tag=\"hold\" Then Cancel=True\nEnd Sub\nSub CloseSelf()\nDoCmd.Close\nMsgBox \"Synthetic retained document\"\nEnd Sub\n"
        "Sub Form_Activate()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('activate')\"\nEnd Sub\nSub Form_Deactivate()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('deactivate')\"\nEnd Sub\n"
        "Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)\nIf KeyCode=vbKeyF11 And Me.Tag=\"block-f11\" Then KeyCode=0\nIf KeyCode=vbKeyF6 And Me.Tag=\"block-cycle\" Then KeyCode=0\nIf Me.Tag=\"block-project\" And (Shift And acCtrlMask)<>0 Then\nIf KeyCode=vbKeyO Or KeyCode=vbKeyN Then KeyCode=0\nEnd If\nEnd Sub\n"
        "Sub editor_BeforeUpdate(Cancel As Integer)\nIf Me.Tag=\"cancel-focus\" Then Cancel=True\nIf Me.Tag=\"close-cycle\" Then DoCmd.Close acReport,\"report\"\nEnd Sub\n"
        "Sub BackgroundSave()\nOn Error Resume Next\nDoCmd.RunCommand 97\nMe.Tag=CStr(Err.Number)\nEnd Sub\n"
        "Sub BackgroundKeys()\nOn Error Resume Next\nSendKeys \"x\",True\nMe.Tag=CStr(Err.Number)\nEnd Sub\n"
        "Sub BackgroundClose()\nOn Error Resume Next\nDoCmd.Close\nMe.Tag=CStr(Err.Number)\nEnd Sub\n"
        "Sub ExplicitSave()\nMe.Dirty=False\nEnd Sub\n"
        "Sub SelectTable()\nDoCmd.SelectObject acTable,\"records\",True\nEnd Sub\n"
        "Sub SelectTableGroup()\nDoCmd.SelectObject acTable,,True\nEnd Sub\n"
        "Sub SelectSourceForm()\nDoCmd.SelectObject acForm,\"ENTRY\",True\nEnd Sub\n"
        "Sub SelectQuery()\nDoCmd.SelectObject acQuery,\"source_query\",True\nEnd Sub\n"
        "Sub SelectEntry()\nDoCmd.SelectObject acForm,\"entry\",False\nEnd Sub\n"
        "Sub SelectReport()\nDoCmd.SelectObject acReport,\"report\",False\nEnd Sub\n"
        "Sub ProjectKeysWhileRunning()\nSendKeys \"^o\",True\nMe.Tag=\"returned\"\nEnd Sub\n";
    const int formId=artifact(db,"form","entry",form); artifact(db,"vba","Form_entry",code);
    const int queryId=artifact(db,"query","source_query","SELECT value FROM records");
    auto report=form; report.replace("Begin Form","Begin Report"); report.replace("OnLoad =\"[Event Procedure]\"\n",""); report.replace("OnUnload =\"[Event Procedure]\"\n","");
    const int reportId=artifact(db,"report","report",report);
    auto modal=form; modal.replace("Begin Form\n","Begin Form\nModal =NotDefault\n"); const int modalId=artifact(db,"form","modal",modal); artifact(db,"vba","Form_modal",code);
    auto cancel=form; cancel.replace("OnLoad =\"[Event Procedure]\"","OnOpen =\"[Event Procedure]\"");
    const int cancelId=artifact(db,"form","cancel",cancel); artifact(db,"vba","Form_cancel","Sub Form_Open(Cancel As Integer)\nCancel=True\nEnd Sub\n");
    const int initializationId=artifact(db,"form","initial_current","Begin Form\nRecordSource =\"records\"\nOnCurrent =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"editor\"\nControlSource =\"value\"\nEnd\nEnd\nEnd");
    artifact(db,"vba","Form_initial_current","Sub Form_Current()\nMe.Tag=\"before\"\nDoCmd.RunCommand 97\nMe.Tag=\"initialized\"\nEnd Sub\n");
    action(window,"Aktualisieren")->trigger();
    window.focusDatabaseWindow();
    openArtifact(window,initializationId); QPointer<AccessFormPreview> initializing=document("initial_current"); assert(initializing);
    assert(makePreviewFormObject(*initializing)->member("Tag",{}).asText()=="initialized"); assert(initializing->close()); dispose(); assert(!initializing);
    openArtifact(window,formId); assert(QTest::qWaitFor([&] { return document("entry")!=nullptr; },3000));
    QPointer<AccessFormPreview> entry=document("entry"); assert(entry->parentWidget()==&window && !entry->isModal() && QApplication::activeModalWidget()==nullptr);
    openArtifact(window,formId); assert(document("entry")==entry);
    QSqlQuery sql(db); assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='load'") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    auto root=makePreviewFormObject(*entry); auto* editor=entry->findChild<QLineEdit*>("access-control-editor"); assert(editor);
    std::vector<linaccess::VbaValue> arguments;
    QApplication::setActiveWindow(entry); editor->setFocus(); root->assign("Tag",linaccess::VbaValue::string("block-project"));
    QTest::keyClick(editor,Qt::Key_O,Qt::ControlModifier); QTest::keyClick(editor,Qt::Key_N,Qt::ControlModifier);
    assert(editor->hasFocus() && !QApplication::activeModalWidget() && db.databaseName()==first);
    root->assign("Tag",linaccess::VbaValue::string(""));
    for(const auto key:{Qt::Key_O,Qt::Key_N}) {
        QApplication::setActiveWindow(entry); editor->setFocus(); cancelProjectShortcut(editor,key);
        assert(entry && db.databaseName()==first && db.isOpen());
    }
    QApplication::setActiveWindow(entry); editor->setFocus(); QApplication::processEvents();
    assert(entry->executeProcedure("ProjectKeysWhileRunning",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(entry && db.databaseName()==first && root->member("Tag",{}).asText()=="returned" && !QApplication::activeModalWidget());
    root->assign("Tag",linaccess::VbaValue::string(""));
    QApplication::setActiveWindow(entry); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"fresh");
    assert(entry->hasUnsavedChanges());
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer"); assert(explorer);
    for(const auto* tag:{"block-cycle","cancel-focus"}) {
        root->assign("Tag",linaccess::VbaValue::string(tag));
        for(const auto modifiers:{Qt::KeyboardModifiers(Qt::ControlModifier),Qt::ControlModifier|Qt::ShiftModifier}) {
            QTest::keyClick(editor,Qt::Key_F6,modifiers); QApplication::processEvents();
            assert(activePreviewDocument()==entry && editor->hasFocus() && !AccessFormPreview::databaseWindowActive());
        }
    }
    root->assign("Tag",linaccess::VbaValue::string("block-f11")); QTest::keyClick(editor,Qt::Key_F11); QApplication::processEvents();
    assert(editor->hasFocus() && !AccessFormPreview::databaseWindowActive());
    root->assign("Tag",linaccess::VbaValue::string("cancel-focus")); QTest::keyClick(editor,Qt::Key_F11); QApplication::processEvents();
    assert(editor->hasFocus() && !AccessFormPreview::databaseWindowActive() && root->objectMember("editor",{}).object->member("Text",{}).asText()=="fresh");
    root->assign("Tag",linaccess::VbaValue::string(""));
    auto* splitter=qobject_cast<QSplitter*>(explorer->parentWidget()); assert(splitter); splitter->setSizes({0,window.width()});
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='deactivate'") && sql.next()); const int beforeDatabase=sql.value(0).toInt(); sql.finish();
    QTest::keyClick(editor,Qt::Key_F11); assert(QTest::qWaitFor([&] { return explorer->hasFocus(); },3000));
    assert(AccessFormPreview::databaseWindowActive(&window) && QApplication::activeWindow()==&window && splitter->sizes()[0]>0);
    assert(root->objectMember("editor",{}).object->member("Value",{}).asText()=="fresh" && entry->hasUnsavedChanges());
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="old"); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='deactivate'") && sql.next() && sql.value(0).toInt()==beforeDatabase+1); sql.finish();
    QTest::keyClick(explorer,Qt::Key_F11); assert(explorer->hasFocus() && explorer->isVisible() && splitter->sizes()[0]>0);
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='deactivate'") && sql.next() && sql.value(0).toInt()==beforeDatabase+1); sql.finish();
    auto screen=makePreviewApplicationObject(*entry)->member("Screen",{}).object;
    for(const auto& property:std::vector<std::pair<const char*,int>>{{"ActiveForm",2450},{"ActiveControl",2474}}) {
        int error=0; try { screen->member(property.first,{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==property.second);
    }
    for(const auto& procedure:std::vector<std::pair<const char*,const char*>>{{"BackgroundSave","2046"},{"BackgroundKeys","3251"},{"BackgroundClose","3251"}}) {
        assert(entry->executeProcedure(procedure.first,arguments).status==linaccess::VbaExecutionStatus::Completed);
        assert(root->member("Tag",{}).asText()==procedure.second && entry->isVisible());
    }
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="old"); sql.finish();
    assert(entry->executeProcedure("ExplicitSave",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh"); sql.finish();
    QApplication::setActiveWindow(entry); editor->setFocus(); QApplication::processEvents();
    assert(!AccessFormPreview::databaseWindowActive() && screen->member("ActiveForm",{}).object->member("Name",{}).asText()=="entry");
    editor->selectAll(); QTest::keyClicks(editor,"second"); root->assign("Tag",linaccess::VbaValue::string("cancel-focus"));
    auto selection=entry->executeProcedure("SelectTable",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Error && selection.errorNumber==2501 && editor->hasFocus());
    root->assign("Tag",linaccess::VbaValue::string("")); selection=entry->executeProcedure("SelectTable",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Completed);
    assert(QTest::qWaitFor([&] { return explorer->hasFocus(); },3000));
    assert(explorer->currentItem()->data(0,Qt::UserRole+2).toString()=="table" && explorer->currentItem()->data(0,Qt::UserRole+3).toString()=="records");
    assert(entry->hasUnsavedChanges() && sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh"); sql.finish();
    assert(entry->executeProcedure("SelectTableGroup",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(explorer->currentItem()->data(0,Qt::UserRole+4).toBool() && explorer->currentItem()->data(0,Qt::UserRole+2).toString()=="table");
    assert(entry->executeProcedure("SelectSourceForm",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(explorer->currentItem()->data(0,Qt::UserRole+1).toInt()==formId && explorer->currentItem()->parent()->data(0,Qt::UserRole+2).toString()=="form");
    entry->hide(); assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(entry->isVisible() && activePreviewDocument()==entry && !AccessFormPreview::databaseWindowActive());
    assert(entry->executeProcedure("SelectQuery",arguments).status==linaccess::VbaExecutionStatus::Completed && explorer->currentItem()->data(0,Qt::UserRole+1).toInt()==queryId);
    selection=entry->executeProcedure("SelectReport",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Error && selection.errorNumber==2451 && AccessFormPreview::databaseWindowActive());
    openArtifact(window,reportId); QPointer<AccessFormPreview> selectedReport=document("report"); assert(selectedReport);
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed && activePreviewDocument()==entry);
    assert(entry->executeProcedure("SelectReport",arguments).status==linaccess::VbaExecutionStatus::Completed && activePreviewDocument()==selectedReport);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh" && entry->hasUnsavedChanges()); sql.finish();
    const auto cycle=[&](QWidget* source,bool backwards,AccessFormPreview* expected) {
        QTest::keyClick(source,Qt::Key_F6,backwards ? Qt::ControlModifier|Qt::ShiftModifier : Qt::KeyboardModifiers(Qt::ControlModifier));
        assert(QTest::qWaitFor([&] { return expected ? activePreviewDocument()==expected && QApplication::activeWindow()==expected : AccessFormPreview::databaseWindowActive(&window) && explorer->hasFocus(); },3000));
    };
    auto* reportEditor=selectedReport->findChild<QLineEdit*>("access-control-editor"); assert(reportEditor);
    QApplication::setActiveWindow(selectedReport); reportEditor->setFocus(); QApplication::processEvents();
    cycle(reportEditor,false,nullptr);
    cycle(explorer,false,entry); assert(editor->hasFocus());
    cycle(editor,false,selectedReport); assert(reportEditor->hasFocus());
    cycle(reportEditor,true,entry);
    cycle(editor,true,nullptr);
    cycle(explorer,true,selectedReport);
    {
        assert(seed.open());
        AccessFormPreview unrelated("unrelated","Begin Form\nBegin Section\nBegin TextBox\nName =\"other\"\nEnd\nEnd\nEnd",{},seed,&window);
        unrelated.show(); window.focusDatabaseWindow(); cycle(explorer,true,selectedReport);
        assert(unrelated.close()); seed.close();
    }
    selectedReport->hide(); window.focusDatabaseWindow();
    cycle(explorer,true,entry); assert(!selectedReport->isVisible());
    selectedReport->showMinimized(); window.focusDatabaseWindow();
    cycle(explorer,true,selectedReport); assert(!selectedReport->isMinimized());
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh" && entry->hasUnsavedChanges()); sql.finish();
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    QApplication::setActiveWindow(entry); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"temporary");
    root->assign("Tag",linaccess::VbaValue::string("close-cycle")); cycle(editor,false,entry);
    dispose(); assert(!selectedReport);
    root->assign("Tag",linaccess::VbaValue::string("")); editor->selectAll(); QTest::keyClicks(editor,"second");
    window.focusDatabaseWindow(); cycle(explorer,true,entry); cycle(editor,false,nullptr); cycle(explorer,false,entry);
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    auto command=makePreviewApplicationObject(*entry)->member("DoCmd",{}).object;
    using linaccess::VbaValue;
    const auto badSelection=[&](const std::vector<VbaValue>& values,int expected) {
        int error=0; try { command->member("SelectObject",values); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==expected);
    };
    badSelection({},450); badSelection({VbaValue{}},94); badSelection({VbaValue::numeric(4)},3251);
    badSelection({VbaValue::numeric(0)},450);
    badSelection({VbaValue::numeric(0),VbaValue::string("records"),VbaValue::logical(false)},3251);
    badSelection({VbaValue::numeric(2),VbaValue::string("absent")},2450);
    badSelection({VbaValue::numeric(0),VbaValue::string("absent"),VbaValue::logical(true)},7874);
    badSelection({VbaValue::numeric(0),VbaValue::string("records"),VbaValue::string("bad")},13);
    badSelection({VbaValue::numeric(0),VbaValue::string("records"),VbaValue::logical(true),VbaValue::numeric(0)},450);
    class SelectionName final:public linaccess::VbaObject {
    public:
        int reads{};
        VbaValue scalarValue() override { ++reads; return VbaValue::string("records"); }
        VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No members"); }
    };
    auto scalarName=std::make_shared<SelectionName>(); command->member("SelectObject",{VbaValue::numeric(0),VbaValue::reference(scalarName),VbaValue::logical(true)});
    assert(scalarName->reads==1 && explorer->currentItem()->data(0,Qt::UserRole+3).toString()=="records");
    assert(sql.exec("INSERT INTO _la_objects(kind,name,definition) VALUES('form','ENTRY','{}')")); action(window,"Aktualisieren")->trigger();
    badSelection({VbaValue::numeric(2),VbaValue::string("entry"),VbaValue::logical(true)},5);
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='form' AND name='ENTRY'")); action(window,"Aktualisieren")->trigger();
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(entry->executeProcedure("ExplicitSave",arguments).status==linaccess::VbaExecutionStatus::Completed);
    root->assign("Tag",linaccess::VbaValue::string("hold"));
    QApplication::setActiveWindow(entry); editor->setFocus(); openProject(window,second,editor); assert(db.databaseName()==first && db.isOpen() && entry && entry->isVisible());
    root->assign("Tag",linaccess::VbaValue::string(""));
    QApplication::setActiveWindow(entry); editor->setFocus(); QApplication::processEvents();
    bool retained=false;
    QTimer::singleShot(0,&window,[&] {
        auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box);
        dispose(); retained=entry!=nullptr; assert(retained); box->button(QMessageBox::Ok)->click();
    });
    const auto closed=entry->executeProcedure("CloseSelf",arguments); assert(closed.status==linaccess::VbaExecutionStatus::Completed && retained);
    dispose(); assert(!entry);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="second"); sql.finish();
    openArtifact(window,reportId); QPointer<AccessFormPreview> reportWindow=document("report"); assert(reportWindow && !reportWindow->isModal() && QApplication::activeModalWidget()==nullptr);
    assert(reportWindow->findChild<QLineEdit*>("access-control-editor")->isReadOnly()); assert(reportWindow->close()); dispose(); assert(!reportWindow);
    openArtifact(window,modalId); QPointer<AccessFormPreview> modalWindow=document("modal"); assert(modalWindow && modalWindow->isModal() && QApplication::activeModalWidget()==modalWindow);
    auto* modalEditor=modalWindow->findChild<QLineEdit*>("access-control-editor"); assert(modalEditor); modalEditor->setFocus(); QTest::keyClick(modalEditor,Qt::Key_F11);
    assert(QApplication::activeModalWidget()==modalWindow && !AccessFormPreview::databaseWindowActive());
    QTest::keyClick(modalEditor,Qt::Key_O,Qt::ControlModifier); QTest::keyClick(modalEditor,Qt::Key_N,Qt::ControlModifier); assert(QApplication::activeModalWidget()==modalWindow);
    QTest::keyClick(modalEditor,Qt::Key_F6,Qt::ControlModifier); QTest::keyClick(modalEditor,Qt::Key_F6,Qt::ControlModifier|Qt::ShiftModifier);
    assert(QApplication::activeModalWidget()==modalWindow && !AccessFormPreview::databaseWindowActive());
    selection=modalWindow->executeProcedure("SelectTable",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Error && selection.errorNumber==2046 && QApplication::activeModalWidget()==modalWindow);
    assert(modalWindow->close()); dispose(); assert(!modalWindow && QApplication::activeModalWidget()==nullptr);
    openArtifact(window,cancelId); assert(!document("cancel") && QApplication::activeModalWidget()==nullptr);
    openProject(window,second,explorer); assert(QSqlDatabase::database("linaccess-main").databaseName()==second);
    const auto third=temporary.filePath("third.linaccess"); bool created=false;
    QTimer::singleShot(0,&window,[&] { auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog); created=true; dialog->selectFile(third); static_cast<QDialog*>(dialog)->accept(); });
    QTest::keyClick(explorer,Qt::Key_N,Qt::ControlModifier);
    assert(created && QFileInfo::exists(third) && QSqlDatabase::database("linaccess-main").databaseName()==third);
    const auto hash=[&] { QFile input(first); assert(input.open(QIODevice::ReadOnly)); return QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256); };
    const auto beforeRefusal=hash(); bool refused=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->setOption(QFileDialog::DontConfirmOverwrite,true); dialog->selectFile(first);
        QTimer::singleShot(0,&window,[&] { auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box); refused=true; box->accept(); });
        static_cast<QDialog*>(dialog)->accept();
    });
    QTest::keyClick(explorer,Qt::Key_N,Qt::ControlModifier);
    assert(refused && hash()==beforeRefusal && QSqlDatabase::database("linaccess-main").databaseName()==third);
    assert(window.close()); dispose(); watchdog.stop();
    std::cout<<"Main window: native project/explorer opening, independent nonmodal form/report lifetime, deduplication, modal metadata, cancellation, SQL save and retained close/MsgBox passed\n";
}
