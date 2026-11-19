#include "native_project_gui_extensions.hpp"
#include "native_macro_storage.hpp"
#include "native_access_cursor.hpp"
#include <QStatusBar>
#include <QProgressBar>
#include "native_standard_module_host.hpp"
#include <QTextBrowser>
#include "native_schema_editor.hpp"
#include "native_macro.hpp"
#include "native_module_definition.hpp"
#include "vba_module_context.hpp"
#include "native_stored_form.hpp"
#include "native_stored_report.hpp"
#include <QDateTime>
#include <QLocale>
#include <QScopeGuard>
#include <bit>
#include <limits>
#include "native_relation_designer.hpp"
#include "native_datasheet_page_store.hpp"
#include <QInputDialog>
#include "mainwindow.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_view.hpp"
#include "native_query_view.hpp"
#include "access_query.hpp"
#include "native_datasheet_print.hpp"
#include "access_table_import.hpp"
#include "access_autonumber.hpp"
#include "dao_database.hpp"
#include "native_relation_designer.hpp"
#include "access_catalog.hpp"
#include <QProcess>
#include <QPrintDialog>
#include <QPrintPreviewWidget>
#include <QPrinter>
#include <QRegularExpression>
#include <QAction>
#include <QApplication>
#include <QCryptographicHash>
#include <QClipboard>
#include <QCheckBox>
#include <QFontDialog>
#include <QFontInfo>
#include <QFontDatabase>
#include <QColorDialog>
#include <QStyleOptionButton>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QDoubleSpinBox>
#include <QFileDialog>
#include <QFile>
#include <QFileInfo>
#include <QLineEdit>
#include <QListWidget>
#include <QSpinBox>
#include <QSettings>
#include "native_classic_style.hpp"
#include <QToolBar>
#include <QJsonArray>
#include <QJsonDocument>
#include <QLabel>
#include <QPlainTextEdit>
#include <QTableView>
#include <QTableWidget>
#include <QTabWidget>
#include <QTabBar>
#include <QContextMenuEvent>
#include <QInputDialog>
#include <QMessageBox>
#include <QMouseEvent>
#include <QPointer>
#include <QMenu>
#include <QPushButton>
#include <QSqlQuery>
#include <QSqlDriver>
#include <QLibrary>
#include <QSplitter>
#include <QStandardPaths>
#include <QTemporaryDir>
#include <QTest>
#include <QPrintPreviewWidget>
#include <QTimer>
#include <QDateTime>
#include <QTreeWidget>
#include <QTreeWidgetItemIterator>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <cstdlib>
#include <iostream>

struct sqlite3;
struct sqlite3_context;
struct sqlite3_value;

namespace {
void disposeStoredFormEvents() {QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);QApplication::processEvents();}
void testNativeBoundPrefixes(QSqlDatabase db) {
    QSqlQuery query(db);
    assert(query.exec(R"SQL(CREATE TABLE prefix_fields(
        "Customer Name" TEXT, "customer-name" TEXT, "1st" TEXT, "_1st" TEXT,
        "BTNneu" TEXT, "btnSpeichern" TEXT, "btnAktion" TEXT
    ))SQL"));
    const auto rejected=[&](const QJsonObject& definition) {
        try { (void)linaccess::nativeStoredForm("AmbiguousPrefixes",definition,db); }
        catch(const linaccess::VbaRuntimeError& error) {
            assert(error.number==5);
            return;
        }
        assert(false && "Expected native stored form to reject an event prefix collision");
    };
    rejected({{"table","prefix_fields"},{"fields",QJsonArray{"Customer Name","customer-name"}}});
    rejected({{"table","prefix_fields"},{"fields",QJsonArray{"1st","_1st"}},{"controls",QJsonArray{}}});
    for(const auto* name:{"BTNneu","btnSpeichern","btnAktion"})
        rejected({{"table","prefix_fields"},{"fields",QJsonArray{QString::fromLatin1(name)}}});
    rejected({{"table","prefix_fields"},{"fields",QJsonArray{"Customer Name"}},
              {"controls",QJsonArray{QJsonObject{{"name","customer_name"},{"type","TextBox"}}}}});

    const QJsonObject definition{{"table","prefix_fields"},{"fields",QJsonArray{"Customer Name","1st"}},
                                 {"vba","Private Sub Customer_Name_Change()\nEnd Sub\n"}};
    const auto native=linaccess::nativeStoredForm("DistinctPrefixes",definition,db);
    const auto parsed=linaccess::parseAccessFormLayout(QString::fromUtf8(native.layout));
    int boundControls=0;
    for(const auto& control:parsed.controls) {
        if(control.type!="TextBox") continue;
        ++boundControls;
        assert(control.name==control.controlSource);
        if(control.name=="Customer Name") {
            assert(control.eventProcPrefix=="Customer_Name");
            assert(control.onChange=="[Event Procedure]");
        } else {
            assert(control.name=="1st");
            assert(control.eventProcPrefix=="_1st");
        }
    }
    assert(boundControls==2);
    assert(query.exec("DROP TABLE prefix_fields"));
}

void exerciseTextBoxMouseMove(QSqlDatabase db) {
    using linaccess::AccessDocumentKind;
    using linaccess::VbaExecutionStatus;
    const QString code=R"VBA(Private Sub first_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If VarType(Button)<>2 Or VarType(Shift)<>2 Or VarType(X)<>4 Or VarType(Y)<>4 Then Err.Raise 13
If X=75 Then Err.Raise 5
If X=90 Then
    DoCmd.Close acForm,Me.Name
Else
    Me.Tag=CStr(Button) & ":" & CStr(Shift) & ":" & CStr(X) & ":" & CStr(Y)
End If
End Sub
Private Function RememberMove()
Me.Tag="expression"
End Function
Private Sub Extra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Tag="extra"
End Sub
Private Sub Other_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Tag="unexpected"
End Sub
)VBA";
    const auto stored=linaccess::nativeStoredForm("Mouse Move Form",{
        {"table","native_designed_records"},
        {"fields",QJsonArray{"first","second value"}},
        {"field_events",QJsonObject{{"second value",QJsonObject{{"onmousemove","=RememberMove()"}}}}},
        {"controls",QJsonArray{
            QJsonObject{{"type","TextBox"},{"name","Extra"}},
            QJsonObject{{"type","TextBox"},{"name","Silent"}},
            QJsonObject{{"type","CommandButton"},{"name","Other"}}
        }},
        {"vba",code}
    },db);
    const auto imported=linaccess::parseAccessFormLayout(QString::fromUtf8(stored.layout));
    QMap<QString,QString> bindings;
    for(const auto& control:imported.controls) bindings.insert(control.name,control.onMouseMove);
    assert(bindings.value("first")=="[Event Procedure]" && bindings.value("second value")=="=RememberMove()");
    assert(bindings.value("Extra")=="[Event Procedure]" && bindings.value("Silent").isEmpty() && bindings.value("Other").isEmpty());

    QPointer<AccessFormPreview> form=AccessFormPreview::showDocument(AccessDocumentKind::Form,stored.name,stored.layout,stored.code,db);
    assert(form);
    auto* first=form->findChild<QLineEdit*>("access-control-first");
    auto* second=form->findChild<QLineEdit*>("access-control-second value");
    auto* extra=form->findChild<QLineEdit*>("access-control-Extra");
    auto* silent=form->findChild<QLineEdit*>("access-control-Silent");
    auto* other=form->findChild<QPushButton*>("access-control-Other");
    assert(first && first->hasMouseTracking() && second && second->hasMouseTracking());
    assert(extra && extra->hasMouseTracking() && silent && other);
    const auto move=[](QWidget* widget,QPointF position,Qt::MouseButtons buttons=Qt::NoButton,
        Qt::KeyboardModifiers modifiers=Qt::NoModifier) {
        QMouseEvent event(QEvent::MouseMove,position,QPointF(widget->mapToGlobal(position.toPoint())),Qt::NoButton,buttons,modifiers);
        QApplication::sendEvent(widget,&event);
    };
    const auto root=makePreviewFormObject(*form);
    const auto tag=[&] { return root->member("Tag",{}).asText(); };
    const QString firstEvent="control:first:MouseMove";
    assert(!form->eventResult(firstEvent));
    move(first,{3.5,4});
    assert(tag()=="0:0:52.5:60" && form->eventResult(firstEvent)->status==VbaExecutionStatus::Completed);
    move(first,{4,3},Qt::LeftButton|Qt::RightButton|Qt::MiddleButton,Qt::ShiftModifier|Qt::ControlModifier|Qt::AltModifier);
    assert(tag()=="7:7:60:45" && !form->hasUnsavedChanges() && first->text()=="before");
    move(silent,{3,4},Qt::LeftButton);
    move(other,{3,4},Qt::LeftButton);
    assert(tag()=="7:7:60:45" && !form->eventResult("control:Silent:MouseMove") && !form->eventResult("control:Other:MouseMove"));
    move(first,{4,3},Qt::ExtraButton1);
    assert(form->eventResult(firstEvent)->status==VbaExecutionStatus::Unsupported && form->eventResult(firstEvent)->errorNumber==3251);
    move(first,{5,3});
    assert(form->eventResult(firstEvent)->status==VbaExecutionStatus::Error && form->eventResult(firstEvent)->errorNumber==5);
    move(second,{3,4});
    assert(tag()=="expression" && form->eventResult("control:second value:MouseMove")->status==VbaExecutionStatus::Completed);
    move(extra,{3,4});
    assert(tag()=="extra" && !form->hasUnsavedChanges());
    move(first,{6,3});
    assert(form && !form->isVisible() && form->eventResult(firstEvent)->status==VbaExecutionStatus::Completed);
    assert(!AccessFormPreview::projectObjectLoaded(db,AccessDocumentKind::Form,stored.name));
    disposeStoredFormEvents();
    assert(!form);
}
}
void exerciseStoredForm(MainWindow& window,QSqlDatabase db) {
    testNativeBoundPrefixes(db);
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE native_designed_records(id INTEGER PRIMARY KEY,first TEXT,\"second value\" TEXT)"));
    assert(sql.exec("INSERT INTO native_designed_records VALUES(1,'before','initial')"));
    sql.finish();
    exerciseTextBoxMouseMove(db);
    const QString module=R"VBA(Private state As Long
Private mouseTrace As String
Private heldOwner As Access.Form
Private heldControl As Control
Function TypedSelf() As Form
Set TypedSelf=Me
End Function
Sub ClearForm(ByRef value As Access.Form)
Set value=Nothing
End Sub
Sub MarkControl(ByVal value As Access.Control)
value.Tag="typed-control"
End Sub
Function IsFirstControl(ByVal unused As Long, ByVal value As Control) As Boolean
IsFirstControl=value.Name="first"
End Function
Function Len(ByVal value As Control) As Long
Len=VBA.Len(value.Name)
End Function
Function EmptyControl(Optional ByVal value As Control=Nothing) As Boolean
EmptyControl=value Is Nothing
End Function
Sub EditTypedField()
Dim target As TextBox
Set target=Me!first
target="typed edit"
End Sub
Private Sub Form_Load()
state=1
Dim owner As Form, field As Access.TextBox, wrong As Object, slotError As Long, argumentError As Long
Set owner=TypedSelf()
Set heldOwner=owner
Set field=Me!first
Set heldControl=field
MarkControl Me!first
If Len(Me!first)<>5 Then Err.Raise 5
If Not IsFirstControl(value:=Me!first,unused:=1) Then Err.Raise 5
ClearForm owner
If Not owner Is Nothing Then Err.Raise 5
If Not EmptyControl() Then Err.Raise 5
Set wrong=CurrentDb
On Error Resume Next
Set heldOwner=wrong
slotError=Err.Number
Err.Clear
MarkControl wrong
argumentError=Err.Number
On Error GoTo 0
If slotError<>13 Or argumentError<>13 Then Err.Raise 5
If Not heldOwner Is Me Then Err.Raise 5
If field.Name<>"first" Then Err.Raise 5
If Not TypeOf Me Is Access.Form Then Err.Raise 5
If Not TypeOf Me!first Is TextBox Then Err.Raise 5
If Not TypeOf Me!btnSpeichern Is CommandButton Then Err.Raise 5
If Not TypeOf Me.Controls Is Access.Controls Then Err.Raise 5
If TypeOf Me Is Report Then Err.Raise 5
End Sub
Private Sub first_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button=1 Then mouseTrace=CStr(Button) & ":" & CStr(Shift) & ":" & CStr(X) & ":" & CStr(Y)
End Sub
Private Sub first_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button=1 Then mouseTrace=mouseTrace & "|up"
End Sub
Private Sub first_Click()
Me.Tag=mouseTrace & "|click"
Me!first.SelStart=0
Me!first.SelLength=0
End Sub
Private Sub first_DblClick(Cancel As Integer)
Me.Tag="text-double"
Cancel=True
End Sub
Private Sub btnSpeichern_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Tag="save-down"
End Sub
Private Sub btnSpeichern_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.Tag<>"save-down" Then Err.Raise 5
Me!first=Me!first & "|up"
End Sub
Private Sub btnAktion_Click()
state=state+1
Me![second value]=CStr(state)
DoCmd.RunCommand acCmdSaveRecord
End Sub
Sub OpenOwnedChild()
DoCmd.OpenForm "Owned Child"
DoCmd.Close acForm,"Owned Native Form"
End Sub
)VBA";
    auto* toolbar=window.findChild<QToolBar*>(); assert(toolbar);
    QAction* create=nullptr; for(auto* action:toolbar->actions()) if(action->text()=="+ Formular") create=action;
    assert(create);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-form-designer");
        dialog->findChild<QLineEdit*>("native-form-name")->setText("Owned Native Form");
        dialog->findChild<QLineEdit*>("native-form-caption")->setText("Owned form caption");
        dialog->findChild<QComboBox*>("native-form-source")->setCurrentText("native_designed_records");
        auto* fields=dialog->findChild<QListWidget*>("native-form-fields"); assert(fields && fields->count()==2);
        auto* moved=fields->takeItem(1); assert(moved->data(Qt::UserRole).toString()=="second value");
        moved->setText("Second \"caption\""); fields->insertItem(0,moved);
        fields->setCurrentItem(moved);
        dialog->findChild<QSpinBox*>("native-form-field-left")->setValue(2400);
        dialog->findChild<QSpinBox*>("native-form-field-width")->setValue(2400);
        auto* code=dialog->findChild<QPlainTextEdit*>("native-form-vba"); code->setPlainText(module);
        QApplication::setActiveWindow(dialog); code->setFocus(); QApplication::processEvents();
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(!dialog->isWindowModified());
        assert(dialog->isVisible() && QApplication::activeModalWidget()==dialog);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='form' AND name='Owned Native Form'") && sql.next());
        assert(QJsonDocument::fromJson(sql.value(0).toByteArray()).object().value("vba").toString()==module); sql.finish();
        auto* undo=dialog->findChild<QPushButton*>("native-designer-undo");auto* redo=dialog->findChild<QPushButton*>("native-designer-redo");assert(undo && redo);
        auto* caption=dialog->findChild<QLineEdit*>("native-form-caption");auto* source=dialog->findChild<QComboBox*>("native-form-source");
        caption->setText("Unsaved designer caption");source->setCurrentText("SELECT id FROM native_designed_records");
        fields=dialog->findChild<QListWidget*>("native-form-fields");assert(fields->count()==1);fields->setCurrentRow(0);
        fields->item(0)->setData(Qt::UserRole+2,QJsonObject{{"OnClick","[Event Procedure]"}});
        dialog->findChild<QSpinBox*>("native-form-field-left")->setValue(2700);code->setPlainText(module+"\n' unsaved authoring change\n");
        assert(undo->isEnabled() && dialog->isWindowModified());undo->click();
        assert(!dialog->isWindowModified() && caption->text()=="Owned form caption" && source->currentText()=="native_designed_records");
        assert(fields->count()==2 && fields->item(0)->data(Qt::UserRole).toString()=="second value" && fields->item(0)->text()=="Second \"caption\"");
        assert(fields->item(0)->data(Qt::UserRole+2).toJsonObject().isEmpty());
        assert(dialog->findChild<QSpinBox*>("native-form-field-left")->value()==2400 && code->toPlainText()==module && redo->isEnabled());
        redo->click();assert(dialog->isWindowModified() && caption->text()=="Unsaved designer caption" && fields->count()==1);
        assert(fields->item(0)->data(Qt::UserRole+2).toJsonObject().value("OnClick").toString()=="[Event Procedure]");
        assert(dialog->findChild<QSpinBox*>("native-form-field-left")->value()==2700 && code->toPlainText()!=module);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='form' AND name='Owned Native Form'") && sql.next());
        assert(QJsonDocument::fromJson(sql.value(0).toByteArray()).object().value("caption").toString()=="Owned form caption");sql.finish();
        undo->click();assert(!dialog->isWindowModified() && code->toPlainText()==module);
        dialog->findChild<QPushButton*>("native-form-save")->click();
    });
    create->trigger(); QApplication::processEvents();
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='form' AND name='Owned Native Form'") && sql.next());
    const auto saved=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(saved.value("fields").toArray().first().toString()=="second value" && saved.value("vba").toString()==module);
    assert(saved.value("field_layout").toObject().value("second value").toObject().value("width").toInt()==2400);
    auto child=saved; child.insert("vba",""); child.insert("caption","Owned child caption");
    sql.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('form','Owned Child',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(child).toJson(QJsonDocument::Compact))); assert(sql.exec()); sql.finish();
    auto* explorer=window.findChild<QTreeWidget*>(); assert(explorer);
    for(auto* action:toolbar->actions()) if(action->text()=="Aktualisieren") action->trigger();
    QTreeWidgetItem* selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="form" && (*item)->text(0)=="Owned Native Form") selected=*item;
    assert(selected); explorer->setCurrentItem(selected);
    const auto open=[&] { assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0))); QApplication::processEvents(); };
    const auto locate=[&](const QString& name) {
        QPointer<AccessFormPreview> result;
        for(const auto& document:AccessFormPreview::projectDocuments(db)) if(document && document->documentName()==name) result=document;
        return result;
    };
    assert(sql.exec("CREATE TEMP TABLE _la_objects(kind TEXT,name TEXT,definition TEXT)"));
    assert(sql.exec("INSERT INTO temp._la_objects VALUES('form','Owned Native Form','{}')"));
    open(); auto form=locate("Owned Native Form"); assert(form && form->hasDataSource() && form->windowTitle().contains("Owned form caption"));
    assert(sql.exec("DROP TABLE temp._la_objects"));
    auto* first=form->findChild<QLineEdit*>("access-control-first"); auto* second=form->findChild<QLineEdit*>("access-control-second value");
    assert(first && second && second->y()<first->y() && second->x()>first->x() && second->width()<first->width() && first->text()=="before");
    bool caption=false; for(auto* label:form->findChildren<QLabel*>()) if(label->text()=="Second \"caption\"") caption=true;
    assert(caption && AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Native Form"));
    auto formObject=makePreviewFormObject(*form);
    const auto loadResult=form->eventResult("OnLoad");
    if(loadResult && loadResult->status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Typed Access load error="<<loadResult->errorNumber<<" line="<<loadResult->line<<" explanation="<<loadResult->explanation<<"\n";
    assert(loadResult && loadResult->status==linaccess::VbaExecutionStatus::Completed);
    assert(first->property("la_access_tag").toString()=="typed-control");
    std::vector<linaccess::VbaValue> typedArguments;
    assert(form->executeProcedure("EditTypedField",typedArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text()=="typed edit" && form->hasUnsavedChanges());
    assert(sql.exec("SELECT first FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="before"); sql.finish();
    (void)formObject->member("Undo",{}); assert(first->text()=="before" && !form->hasUnsavedChanges());
    QApplication::setActiveWindow(form); first->setFocus();
    QTest::mouseClick(first,Qt::LeftButton,Qt::ShiftModifier|Qt::ControlModifier,QPoint(3,4));
    assert(formObject->member("Tag",{}).asText()=="1:3:45:60|up|click" && !formObject->member("Dirty",{}).boolean && first->cursorPosition()==0);
    formObject->assign("Tag",linaccess::VbaValue::string("right-click"));
    QTest::mouseClick(first,Qt::RightButton);
    assert(formObject->member("Tag",{}).asText()=="right-click");
    QTest::mouseDClick(first,Qt::LeftButton);
    assert(formObject->member("Tag",{}).asText()=="text-double");
    first->setFocus(); first->selectAll(); QTest::keyClicks(first,"native save");
    auto* save=form->findChild<QPushButton*>("access-control-btnSpeichern"); auto* action=form->findChild<QPushButton*>("access-control-btnAktion"); assert(save && action);
    QTest::mouseClick(save,Qt::LeftButton); QApplication::processEvents();
    assert(sql.exec("SELECT first FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="native save|up"); sql.finish();
    QTest::mouseClick(action,Qt::LeftButton); QTest::mouseClick(action,Qt::LeftButton); QApplication::processEvents();
    assert(sql.exec("SELECT \"second value\" FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="3"); sql.finish();
    std::vector<linaccess::VbaValue> arguments;
    assert(form->executeProcedure("OpenOwnedChild",arguments).status==linaccess::VbaExecutionStatus::Completed); disposeStoredFormEvents();
    auto ownedChild=locate("Owned Child"); assert(ownedChild && ownedChild->isVisible() && !AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Native Form"));
    assert(ownedChild->close()); disposeStoredFormEvents();
    {
        const QByteArray layout="Begin Form\nBegin Section\nBegin Subform\nName =\"child\"\nSourceObject =\"Form.Owned Child\"\nWidth =6000\nHeight =3000\nEnd\nEnd\nEnd\n";
        AccessFormPreview parent("Owned Parent",layout,{},db,&window);
        AccessFormPreview* embedded=nullptr;
        for(auto* dialog:parent.findChildren<QDialog*>()) if(auto* document=dynamic_cast<AccessFormPreview*>(dialog)) embedded=document;
        assert(embedded && embedded->documentName()=="Owned Child" && !embedded->isWindow() && embedded->hasDataSource());
        auto childObject=makePreviewFormObject(*embedded);
        childObject->assign("first",linaccess::VbaValue::string("embedded save"));
        assert(childObject->member("Dirty",{}).boolean);
        childObject->assign("Dirty",linaccess::VbaValue::logical(false));
        assert(sql.exec("SELECT first FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="embedded save"); sql.finish();
        assert(AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Child"));
        assert(parent.close() && !AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Child"));
    }
    open(); form=locate("Owned Native Form"); assert(form);
    auto* edit=window.findChild<QAction*>("native-form-edit"); assert(edit);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-form-designer");
        assert(!AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Native Form"));
        const auto* fields=dialog->findChild<QListWidget*>("native-form-fields");
        assert(fields->item(0)->data(Qt::UserRole).toString()=="second value" && fields->item(0)->text()=="Second \"caption\"");
        dialog->findChild<QListWidget*>("native-form-fields")->setCurrentRow(0);
        assert(dialog->findChild<QSpinBox*>("native-form-field-left")->value()==2400);
        assert(dialog->findChild<QSpinBox*>("native-form-field-width")->value()==2400);
        assert(dialog->findChild<QPlainTextEdit*>("native-form-vba")->toPlainText()==module);
        dialog->reject();
    });
    edit->trigger(); disposeStoredFormEvents();
    const QString totalsSql="SELECT COUNT(*) AS total FROM [Owned Form ZBase]";
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('query','Owned Form ZBase',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql","SELECT id FROM native_designed_records"}}).toJson(QJsonDocument::Compact)));
    assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('query','Owned Form Totals',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql",totalsSql}}).toJson(QJsonDocument::Compact)));
    assert(sql.exec()); sql.finish();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->findChild<QLineEdit*>("native-form-name")->setText("Owned Query Form");
        auto* source=dialog->findChild<QComboBox*>("native-form-source"); assert(source->isEditable());
        source->setCurrentText("SELECT 42 AS probe");
        auto* fields=dialog->findChild<QListWidget*>("native-form-fields");
        assert(fields->count()==1 && fields->item(0)->data(Qt::UserRole).toString()=="probe");
        source->setCurrentText("Owned Form Totals");
        assert(fields->count()==1 && fields->item(0)->data(Qt::UserRole).toString()=="total");
        dialog->findChild<QPushButton*>("native-form-save")->click();
    });
    create->trigger(); QApplication::processEvents();
    selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="form" && (*item)->text(0)=="Owned Query Form") selected=*item;
    assert(selected); explorer->setCurrentItem(selected); open();
    auto queryForm=locate("Owned Query Form"); assert(queryForm && queryForm->hasDataSource());
    auto* total=queryForm->findChild<QLineEdit*>("access-control-total"); assert(total && total->isReadOnly());
    assert(sql.exec(totalsSql) && sql.next()); const int beforeCount=sql.value(0).toInt(); sql.finish();
    assert(total->text().toInt()==beforeCount);
    assert(!makePreviewFormObject(*queryForm)->member("AllowAdditions",{}).boolean);
    assert(!queryForm->findChild<QPushButton*>("access-control-btnNeu"));
    assert(!queryForm->findChild<QPushButton*>("access-control-btnSpeichern"));
    assert(sql.exec("INSERT INTO native_designed_records(first) VALUES('query change')")); sql.finish();
    (void)makePreviewFormObject(*queryForm)->member("Requery",{});
    assert(total->text().toInt()==beforeCount+1);
    assert(makePreviewApplicationObject(*queryForm)->member("Eval",{linaccess::VbaValue::string(R"VBA(DLookup("id","Owned Form ZBase","id=1"))VBA")}).asNumber()==1);
    sql.prepare("UPDATE main._la_objects SET definition=? WHERE kind='query' AND name='Owned Form ZBase'");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql","SELECT id FROM native_designed_records WHERE id<0"}}).toJson(QJsonDocument::Compact)));
    assert(sql.exec()); sql.finish();
    (void)makePreviewFormObject(*queryForm)->member("Requery",{});
    assert(total->text()=="0");
    assert(queryForm->close()); disposeStoredFormEvents();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        assert(dialog->findChild<QComboBox*>("native-form-source")->currentText()=="Owned Form Totals");
        assert(dialog->findChild<QListWidget*>("native-form-fields")->count()==1);
        dialog->reject();
    });
    explorer->setCurrentItem(selected); edit->trigger(); disposeStoredFormEvents();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-query-designer");
        dialog->findChild<QLineEdit*>("native-query-name")->setText("Shortcut Saved Query");
        auto* code=dialog->findChild<QPlainTextEdit*>("native-query-sql"); code->setPlainText("SELECT 42 AS value");
        assert(dialog->isWindowModified());
        QApplication::setActiveWindow(dialog); code->setFocus(); QApplication::processEvents();
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(!dialog->isWindowModified());
        assert(dialog->isVisible() && QApplication::activeModalWidget()==dialog);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='query' AND name='Shortcut Saved Query'") && sql.next());
        auto definition=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
        assert(!dialog->isWindowModified() && !dialog->property("la_designer_new").toBool());
        assert(definition.value("sql").toString()=="SELECT 42 AS value"); definition.insert("archive","preserve");
        sql.prepare("UPDATE main._la_objects SET definition=? WHERE kind='query' AND name='Shortcut Saved Query'");
        sql.addBindValue(QString::fromUtf8(QJsonDocument(definition).toJson(QJsonDocument::Compact))); assert(sql.exec()); sql.finish();
        code->setPlainText("SELECT 43 AS value");
        assert(dialog->isWindowModified());
        assert(dialog->findChild<QLabel*>("native-designer-save-status")->text()=="Nicht gespeicherte Änderungen");
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(!dialog->isWindowModified());
        assert(dialog->isVisible() && QApplication::activeModalWidget()==dialog);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='query' AND name='Shortcut Saved Query'") && sql.next());
        definition=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
        assert(definition.value("sql").toString()=="SELECT 43 AS value" && definition.value("archive").toString()=="preserve");
        dialog->findChild<QPushButton*>("native-query-run")->click();
        auto* results=dialog->findChild<QTableView*>(); assert(results && results->model()->rowCount()==1 && results->model()->index(0,0).data().toInt()==43);
        dialog->reject();
    });
    QAction* createQuery=nullptr;
    for(auto* candidate:toolbar->actions()) if(candidate->text()=="+ Abfrage") createQuery=candidate;
    assert(createQuery); createQuery->trigger(); disposeStoredFormEvents();
    assert(sql.exec("DELETE FROM main._la_objects WHERE kind='query' AND name='Shortcut Saved Query'")); sql.finish();

    assert(sql.exec("CREATE TABLE free_control_effects(value INTEGER)"));
    const QString freeCode=R"VBA(Sub Form_Load()
Me.Choice.Requery
End Sub
Sub ApplyValue_Click()
CurrentDb.Execute "INSERT INTO free_control_effects(value) VALUES (" & CLng(Me.EntryValue) & ")"
Me.ResultLabel.Caption=CStr(Me.EntryValue)
Me.Listing.Requery
End Sub
)VBA";
    QTimer::singleShot(0,&window,[&] {
        auto* designer=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(designer && designer->objectName()=="native-form-designer");
        designer->findChild<QLineEdit*>("native-form-name")->setText("Free Controls Form");
        designer->findChild<QLineEdit*>("native-form-caption")->setText("Freies Formular");
        assert(designer->findChild<QComboBox*>("native-form-source")->currentText().isEmpty());
        assert(designer->findChild<QListWidget*>("native-form-fields")->count()==0);
        designer->findChild<QPlainTextEdit*>("native-form-vba")->setPlainText(freeCode);
        QTimer::singleShot(0,&window,[&] {
            auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-form-controls-editor");
            auto* grid=editor->findChild<QTableWidget*>("native-form-controls-grid");auto* add=editor->findChild<QPushButton*>("native-form-controls-add");assert(grid && add);
            const QStringList types{"TextBox","TextBox","Label","CommandButton","CheckBox","ComboBox","ListBox","TextBox","Label"};
            const QStringList names{"EntryValue","ComputedValue","ResultLabel","ApplyValue","EnabledFlag","Choice","Listing","MaskedValue","HiddenLabel"};
            for(int row=0;row<names.size();++row) {add->click();qobject_cast<QComboBox*>(grid->cellWidget(row,0))->setCurrentText(types[row]);grid->item(row,1)->setText(names[row]);}
            grid->item(1,3)->setText("=[EntryValue]*2");grid->item(2,2)->setText("Bereit");grid->item(3,2)->setText("Übernehmen");grid->item(4,2)->setText("Aktiv");
            qobject_cast<QComboBox*>(grid->cellWidget(5,8))->setCurrentText("Value List");grid->item(5,9)->setText("alpha;1;beta;2");
            grid->item(5,11)->setText("2");grid->item(5,12)->setText("2");grid->item(5,13)->setText("1800;0");
            grid->item(6,9)->setText("SELECT value,value*2 AS doubled FROM free_control_effects");grid->item(6,7)->setText("700");
            grid->item(6,11)->setText("2");grid->item(6,12)->setText("2");grid->item(6,13)->setText("0;3cm");
            grid->item(0,14)->setText("3");grid->item(1,15)->setText("0000");grid->item(4,14)->setText("False");
            grid->item(7,14)->setText("\"mask-value\"");grid->item(7,16)->setText("Password");
            grid->item(7,5)->setText("3600");grid->item(8,2)->setText("Hidden metadata label");
            const auto stateRow=[&](int row) {grid->setCurrentCell(row,1);};
            stateRow(0);editor->findChild<QSpinBox*>("native-form-control-tab-index")->setValue(10);
            stateRow(1);editor->findChild<QCheckBox*>("native-form-control-tab-stop")->setChecked(false);
            for(const auto pair:{std::pair{5,11},std::pair{6,12},std::pair{3,13},std::pair{4,14}}) {stateRow(pair.first);editor->findChild<QSpinBox*>("native-form-control-tab-index")->setValue(pair.second);}
            stateRow(7);editor->findChild<QCheckBox*>("native-form-control-locked")->setChecked(true);editor->findChild<QCheckBox*>("native-form-control-tab-stop")->setChecked(false);
            stateRow(8);editor->findChild<QCheckBox*>("native-form-control-visible")->setChecked(false);editor->findChild<QCheckBox*>("native-form-control-enabled")->setChecked(false);
            assert(!editor->findChild<QCheckBox*>("native-form-control-tab-stop")->isEnabled());
            stateRow(0);grid->item(0,7)->setText("420");
            editor->findChild<QLineEdit*>("native-form-control-tag")->setText("entry metadata");editor->findChild<QComboBox*>("native-form-control-alignment")->setCurrentIndex(2);
            QTimer::singleShot(0,&window,[&] {auto* picker=qobject_cast<QFontDialog*>(QApplication::activeModalWidget());assert(picker);QFont font;bool found=false;for(const auto& family:QFontDatabase::families()) {for(const auto& style:QFontDatabase::styles(family)) if(QFontDatabase::bold(family,style) && QFontDatabase::italic(family,style)) {font=QFontDatabase::font(family,style,12);found=true;break;}if(found) break;}assert(found);font.setUnderline(true);picker->setCurrentFont(font);picker->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();});
            QTest::mouseClick(editor->findChild<QPushButton*>("native-form-control-font"),Qt::LeftButton);assert(grid->currentRow()==0);
            for(const auto& choice:QList<QPair<QString,QColor>>{{"native-form-control-fore-color",QColor(12,34,56)},{"native-form-control-back-color",QColor(230,220,210)}}) {
                const auto wanted=choice.second;QTimer::singleShot(0,&window,[wanted] {auto* picker=qobject_cast<QColorDialog*>(QApplication::activeModalWidget());assert(picker);picker->setCurrentColor(wanted);picker->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();});
                QTest::mouseClick(editor->findChild<QPushButton*>(choice.first),Qt::LeftButton);
                assert(grid->currentRow()==0);
            }
            stateRow(2);editor->findChild<QCheckBox*>("native-form-control-transparent")->setChecked(true);
            editor->findChild<QComboBox*>("native-form-control-alignment")->setCurrentIndex(3);
            stateRow(8);const QString longCaption(33000,'x');grid->item(8,2)->setText(longCaption);grid->setCurrentCell(8,2);grid->scrollToItem(grid->item(8,2));QApplication::setActiveWindow(editor);grid->setFocus();grid->editItem(grid->item(8,2));QApplication::processEvents();
            auto* propertyEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget());
            if(!propertyEditor || !grid->isAncestorOf(propertyEditor)) {auto* focus=QApplication::focusWidget();std::cerr<<"Property editor focus diagnostic: "<<(focus ? focus->metaObject()->className() : "none")<<" row="<<grid->currentRow()<<" col="<<grid->currentColumn()<<" editors="<<grid->findChildren<QLineEdit*>().size()<<'\n';}
            assert(propertyEditor && grid->isAncestorOf(propertyEditor));
            assert(propertyEditor->maxLength()>=33001 && propertyEditor->text()==longCaption);propertyEditor->setCursorPosition(propertyEditor->text().size());QTest::keyClicks(propertyEditor,"z");
            editor->findChild<QPushButton*>("native-form-control-font")->setFocus();QApplication::processEvents();assert(grid->item(8,2)->text()==longCaption+"z");
            grid->item(1,1)->setText("EntryValue");editor->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();
            assert(editor->isVisible() && editor->findChild<QLabel*>("native-form-controls-status")->text().contains("(5)"));
            grid->item(1,1)->setText("ComputedValue");editor->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();
        });
        designer->findChild<QPushButton*>("native-form-controls")->click();
        auto* undo=designer->findChild<QPushButton*>("native-designer-undo");auto* redo=designer->findChild<QPushButton*>("native-designer-redo");assert(undo->isEnabled());undo->click();redo->click();
        designer->findChild<QPushButton*>("native-form-save")->click();
    });
    create->trigger();disposeStoredFormEvents();
    assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='form' AND name='Free Controls Form'") && sql.next());
    const auto freeDefinition=QJsonDocument::fromJson(sql.value(0).toByteArray()).object();sql.finish();
    assert(freeDefinition.value("table").toString().isEmpty() && freeDefinition.value("fields").toArray().isEmpty() && freeDefinition.value("controls").toArray().size()==9);
    selected=nullptr;for(QTreeWidgetItemIterator item(explorer);*item;++item) if((*item)->data(0,Qt::UserRole).toString()=="form" && (*item)->text(0)=="Free Controls Form") selected=*item;
    assert(selected);explorer->setCurrentItem(selected);open();auto freeForm=locate("Free Controls Form");assert(freeForm && !freeForm->hasDataSource());
    auto freeObject=makePreviewFormObject(*freeForm);
    auto* entry=freeForm->findChild<QLineEdit*>("access-control-EntryValue");auto* computed=freeForm->findChild<QLineEdit*>("access-control-ComputedValue");
    auto* button=freeForm->findChild<QPushButton*>("access-control-ApplyValue");auto* check=freeForm->findChild<QCheckBox*>("access-control-EnabledFlag");auto* choice=freeForm->findChild<QComboBox*>("access-control-Choice");
    assert(entry && computed && button && check && choice && !entry->isReadOnly() && computed->isReadOnly() && choice->count()==2);
    if(entry->font().pointSize()!=12 || entry->font().weight()!=QFont::Bold || !entry->font().italic() || !entry->font().underline()) {
        const auto metadata=freeDefinition.value("controls").toArray().first().toObject().value("font").toObject();
        std::cerr<<"Authored font diagnostic: widget="<<entry->font().pointSize()<<','<<entry->font().weight()<<','<<entry->font().italic()<<','<<entry->font().underline()<<" stored="<<metadata.value("size").toInt()<<','<<metadata.value("weight").toInt()<<','<<metadata.value("italic").toBool()<<','<<metadata.value("underline").toBool()<<'\n';
    }
    assert(entry->alignment()&Qt::AlignHCenter);
    assert(freeObject->member("Controls",{linaccess::VbaValue::string("EntryValue")}).object->member("Tag",{}).asText()=="entry metadata");
    assert(entry->font().pointSize()==12 && entry->font().weight()==QFont::Bold && entry->font().italic() && entry->font().underline());
    assert(QFontInfo(entry->font()).weight()>=QFont::Bold && QFontInfo(entry->font()).italic());
    if(entry->palette().color(QPalette::Active,QPalette::Text)!=QColor(12,34,56) || entry->palette().color(QPalette::Active,QPalette::Base)!=QColor(230,220,210)) {
        const auto metadata=freeDefinition.value("controls").toArray().first().toObject();
        std::cerr<<"Authored colors diagnostic: text="<<entry->palette().color(QPalette::Active,QPalette::Text).name().toStdString()<<" base="<<entry->palette().color(QPalette::Active,QPalette::Base).name().toStdString()<<" stored="<<metadata.value("fore_color").toInt(-1)<<','<<metadata.value("back_color").toInt(-1)<<'\n';
    }
    assert(entry->palette().color(QPalette::Active,QPalette::Text)==QColor(12,34,56) && entry->palette().color(QPalette::Active,QPalette::Base)==QColor(230,220,210));
    auto* resultLabel=freeForm->findChild<QLabel*>("access-control-ResultLabel");assert(resultLabel && !resultLabel->autoFillBackground() && resultLabel->palette().color(QPalette::Active,QPalette::Window).alpha()==0);
    assert(resultLabel->alignment()&Qt::AlignRight);
    auto* masked=freeForm->findChild<QLineEdit*>("access-control-MaskedValue");assert(masked && masked->isReadOnly() && masked->echoMode()==QLineEdit::Password && masked->text()=="mask-value" && masked->displayText()!=masked->text());
    auto* hiddenLabel=freeForm->findChild<QLabel*>("access-control-HiddenLabel");assert(hiddenLabel && hiddenLabel->isHidden() && !hiddenLabel->isEnabled());
    assert(hiddenLabel->text().size()==33001 && hiddenLabel->text().endsWith('z'));
    assert(!(computed->focusPolicy()&Qt::TabFocus) && !(masked->focusPolicy()&Qt::TabFocus));
    assert(freeObject->item("MaskedValue").asText()=="mask-value" && entry->text()=="3" && computed->text()=="0006");
    assert(!check->isChecked() && freeObject->item("EnabledFlag").kind==linaccess::VbaValue::Kind::Boolean);
    assert(sql.exec("SELECT COUNT(*) FROM free_control_effects") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    QApplication::setActiveWindow(freeForm);entry->setFocus();entry->selectAll();QTest::keyClicks(entry,"17");entry->clearFocus();QApplication::processEvents();
    assert(computed->text()=="0034" && freeObject->item("ComputedValue").asNumber()==34);
    entry->setFocus();QTest::keyClick(entry,Qt::Key_Tab);auto* focused=QApplication::focusWidget();assert(focused==choice || (focused && choice->isAncestorOf(focused)));QStyleOptionButton checkboxStyle;checkboxStyle.initFrom(check);const auto indicator=check->style()->subElementRect(QStyle::SE_CheckBoxIndicator,&checkboxStyle,check);QTest::mouseClick(check,Qt::LeftButton,Qt::NoModifier,indicator.center());assert(check->isChecked() && freeObject->item("EnabledFlag").truthy());choice->setCurrentIndex(1);assert(choice->currentText()=="beta" && freeObject->item("Choice").asText()=="2");
    auto* choiceView=qobject_cast<QTreeView*>(choice->view());assert(choiceView && choiceView->isColumnHidden(1));
    QTest::mouseClick(button,Qt::LeftButton);QApplication::processEvents();
    const auto clicked=freeForm->eventResult("control:ApplyValue:Click");assert(clicked && clicked->status==linaccess::VbaExecutionStatus::Completed);
    assert(freeForm->findChild<QLabel*>("access-control-ResultLabel")->text()=="17");
    assert(freeObject->member("Controls",{linaccess::VbaValue::string("Listing")}).object->member("ListCount",{}).asNumber()==1);
    auto* listing=freeForm->findChild<QTreeView*>("access-control-Listing");assert(listing && listing->model()->columnCount()==2 && listing->isColumnHidden(0));
    QTest::mouseClick(listing->viewport(),Qt::LeftButton,Qt::NoModifier,listing->visualRect(listing->model()->index(0,1)).center());
    assert(freeObject->item("Listing").asNumber()==34);
    assert(sql.exec("SELECT COUNT(*),MIN(value),MAX(value) FROM free_control_effects") && sql.next() && sql.value(0).toInt()==1 && sql.value(1).toInt()==17 && sql.value(2).toInt()==17);sql.finish();
    assert(freeForm->close());disposeStoredFormEvents();
    const auto blank=linaccess::nativeStoredForm("Blank",QJsonObject{{"table",""},{"fields",QJsonArray{}}},db);
    assert(linaccess::parseAccessFormLayout(QString::fromUtf8(blank.layout)).controls.isEmpty());
}

