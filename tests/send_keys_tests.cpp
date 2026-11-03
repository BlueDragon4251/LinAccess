#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testSendKeys() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-sendkeys"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,first_value TEXT,second_value TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'before','start')"));
    const QByteArray definition=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Height =2000
Begin TextBox
Name ="first"
ControlSource ="first_value"
Top =100
End
Begin TextBox
Name ="second"
ControlSource ="second_value"
Top =600
End
Begin TextBox
Name ="keys"
Visible =0
End
Begin TextBox
Name ="unset"
Visible =0
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Immediate()
DoCmd.GoToControl "[first]"
SendKeys "hello{TAB}^aworld",True
End Sub
Sub Queued()
With CodeContextObject
.keys.Requery
DoCmd.GoToControl "first"
SendKeys .keys,False
End With
End Sub
Sub QueuedPair()
DoCmd.GoToControl "first"
SendKeys "abc",False
DoCmd.SendKeys "XYZ",False
End Sub
Sub QueuedThenYield()
DoCmd.GoToControl "first"
SendKeys "abc",False
DoCmd.SendKeys "XYZ",False
DoEvents
If Me.first.Text<>"abcXYZ" Then Err.Raise 5
Me.Tag="queue-drained"
End Sub
Sub NullDefault()
With CodeContextObject
SendKeys .unset,False
End With
End Sub
Sub ObjectNullDefault()
SendKeys Me.Controls("unset"),False
End Sub
Sub Invalid()
SendKeys "^a{UNSUPPORTED_KEY}",True
End Sub
Sub Unicode()
SendKeys "Grüße😀",True
End Sub
Sub Escaped()
SendKeys "{+}{^}{%}{(}{)}{{}{}}",True
End Sub
)VBA";
    AccessFormPreview form("keyboard input",definition,code,db); std::vector<VbaValue> none;
    assert(form.executeProcedure("Immediate",none).errorNumber==2110);
    auto* first=form.findChild<QLineEdit*>("access-control-first");
    auto* second=form.findChild<QLineEdit*>("access-control-second"); assert(first && second);
    form.show(); QApplication::setActiveWindow(&form); first->setFocus(); QApplication::processEvents();
    assert(QApplication::focusWidget()==first); first->deselect();
    assert(form.executeProcedure("Immediate",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text()=="hello" && second->text()=="world" && QApplication::focusWidget()==second);
    assert(form.close());
    assert(query.exec("SELECT first_value,second_value FROM records WHERE id=1") && query.next());
    assert(query.value(0).toString()=="hello" && query.value(1).toString()=="world"); query.finish();
    auto unbound=definition; unbound.replace("RecordSource =\"records\"",""); unbound.replace("ControlSource =\"first_value\"",""); unbound.replace("ControlSource =\"second_value\"","");
    AccessFormPreview queued("queued keys",unbound,code,db);
    auto object=makePreviewFormObject(queued);
    first=queued.findChild<QLineEdit*>("access-control-first"); second=queued.findChild<QLineEdit*>("access-control-second");
    queued.show(); QApplication::setActiveWindow(&queued); first->setFocus(); QApplication::processEvents();
    assert(queued.executeProcedure("NullDefault",none).errorNumber==94);
    assert(queued.executeProcedure("ObjectNullDefault",none).errorNumber==94);
    object->assignItem("keys",VbaValue::string("queued"));
    assert(queued.executeProcedure("Queued",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text().isEmpty() && object->item("keys").asText()=="queued"); assert(QTest::qWaitFor([&] { return first->text()=="queued"; },2000));
    assert(queued.executeProcedure("Invalid",none).errorNumber==3251);
    assert(first->text()=="queued" && first->selectedText().isEmpty());
    first->clear(); assert(queued.executeProcedure("Unicode",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text()==QString::fromUtf8("Grüße😀"));
    first->clear(); assert(queued.executeProcedure("Escaped",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text()=="+^%(){}");
    first->clear(); assert(queued.executeProcedure("QueuedPair",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(QTest::qWaitFor([&] { return first->text()=="abcXYZ"; },2000));
    first->clear(); assert(queued.executeProcedure("QueuedThenYield",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text()=="abcXYZ" && object->member("Tag",{}).asText()=="queue-drained");
    first->clear(); assert(queued.executeProcedure("Queued",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(queued.close()); QApplication::processEvents(); assert(first->text().isEmpty());

    const QByteArray aliasForm="Begin Form\nBegin Section\nBegin TextBox\nName =\"unit price\"\nEnd\nEnd\nEnd\n";
    const QByteArray aliasCode=R"VBA(Sub WriteAlias()
Me.unit_price=5
Me.unit_price.Requery
Me.unit_price.Visible=True
Me.unit_price.SetFocus
If Me.unit_price<>5 Then Err.Raise 5
If Forms("aliases").unit_price.Value<>5 Then Err.Raise 5
If Me![unit price]<>5 Then Err.Raise 5
End Sub
Sub LiteralBang()
Me!unit_price.Requery
End Sub
Sub LiteralBangValue()
Me!unit_price=7
End Sub
)VBA";
    AccessFormPreview aliases("aliases",aliasForm,aliasCode,db);
    aliases.show(); QApplication::setActiveWindow(&aliases); QApplication::processEvents();
    assert(aliases.executeProcedure("WriteAlias",none).status==linaccess::VbaExecutionStatus::Completed);
    auto aliasObject=makePreviewFormObject(aliases);
    assert(aliasObject->objectMember("unit_price",{}).object->member("Name",{}).asText()=="unit price");
    assert(aliases.executeProcedure("LiteralBang",none).errorNumber==2465);
    assert(aliases.executeProcedure("LiteralBangValue",none).errorNumber==2465);
    bool literalRejected=false;
    try { (void)aliasObject->member("Controls",{VbaValue::string("unit_price")}); }
    catch(const linaccess::VbaRuntimeError& error) { literalRejected=error.number==2465; }
    assert(literalRejected && aliasObject->item("unit price").asNumber()==5);
    assert(aliases.close());
    auto ambiguous=aliasForm;
    ambiguous.replace("End\nEnd\nEnd\n","End\nBegin TextBox\nName =\"unit-price\"\nEnd\nEnd\nEnd\n");
    AccessFormPreview collision("collision",ambiguous,"Sub ReadAlias()\nMe.unit_price=5\nEnd Sub\n",db);
    assert(collision.executeProcedure("ReadAlias",none).errorNumber==2465);
    assert(collision.close());
    ambiguous.replace("End\nEnd\nEnd\n","End\nBegin TextBox\nName =\"unit_price\"\nEnd\nEnd\nEnd\n");
    AccessFormPreview exact("exact",ambiguous,"Sub ReadAlias()\nMe.unit_price=8\nEnd Sub\n",db);
    assert(exact.executeProcedure("ReadAlias",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(makePreviewFormObject(exact)->item("unit_price").asNumber()==8);
    assert(makePreviewFormObject(exact)->item("unit price").kind==VbaValue::Kind::Null);
    assert(exact.close());
}
