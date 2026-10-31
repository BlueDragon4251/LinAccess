#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include <QLineEdit>
#include <QCheckBox>
#include <QComboBox>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testFilters() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-filter-tests"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE items(id INTEGER PRIMARY KEY,amount REAL,label TEXT,stamp TEXT)"));
    assert(query.exec("INSERT INTO items VALUES(1,1.5,'first','2003-01-01 00:00:00'),(2,3.5,'second','2004-01-01 00:00:00'),(3,NULL,'nullable','2005-01-01 00:00:00')"));
    assert(query.exec("CREATE TABLE events(name TEXT)"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray definition=R"FORM(Begin Form
RecordSource ="items"
OnCurrent ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="amount"
ControlSource ="amount"
End
Begin TextBox
Name ="total"
ControlSource ="=Sum([amount])"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Apply(criteria As String)
Me.Filter=criteria
Me.FilterOn=True
End Sub
Sub Refresh()
Me.Requery
End Sub
Sub Clear()
Me.FilterOn=False
End Sub
Sub Form_Current()
CurrentDb.Execute "INSERT INTO events VALUES('current')"
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
If Me.Tag = "stay" Then
Cancel=1
End If
End Sub
)VBA";
    AccessFormPreview form("native filter",definition,code,db);
    auto object=makePreviewFormObject(form);
    auto* selector=form.findChild<QComboBox*>("access-record-selector");
    auto* editor=form.findChild<QLineEdit*>("access-filter");
    auto* toggle=form.findChild<QCheckBox*>("access-filter-on");
    assert(selector && editor && toggle && selector->count()==4);
    std::vector<linaccess::VbaValue> args{linaccess::VbaValue::string("[amount]>2")};
    assert(form.executeProcedure("Apply",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(selector->count()==2 && object->member("FilterOn",{}).truthy());
    assert(object->member("Filter",{}).asText()=="[amount]>2");
    assert(object->item("amount").asNumber()==3.5 && object->item("total").asNumber()==3.5);
    assert(toggle->isChecked() && editor->text()=="[amount]>2");
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==2); query.finish();
    assert(query.exec("UPDATE items SET amount=3.75 WHERE id=2")); query.finish();
    std::vector<linaccess::VbaValue> refreshArgs;
    assert(form.executeProcedure("Refresh",refreshArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(object->item("amount").asNumber()==3.75 && object->item("total").asNumber()==3.75);
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==3); query.finish();

    args.front()=linaccess::VbaValue::string("[amount]>");
    const auto invalid=form.executeProcedure("Apply",args);
    assert(invalid.status==linaccess::VbaExecutionStatus::Error && invalid.errorNumber==3075);
    assert(selector->count()==2 && object->item("amount").asNumber()==3.75 && object->member("Filter",{}).asText()=="[amount]>2");
    std::vector<linaccess::VbaValue> none;
    assert(form.executeProcedure("Clear",none).status==linaccess::VbaExecutionStatus::Completed && selector->count()==4);
    assert(!object->member("FilterOn",{}).truthy());
    editor->setText("[stamp]>=#2004-01-01# AND [label] Like 's*'"); toggle->click();
    assert(toggle->isChecked() && selector->count()==2 && object->item("amount").asNumber()==3.75);
    object->assign("Tag",linaccess::VbaValue::string("stay"));
    object->member("Controls",{linaccess::VbaValue::string("amount")}).object->assign("Value",linaccess::VbaValue::numeric(4));
    const auto cancelled=form.executeProcedure("Clear",none);
    assert(cancelled.status==linaccess::VbaExecutionStatus::Error && cancelled.errorNumber==2501);
    assert(form.hasUnsavedChanges() && toggle->isChecked() && selector->count()==2);
    object->assign("Tag",linaccess::VbaValue::string(""));
    assert(form.executeProcedure("Clear",none).status==linaccess::VbaExecutionStatus::Completed && selector->count()==4);
    assert(query.exec("SELECT amount FROM items WHERE id=2") && query.next() && query.value(0).toDouble()==4); query.finish();
    assert(form.close());

    const auto saved=QByteArray("Filter =\"[id]=2\"\nFilterOnLoad =-1\n");
    auto savedDefinition=definition; savedDefinition.replace("RecordSource =\"items\"\n","RecordSource =\"items\"\n"+saved);
    AccessFormPreview savedForm("saved filter",savedDefinition,{},db);
    assert(savedForm.findChild<QComboBox*>("access-record-selector")->count()==2);
    assert(makePreviewFormObject(savedForm)->item("amount").asNumber()==4);
    assert(savedForm.close());
    auto emptySaved=savedDefinition; emptySaved.replace("[id]=2","");
    AccessFormPreview emptySavedForm("empty saved filter",emptySaved,{},db);
    assert(!makePreviewFormObject(emptySavedForm)->member("FilterOn",{}).truthy());
    assert(emptySavedForm.findChild<QComboBox*>("access-record-selector")->count()==4);
    assert(emptySavedForm.close());


    auto queryDefinition=definition; queryDefinition.replace("RecordSource =\"items\"","RecordSource =\"SELECT id,amount,label,stamp FROM items\"");
    AccessFormPreview readonly("query filter",queryDefinition,code,db);
    args.front()=linaccess::VbaValue::string("[amount] Is Null");
    assert(readonly.executeProcedure("Apply",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(readonly.findChild<QComboBox*>("access-record-selector")->count()==2);
    assert(readonly.findChild<QLineEdit*>("access-control-amount")->isReadOnly());
    assert(readonly.close());

    auto brokenDefinition=savedDefinition; brokenDefinition.replace("[id]=2","[id]>");
    AccessFormPreview broken("invalid saved filter",brokenDefinition,{},db);
    assert(!broken.hasDataSource());
    makePreviewFormObject(broken)->assign("FilterOn",linaccess::VbaValue::logical(false));
    assert(broken.hasDataSource() && broken.findChild<QComboBox*>("access-record-selector")->count()==4);
    assert(broken.close());

    const auto top=linaccess::compileAccessQuery("SELECT TOP 1 id,amount FROM items ORDER BY id DESC");
    const auto filtered=linaccess::filterAccessQueryPlan(db,top,"id<3");
    assert(linaccess::executeAccessQueryPlan(db,filtered,{}).rows.empty());
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','saved_top',?)");
    query.addBindValue(QByteArray("SELECT TOP 2 id,amount FROM items ORDER BY id DESC")); assert(query.exec()); query.finish();
    auto topDefinition=definition; topDefinition.replace("RecordSource =\"items\"","RecordSource =\"saved_top\"");
    AccessFormPreview savedTop("saved top filter",topDefinition,code,db);
    args.front()=linaccess::VbaValue::string("[id]=1");
    assert(savedTop.executeProcedure("Apply",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(savedTop.hasDataSource() && savedTop.findChild<QComboBox*>("access-record-selector")->count()==1);
    assert(savedTop.close());

    const QByteArray unboundDefinition=R"FORM(Begin Form
OnCurrent ="[Event Procedure]"
Begin Section
End
End
)FORM";
    const QByteArray unboundCode=R"VBA(Sub Form_Current()
CurrentDb.Execute "INSERT INTO events VALUES('unbound_current')"
DoCmd.Requery
DoCmd.Requery
End Sub
)VBA";
    AccessFormPreview unbound("unbound requery",unboundDefinition,unboundCode,db);
    assert(query.exec("SELECT COUNT(*) FROM events WHERE name='unbound_current'") && query.next() && query.value(0).toInt()==1); query.finish();
    assert(unbound.eventResult("OnCurrent") && unbound.eventResult("OnCurrent")->status==linaccess::VbaExecutionStatus::Completed);
    assert(unbound.close());

}
