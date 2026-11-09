#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include <QSqlQuery>
#include <QLineEdit>
#include <iostream>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
namespace {
void testNativeDomainLookup() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-domain-lookup"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE lookup_values(id INTEGER PRIMARY KEY,label TEXT,qty REAL,moment DATETIME,flag BOOLEAN,cash REAL)"));
    assert(query.exec("INSERT INTO lookup_values VALUES(1,'before',10.5,'2026-11-04 00:00:00',1,12.3456),(2,'O''Brien',3.5,NULL,0,0),(3,'',NULL,NULL,NULL,NULL),(4,NULL,NULL,NULL,NULL,NULL)"));
    assert(query.exec("CREATE TABLE _la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT)"));
    query.prepare("INSERT INTO _la_table_schema VALUES('lookup_values',?)");
    query.addBindValue(QString(R"JSON([{"name":"cash","source_type":"CURRENCY"}])JSON")); assert(query.exec()); query.finish();
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','Lookup Alias',?)");
    query.addBindValue(QByteArray("SELECT id,label AS [display name],qty,moment,flag,cash FROM lookup_values")); assert(query.exec()); query.finish();
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','Lookup Ordered',?)");
    query.addBindValue(QByteArray("SELECT id,label,qty,moment,flag,cash FROM lookup_values ORDER BY id")); assert(query.exec()); query.finish();
    const QByteArray source=R"FORM(Begin Form
RecordSource ="lookup_values"
Begin Section
Height =1500
Begin TextBox
Name ="Edited"
ControlSource ="label"
End
Begin TextBox
Name ="LookedUp"
ControlSource ="=DLookup(""label"",""lookup_values"",""id=1"")"
Top =400
End
End
End
)FORM";
    const QByteArray module=R"VBA(Sub Checks()
If DLookup("label","lookup_values","id=1")<>"before" Then Err.Raise 5
If DLookup(Expr:="[display name]",Domain:="lookup alias",Criteria:="id=2")<>"O'Brien" Then Err.Raise 5
If Application.DLookup(Domain:="lookup_values",Expr:="qty",Criteria:="id=2")<>3.5 Then Err.Raise 5
If TypeName(DLookup("moment","lookup_values","id=1"))<>"Date" Then Err.Raise 5
If DLookup("moment","lookup_values","id=1")<>DateSerial(2026,11,4) Then Err.Raise 5
If TypeName(DLookup("flag","lookup_values","id=1"))<>"Boolean" Then Err.Raise 5
If DLookup("flag","lookup_values","id=1")<>True Then Err.Raise 5
If TypeName(DLookup("cash","lookup_values","id=1"))<>"Currency" Then Err.Raise 5
If DLookup("cash","lookup_values","id=1")<>CCur("12.3456") Then Err.Raise 5
If DMin("qty","lookup_values")<>3.5 Then Err.Raise 5
If Application.DMax(Domain:="lookup_values",Expr:="qty")<>10.5 Then Err.Raise 5
If DAvg(Expr:="qty",Domain:="lookup_values")<>7 Then Err.Raise 5
If DAvg("qty","lookup_values","id=1")<>10.5 Then Err.Raise 5
If Not IsNull(DMin("qty","lookup_values","id=99")) Then Err.Raise 5
If Not IsNull(DMax("qty","lookup_values","id=3")) Then Err.Raise 5
If Not IsNull(DAvg("qty","lookup_values","id=99")) Then Err.Raise 5
If TypeName(DMin("moment","lookup_values"))<>"Date" Then Err.Raise 5
If DMax("moment","lookup_values")<>DateSerial(2026,11,4) Then Err.Raise 5
If TypeName(Application.DMax("cash","lookup_values"))<>"Currency" Then Err.Raise 5
If DMax("cash","lookup_values")<>CCur("12.3456") Then Err.Raise 5
If DMin("cash","lookup_values")<>CCur(0) Then Err.Raise 5
If TypeName(DSum("cash","lookup_values"))<>"Currency" Then Err.Raise 5
If DSum("cash","lookup_values")<>CCur("12.3456") Then Err.Raise 5
If TypeName(Application.DSum(Expr:="cash",Domain:="lookup_values"))<>"Currency" Then Err.Raise 5
If Application.DSum("CCur(qty)","lookup_values","qty Is Not Null")<>CCur(14) Then Err.Raise 5
If Application.DCount("*","lookup_values")<>4 Then Err.Raise 5
If Application.DCount(Domain:="lookup_values",Expr:="cash")<>2 Then Err.Raise 5
If Application.DCount("cash","lookup_values","id>2")<>0 Then Err.Raise 5
If Not IsNull(Application.DSum("cash","lookup_values","id>2")) Then Err.Raise 5
If DFirst("label","Lookup Ordered")<>"before" Then Err.Raise 5
If Not IsNull(Application.DLast(Domain:="Lookup Ordered",Expr:="label")) Then Err.Raise 5
If DLast("qty","Lookup Ordered","qty Is Not Null")<>3.5 Then Err.Raise 5
If DFirst(Expr:="qty",Domain:="Lookup Ordered",Criteria:="qty Is Not Null")<>10.5 Then Err.Raise 5
If Not IsNull(DFirst("id","Lookup Ordered","id=99")) Then Err.Raise 5
If Not IsNull(DLast("id","Lookup Ordered","id=99")) Then Err.Raise 5
If TypeName(DLast("cash","Lookup Ordered","id=1"))<>"Currency" Then Err.Raise 5
If DLookup("[qty]*2","lookup_values","id=1")<>21 Then Err.Raise 5
If DLookup("[qty]+[id]","lookup_values","id=2")<>5.5 Then Err.Raise 5
If DLookup("[label] & ':' & [id]","lookup_values","id=2")<>"O'Brien:2" Then Err.Raise 5
If TypeName(DLookup("CCur([qty])","lookup_values","id=1"))<>"Currency" Then Err.Raise 5
If DLookup("DateAdd('d',1,[moment])","lookup_values","id=1")<>DateSerial(2026,11,5) Then Err.Raise 5
If Not IsNull(DLookup("[qty]*2","lookup_values","id=99")) Then Err.Raise 5
If DFirst("[qty]*2","Lookup Ordered")<>21 Then Err.Raise 5
If Not IsNull(DLast("[qty]*2","Lookup Ordered")) Then Err.Raise 5
If DMin("[qty]*2","lookup_values")<>7 Then Err.Raise 5
If DMax("[qty]*2","lookup_values")<>21 Then Err.Raise 5
If DAvg("[qty]*2","lookup_values")<>14 Then Err.Raise 5
If DVar("[qty]*2","lookup_values")<>98 Then Err.Raise 5
If Abs(DStDev("[qty]*2","lookup_values")-Sqr(98))>0.000000001 Then Err.Raise 5
If DLookup("IIf([id]=1,'FROM SELECT','other')","lookup_values","id=1")<>"FROM SELECT" Then Err.Raise 5
If DVar("qty","lookup_values")<>24.5 Then Err.Raise 5
If Abs(Application.DStDev(Domain:="lookup_values",Expr:="qty")-Sqr(24.5))>0.000000001 Then Err.Raise 5
If Not IsNull(DVar("qty","lookup_values","id=1")) Then Err.Raise 5
If Not IsNull(DStDev("qty","lookup_values","id=99")) Then Err.Raise 5
If Not IsNull(DLookup("label","lookup_values","id=99")) Then Err.Raise 5
If Not IsNull(DLookup("label","lookup_values","id=4")) Then Err.Raise 5
If IsNull(DLookup("label","lookup_values","id=3")) Then Err.Raise 5
If Len(DLookup("label","lookup_values","id=3"))<>0 Then Err.Raise 5
Dim arbitrary As Long
arbitrary=DLookup("id","lookup_values")
If arbitrary<1 Or arbitrary>4 Then Err.Raise 5
Me!Edited="draft"
If DLookup("label","lookup_values","id=1")<>"before" Then Err.Raise 5
Me.Dirty=False
If DLookup("label","lookup_values","id=1")<>"draft" Then Err.Raise 5
Me.Tag="lookups checked"
End Sub
Sub MissingDomain()
Dim value As Variant
value=DLookup("label","missing_lookup_domain")
End Sub
Sub UnsupportedExpression()
Dim value As Variant
value=DLookup("Sum(qty)","lookup_values")
End Sub
)VBA";
    AccessFormPreview form("lookup test",source,module,db);
    assert(form.findChild<QLineEdit*>("access-control-LookedUp")->text()=="before");
    std::vector<linaccess::VbaValue> args;
    const auto result=form.executeProcedure("Checks",args);
    if(result.status!=linaccess::VbaExecutionStatus::Completed)
        std::cerr<<"Synthetic DLookup: error="<<result.errorNumber<<", line="<<result.line<<", "<<result.explanation<<'\n';
    assert(result.status==linaccess::VbaExecutionStatus::Completed);
    assert(makePreviewFormObject(form)->member("Tag",{}).asText()=="lookups checked");
    assert(query.exec("SELECT label FROM lookup_values WHERE id=1") && query.next() && query.value(0).toString()=="draft"); query.finish();
    assert(form.executeProcedure("MissingDomain",args).errorNumber==3078);
    assert(form.executeProcedure("UnsupportedExpression",args).errorNumber==3251);
    for(const auto& invalid:QStringList{"qty FROM lookup_values","qty;DELETE FROM lookup_values","(SELECT qty FROM lookup_values)","qty, id","qty AS x","qty) --","unknown_function(qty)","Var(qty)","qty OVER()"}) {
        int refused=0;try{(void)linaccess::accessScalarDomainExpression(invalid);}catch(const linaccess::VbaRuntimeError& failure){refused=failure.number;}assert(refused==3251);
    }
    assert(query.exec("SELECT COUNT(*) FROM lookup_values") && query.next() && query.value(0).toInt()==4);query.finish();
    assert(form.close());
}
}
void testSubforms() {
    testNativeDomainLookup();
    auto db=QSqlDatabase::addDatabase("QSQLITE","embedded-form-test"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE items(id INTEGER PRIMARY KEY,name TEXT,amount REAL)"));
    assert(query.exec("INSERT INTO items VALUES(1,'before',10.5),(2,'second',3.5),(3,'nullable',NULL)"));
    assert(query.exec("CREATE VIEW amount_domain AS SELECT id,amount FROM items WHERE id<3"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray childSource=R"FORM(Begin Form
RecordSource ="items"
OnLoad ="[Event Procedure]"
OnUnload ="[Event Procedure]"
OnClose ="[Event Procedure]"
Begin Section
Height =1500
Begin TextBox
Name ="child_name"
Tag ="field metadata"
ControlSource ="name"
End
Begin TextBox
Name ="allow_close"
Top =400
End
Begin TextBox
Name ="total"
ControlSource ="=Sum([amount])"
Top =800
End
Begin TextBox
Name ="nonnull"
ControlSource ="=Count([amount])"
End
Begin TextBox
Name ="rowcount"
ControlSource ="=Count(*)"
End
Begin TextBox
Name ="localized"
ControlSource ="=Wenn(IstNull([total]);0;[total])"
End
Begin TextBox
Name ="domain_sum"
ControlSource ="=DSum(""amount"",""items"",""id<3"")"
End
Begin TextBox
Name ="domain_none"
ControlSource ="=DSum(""amount"",""items"",""id<0"")"
End
Begin TextBox
Name ="domain_count"
ControlSource ="=DCount(""amount"",""items"")"
End
Begin TextBox
Name ="domain_view"
ControlSource ="=DSum(""amount*2"",""amount_domain"")"
End
Begin TextBox
Name ="eager_error"
ControlSource ="=IIf(True,1,1/0)"
End
End
End
)FORM";
    const QByteArray childCode=R"VBA(
Sub CurrencyCheck()
Me!allow_close = CCur("922337203685477.5807")
End Sub
Sub DomainCheck()
Me!allow_close = DSum("amount", "items", "id=" & 2)
End Sub
Sub DomainFailure()
On Error GoTo Failed
Me!allow_close = DSum("amount", "missing_domain")
Exit Sub
Failed:
Me!allow_close = Err.Number
End Sub
Sub Form_Load()
Me!allow_close = 1
End Sub
Sub Form_Unload(Cancel As Integer)
If Me!allow_close = 0 Then
Cancel = 1
End If
End Sub
Sub Form_Close()
Me!allow_close = 2
End Sub
)VBA";
    const auto artifact=[&](const char* kind,const char* name,const QByteArray& bytes) {
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES(?,?,?)");
        query.addBindValue(kind); query.addBindValue(name); query.addBindValue(bytes); assert(query.exec()); query.finish();
    };
    artifact("form","child",childSource); artifact("vba","Form_child",childCode);
    const QByteArray parentSource=R"FORM(Begin Form
Tag ="parent metadata"
Begin Section
Height =4000
Begin Subform
Name ="embedded records"
SourceObject ="Form.child"
Width =7000
Height =3500
End
Begin TextBox
Name ="offset"
Top =3500
End
Begin TextBox
Name ="combined"
ControlSource ="=[embedded records].[Form]![total]+[offset]"
Top =3800
End
Begin ComboBox
Name ="lookup"
RowSource ="SELECT id,name,amount FROM items ORDER BY id"
RowSourceType ="Table/Query"
BoundColumn =2
ColumnCount =3
End
Begin TextBox
Name ="lookup_id"
ControlSource ="=[lookup].Column(0)"
End
Begin TextBox
Name ="lookup_first"
ControlSource ="=[lookup].Column(1;0)"
End
Begin TextBox
Name ="global_offset"
ControlSource ="=[Formulare]![parent]![offset]+1"
End
Begin ComboBox
Name ="ordinal"
RowSource ="SELECT id,name,amount FROM items ORDER BY id"
BoundColumn =0
ColumnCount =3
End
Begin TextBox
Name ="manual_error"
End
Begin TextBox
Name ="cycle_one"
ControlSource ="=[cycle_two]"
End
Begin TextBox
Name ="cycle_two"
ControlSource ="=[cycle_one]"
End
End
End
)FORM";
    const QByteArray parentCode=R"VBA(
Sub Refresh()
Me![embedded records].Requery
End Sub
Sub Change()
Dim frm As Form
Set frm = Me![embedded records].Form
frm!child_name = "edited"
Me("embedded " & "records").Tag = "child metadata"
Me.Tag = Me.Tag & ":changed"
Me!offset = 2
End Sub
Sub CheckErrors()
Me!offset = CVErr(2015)
Me!manual_error = CStr(IsError(Me!combined))
End Sub
)VBA";
    const QByteArray defaultCode=R"VBA(Sub WriteDefault()
Dim child As Object
Set child=Forms("default parent")("embedded records").Form
child("child_name").Value="default call"
child.Dirty=False
End Sub
Function ReadDefault() As String
ReadDefault=Forms("default parent")("embedded records").Form("child_name")
End Function
Sub RestoreDefault()
Forms("default parent")("embedded records").Form("child_name").Value="before"
Forms("default parent")("embedded records").Form.Dirty=False
End Sub
)VBA";
    {
        AccessFormPreview defaultParent("default parent",parentSource,defaultCode,db);
        std::vector<linaccess::VbaValue> args;
        const auto read=defaultParent.executeProcedure("ReadDefault",args);
        assert(read.status==linaccess::VbaExecutionStatus::Completed && read.typedReturnValue->asText()=="before");
        assert(defaultParent.executeProcedure("WriteDefault",args).status==linaccess::VbaExecutionStatus::Completed);
        assert(query.exec("SELECT name FROM items WHERE id=1") && query.next() && query.value(0).toString()=="default call"); query.finish();
        assert(defaultParent.executeProcedure("RestoreDefault",args).status==linaccess::VbaExecutionStatus::Completed);
        assert(query.exec("SELECT name FROM items WHERE id=1") && query.next() && query.value(0).toString()=="before"); query.finish();
        assert(defaultParent.close());
    }
    AccessFormPreview parent("parent",parentSource,parentCode,db);
    auto* child=dynamic_cast<AccessFormPreview*>(parent.findChild<QWidget*>("access-control-embedded records"));
    assert(child && !child->isWindow() && child->hasDataSource());
    assert(child->eventResult("OnLoad") && child->eventResult("OnLoad")->status==linaccess::VbaExecutionStatus::Completed);
    auto form=makePreviewFormObject(parent)->member("Controls",{linaccess::VbaValue::string("embedded records")}).object->member("Form",{}).object;
    auto controls=form->member("Controls",{}).object;
    auto name=controls->member("Item",{linaccess::VbaValue::string("child_name")}).object;
    assert(name->member("Tag",{}).asText()=="field metadata");
    bool nullTag=false;
    try { name->assign("Tag",{}); } catch(const linaccess::VbaRuntimeError& error) { nullTag=error.number==94; }
    assert(nullTag && name->member("Tag",{}).asText()=="field metadata");
    assert(makePreviewFormObject(parent)->member("Tag",{}).asText()=="parent metadata");
    auto allow=controls->member("Item",{linaccess::VbaValue::string("allow_close")}).object;
    assert(form->item("total").asNumber()==14);
    assert(form->item("domain_sum").asNumber()==14);
    assert(form->item("domain_none").kind==linaccess::VbaValue::Kind::Null);
    assert(form->item("domain_count").asNumber()==2);
    assert(form->item("domain_view").asNumber()==28);
    std::vector<linaccess::VbaValue> domainArgs;
    assert(child->executeProcedure("CurrencyCheck",domainArgs).status==linaccess::VbaExecutionStatus::Completed);
    const auto preservedCurrency=allow->member("Value",{});
    assert(preservedCurrency.kind==linaccess::VbaValue::Kind::Currency && preservedCurrency.asText()=="922337203685477.5807");

    assert(child->executeProcedure("DomainCheck",domainArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(allow->member("Value",{}).asNumber()==3.5);
    assert(child->executeProcedure("DomainFailure",domainArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(allow->member("Value",{}).asNumber()==3078);
    allow->assign("Value",linaccess::VbaValue::numeric(1));

    assert(form->item("nonnull").asNumber()==2);
    assert(form->item("rowcount").asNumber()==3);
    assert(form->item("localized").asNumber()==14);
    const auto eager=form->item("eager_error");
    assert(eager.kind==linaccess::VbaValue::Kind::Error && eager.errorNumber==11);
    const auto parentControls=makePreviewFormObject(parent)->member("Controls",{}).object;
    const auto lookup=parentControls->member("Item",{linaccess::VbaValue::string("lookup")}).object;
    assert(lookup->member("Column",{linaccess::VbaValue::numeric(0)}).kind==linaccess::VbaValue::Kind::Null);
    lookup->assign("Value",linaccess::VbaValue::string("second"));
    assert(lookup->member("Value",{}).asText()=="second");
    assert(lookup->member("Column",{linaccess::VbaValue::numeric(0)}).asNumber()==2);
    assert(lookup->member("Column",{linaccess::VbaValue::numeric(2)}).asNumber()==3.5);
    assert(lookup->member("Column",{linaccess::VbaValue::numeric(1),linaccess::VbaValue::numeric(0)}).asText()=="before");
    assert(makePreviewFormObject(parent)->item("lookup_id").asNumber()==2);
    lookup->assign("Value",linaccess::VbaValue::string("nullable"));
    assert(lookup->member("Column",{linaccess::VbaValue::numeric(2)}).kind==linaccess::VbaValue::Kind::Null);
    const auto ordinal=parentControls->member("Item",{linaccess::VbaValue::string("ordinal")}).object;
    ordinal->assign("Value",linaccess::VbaValue::numeric(1));
    assert(ordinal->member("Column",{linaccess::VbaValue::numeric(0)}).asNumber()==2);
    assert(ordinal->member("Value",{}).asNumber()==1);
    assert(name->member("Value",{}).asText()=="before");
    assert(makePreviewFormsObject(parent)->member("Count",{}).asNumber()==1);
    assert(query.exec("UPDATE items SET name='external' WHERE id=1")); query.finish();
    std::vector<linaccess::VbaValue> arguments;
    assert(parent.executeProcedure("Refresh",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(name->member("Value",{}).asText()=="external");
    assert(parent.executeProcedure("Change",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(makePreviewFormObject(parent)->member("Tag",{}).asText()=="parent metadata:changed");
    assert(makePreviewFormObject(parent)->item("combined").asNumber()==16);
    assert(makePreviewFormObject(parent)->item("global_offset").asNumber()==3);
    assert(makePreviewFormObject(parent)->item("lookup_first").asText()=="before");
    assert(parent.findChild<QLineEdit*>("access-control-combined")->text()=="16");
    assert(parent.executeProcedure("CheckErrors",arguments).status==linaccess::VbaExecutionStatus::Completed);
    const auto propagated=makePreviewFormObject(parent)->item("combined");
    assert(propagated.kind==linaccess::VbaValue::Kind::Error && propagated.errorNumber==2015);
    assert(makePreviewFormObject(parent)->item("manual_error").asText()=="True");
    const auto cyclic=makePreviewFormObject(parent)->item("cycle_one");
    assert(cyclic.kind==linaccess::VbaValue::Kind::Error && cyclic.errorNumber==2455);
    linaccess::AccessQueryModel sql;
    bool rejectedError=false;
    try { sql.execute(db,"SELECT Forms!parent!combined"); }
    catch(const linaccess::VbaRuntimeError& error) { rejectedError=error.number==13; }
    assert(rejectedError);
    parentControls->member("Item",{linaccess::VbaValue::string("offset")}).object->assign("Value",linaccess::VbaValue::numeric(2));
    assert(makePreviewFormObject(parent)->member("Controls",{linaccess::VbaValue::string("embedded records")}).object->member("Tag",{}).asText()=="child metadata");
    allow->assign("Value",linaccess::VbaValue::numeric(0));
    assert(!parent.close());
    assert(form->member("Name",{}).asText()=="child");
    assert(!child->eventResult("OnClose"));
    assert(query.exec("SELECT name FROM items WHERE id=1") && query.next() && query.value(0).toString()=="edited"); query.finish();
    allow->assign("Value",linaccess::VbaValue::numeric(1));
    assert(parent.close());
    assert(child->eventResult("OnClose") && child->eventResult("OnClose")->status==linaccess::VbaExecutionStatus::Completed);
    bool closed=false;
    try { (void)form->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { closed=error.number==2450; }
    assert(closed);

    auto linked=parentSource;
    linked.replace("SourceObject =\"Form.child\"","SourceObject =\"Form.child\"\nLinkMasterFields =\"id\"\nLinkChildFields =\"id\"");
    AccessFormPreview unsupported("linked",linked,parentCode,db);
    const auto rejected=unsupported.executeProcedure("Refresh",arguments);
    assert(rejected.status==linaccess::VbaExecutionStatus::Error && rejected.errorNumber==3251);
    assert(unsupported.close());

    auto cycle=parentSource; cycle.replace("Form.child","Form.cycle"); artifact("form","cycle",cycle);
    AccessFormPreview recursive("cycle",cycle,parentCode,db);
    const auto recursion=recursive.executeProcedure("Refresh",arguments);
    assert(recursion.status==linaccess::VbaExecutionStatus::Error && recursion.errorNumber==2455);
    assert(recursive.close());

    const QByteArray wrappedSource=R"FORM(Begin Form
Caption ="A ""quoted"" "
    "caption"
RecordSource ="SELECT id,na"
    "me,amount FROM items ORDER BY id"
Begin Section
Height =1500
Begin TextBox
Name ="bound name"
ControlSource ="na"
    "me"
End
Begin ComboBox
Name ="wrapped rows"
ColumnCount =2
RowSource ="SELECT id,na"
    "me FROM items ORDER BY id"
End
End
End
)FORM";
    const auto escaped=linaccess::parseAccessFormLayout(QString::fromUtf8(R"FORM(Begin Form
Caption ="A \"quoted\" "
    "caption"
Tag ="C:\data\folder"
Begin Section
Begin TextBox
Name ="escaped expression"
ControlSource ="=DSum(\"amount\",\"items\",\"id<3\")"
End
End
End
)FORM"));
    assert(escaped.caption=="A \"quoted\" caption");
    assert(escaped.tag==R"(C:\data\folder)");
    assert(escaped.controls.size()==1);
    assert(escaped.controls.front().controlSource==R"(=DSum("amount","items","id<3"))");
    const auto decoded=linaccess::parseAccessFormLayout(QString::fromUtf8(wrappedSource));
    assert(decoded.caption=="A \"quoted\" caption" && decoded.recordSource=="SELECT id,name,amount FROM items ORDER BY id");
    assert(linaccess::parseAccessFormText(QString::fromUtf8(wrappedSource)).recordSource==decoded.recordSource);
    AccessFormPreview wrapped("wrapped source",wrappedSource,{},db);
    assert(wrapped.hasDataSource());
    assert(makePreviewFormObject(wrapped)->item("bound name").asText()=="edited");
    const auto wrappedRows=makePreviewFormObject(wrapped)->member("Controls",{linaccess::VbaValue::string("wrapped rows")}).object;
    assert(wrappedRows->member("Column",{linaccess::VbaValue::numeric(1),linaccess::VbaValue::numeric(0)}).asText()=="edited");
    assert(wrapped.close());

    AccessFormPreview failedSource("failed source",parentSource,parentCode,db);
    assert(query.exec("DROP TABLE items")); query.finish();
    const auto failedRequery=failedSource.executeProcedure("Refresh",arguments);
    assert(failedRequery.status==linaccess::VbaExecutionStatus::Error && failedRequery.errorNumber==3078);
    auto* unavailable=dynamic_cast<AccessFormPreview*>(failedSource.findChild<QWidget*>("access-control-embedded records"));
    assert(unavailable && !unavailable->hasDataSource());
    assert(failedSource.close());
}
