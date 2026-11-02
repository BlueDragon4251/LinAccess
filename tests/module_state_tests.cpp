#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_module_context.hpp"
#include "vba_source_codec.hpp"
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testNumericModuleStates();
void testNumericProcedureValues();
void testScalarProcedureValues();

void testModuleStates(const QDir& directory,const QJsonArray& objects) {
    QByteArray globals;
    int classModules=0,classFields=0;
    for (const auto& item:objects) {
        const auto artifact=item.toObject(); if (artifact["kind"]!="vba") continue;
        QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly));
        const auto bytes=input.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1().toLower());
        const auto source=linaccess::decodeVbaSource(bytes).toStdString();
        const auto state=linaccess::createVbaModuleState(source);
        if (artifact["sha256"]=="b58d91036b71523ee9217433bcf809539d8747b7ea7cb9bf459161e86f1c9922") globals=bytes;
        if (linaccess::isVbaClassModule(source) && !state->values.empty()) {
            ++classModules;
            for (const auto& [name,value]:state->values) { (void)name; assert(value.kind==linaccess::VbaValue::Kind::Boolean && !value.boolean); ++classFields; }
        }
    }
    assert(!globals.isEmpty() && classModules==3 && classFields==4);
    const auto declarations=linaccess::createVbaModuleState(linaccess::decodeVbaSource(globals).toStdString());
    assert(declarations->publicVariables.size()==3);
    std::string selected; int strings=0,variants=0;
    for (const auto& [name,value]:declarations->values) {
        if (value.kind==linaccess::VbaValue::Kind::Text) { ++strings; selected=name; }
        else { assert(value.kind==linaccess::VbaValue::Kind::Empty); ++variants; }
    }
    assert(strings==3 && variants==0 && !selected.empty());
    const auto open=[&](const QString& connection) {
        auto db=QSqlDatabase::addDatabase("QSQLITE",connection); db.setDatabaseName(":memory:"); assert(db.open());
        QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','original-global-declarations',?)");
        query.addBindValue(globals); assert(query.exec());
        const auto add=[&](const char* name,const QByteArray& code) {
            query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba',?,?)");
            query.addBindValue(name); query.addBindValue(code); assert(query.exec());
        };
        add("shared helper","Public Function SharedNext() As Variant\nSharedNext=999\nEnd Function\n");
        const auto helper=QByteArray("Private hits As Variant\nPublic Function SharedNext() As Variant\nSharedNext=Advance()\nEnd Function\nPrivate Function Advance() As Variant\nhits=hits+1\nAdvance=hits\nEnd Function\nPublic Function ReadStandard() As Variant\nReadStandard=")+QByteArray::fromStdString(selected)+"\nEnd Function\nPublic Function IllegalMe() As Variant\nIllegalMe=Me.Tag\nEnd Function\nPublic Function IllegalMeRequery() As Variant\nMe.exposed.Requery\nEnd Function\nPublic Function IllegalMeFocus() As Variant\nMe.exposed.SetFocus\nEnd Function\n";
        add("SHARED HELPER",helper); return db;
    };
    auto db=open("original-module-state"); auto otherDb=open("separate-module-state");
    const auto source=QByteArray("Private counter As Variant\nPublic exposed As String\nFunction External() As Variant\nExternal=SharedNext()\nEnd Function\nFunction StandardValue() As Variant\nStandardValue=ReadStandard()\nEnd Function\nFunction PrivateCall() As Variant\nPrivateCall=Advance()\nEnd Function\nFunction BadMe() As Variant\nBadMe=IllegalMe()\nEnd Function\nFunction BadMeRequery() As Variant\nBadMeRequery=IllegalMeRequery()\nEnd Function\nFunction BadMeFocus() As Variant\nBadMeFocus=IllegalMeFocus()\nEnd Function\nSub Explicit()\nMe.counter=9\nMe.exposed=\"public value\"\nMe.Tag=CStr(Me.counter)\nEnd Sub\nSub Bang()\nMe!exposed=1\nEnd Sub\nSub WriteGlobal()\n")+QByteArray::fromStdString(selected)+"=\"session\"\nEnd Sub\nFunction ReadGlobal() As Variant\nReadGlobal="+QByteArray::fromStdString(selected)+"\nEnd Function\nFunction Increment() As Variant\ncounter=counter+1\nIncrement=counter\nEnd Function\n";
    const auto call=[](AccessFormPreview& form,const QString& procedure) {
        std::vector<linaccess::VbaValue> args; const auto result=form.executeProcedure(procedure,args);
        assert(result.status==linaccess::VbaExecutionStatus::Completed); return result;
    };
    AccessFormPreview first("first","Begin Form\nEnd\n",source,db);
    AccessFormPreview second("second","Begin Form\nEnd\n",source,db);
    AccessFormPreview separate("separate","Begin Form\nEnd\n",source,otherDb);
    assert(call(first,"ReadGlobal").typedReturnValue->asText().empty());
    call(first,"WriteGlobal");
    assert(call(second,"ReadGlobal").typedReturnValue->asText()=="session");
    assert(call(separate,"ReadGlobal").typedReturnValue->asText().empty());
    assert(call(first,"External").typedReturnValue->asNumber()==1);
    assert(call(second,"External").typedReturnValue->asNumber()==2);
    std::vector<linaccess::VbaValue> unavailable;
    assert(first.executeProcedure("PrivateCall",unavailable).errorNumber==5);
    assert(first.executeProcedure("BadMe",unavailable).errorNumber==438);
    assert(first.executeProcedure("BadMeRequery",unavailable).errorNumber==438);
    assert(first.executeProcedure("BadMeFocus",unavailable).errorNumber==438);
    const auto controlDefinition=QByteArray("Begin Form\nBegin Section\nBegin TextBox\nName =\"")+QByteArray::fromStdString(selected)+"\"\nEnd\nEnd\nEnd\n";
    AccessFormPreview controlOwner("control scope",controlDefinition,source,db);
    makePreviewFormObject(controlOwner)->assignItem(selected,linaccess::VbaValue::string("control value"));
    assert(call(controlOwner,"ReadGlobal").typedReturnValue->asText()=="control value");
    assert(call(controlOwner,"StandardValue").typedReturnValue->asText()=="session");
    assert(controlOwner.close());
    QSqlQuery conflict(db);
    conflict.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','conflict',?)");
    conflict.addBindValue(QByteArray("Public Function SharedNext() As Variant\nSharedNext=999\nEnd Function\n")); assert(conflict.exec());
    assert(first.executeProcedure("External",unavailable).errorNumber==5);
    assert(conflict.exec("DELETE FROM _la_access_artifacts WHERE name='conflict'"));
    assert(call(second,"External").typedReturnValue->asNumber()==3);
    assert(call(first,"Increment").typedReturnValue->asNumber()==1);
    assert(call(first,"Increment").typedReturnValue->asNumber()==2);
    assert(call(second,"Increment").typedReturnValue->asNumber()==1);
    call(first,"Explicit");
    auto firstObject=makePreviewFormObject(first);
    assert(firstObject->member("Tag",{}).asText()=="9" && firstObject->member("exposed",{}).asText()=="public value");
    assert(makePreviewFormObject(second)->member("exposed",{}).asText().empty());
    bool hidden=false;
    try { (void)firstObject->member("counter",{}); } catch(const linaccess::VbaRuntimeError& error) { hidden=error.number==2465; }
    assert(hidden);
    std::vector<linaccess::VbaValue> none; assert(first.executeProcedure("Bang",none).errorNumber==2465);
    assert(first.close());
    assert(call(second,"ReadGlobal").typedReturnValue->asText()=="session");
    AccessFormPreview reopened("first","Begin Form\nEnd\n",source,db);
    assert(call(reopened,"Increment").typedReturnValue->asNumber()==1);
    linaccess::resetVbaProjectState(db);
    assert(call(second,"ReadGlobal").typedReturnValue->asText().empty());
    assert(call(second,"External").typedReturnValue->asNumber()==1);
    assert(second.close() && reopened.close() && separate.close());
    testNumericModuleStates();
    testNumericProcedureValues();
    testScalarProcedureValues();
}

void testNumericModuleStates() {
    using linaccess::VbaValue;
    const auto open=[](const char* connection) {
        auto db=QSqlDatabase::addDatabase("QSQLITE",QString::fromUtf8(connection)); db.setDatabaseName(":memory:"); assert(db.open());
        QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
        assert(query.exec("CREATE TABLE numeric_audit(value INTEGER)"));
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','numeric helper',?)");
        query.addBindValue(QByteArray("Public projectTotal As Long\nPublic Function NumericSharedNext() As Variant\nprojectTotal=projectTotal+40000\nNumericSharedNext=projectTotal\nEnd Function")); assert(query.exec()); return db;
    };
    auto db=open("numeric-module-values"),otherDb=open("numeric-module-other");
    const QByteArray code=R"VBA(Private small As Integer
Public exposed As Long
Function Increment() As Variant
small=small+1
Increment=small
End Function
Function External() As Variant
External=NumericSharedNext()
End Function
Sub RoundValues()
small=2.5
exposed=3.5
Me.Tag=CStr(small)&":"&CStr(exposed)
End Sub
Function OverflowSmall() As Variant
On Error Resume Next
small=32767
small=small+1
Me.Tag=CStr(Err.Number)
OverflowSmall=small
End Function
Sub Persist()
CurrentDb.Execute "INSERT INTO numeric_audit(value) VALUES("&CStr(exposed)&")"
End Sub
)VBA";
    AccessFormPreview first("numeric first","Begin Form\nEnd",code,db),second("numeric second","Begin Form\nEnd",code,db),separate("numeric separate","Begin Form\nEnd",code,otherDb);
    const auto call=[](AccessFormPreview& form,const char* procedure) {
        std::vector<VbaValue> args; const auto result=form.executeProcedure(QString::fromUtf8(procedure),args); assert(result.status==linaccess::VbaExecutionStatus::Completed); return result;
    };
    auto root=makePreviewFormObject(first);
    assert(root->member("exposed",{}).asNumber()==0);
    assert(call(first,"Increment").typedReturnValue->asNumber()==1 && call(first,"Increment").typedReturnValue->asNumber()==2);
    assert(call(second,"Increment").typedReturnValue->asNumber()==1);
    assert(call(first,"External").typedReturnValue->asNumber()==40000 && call(second,"External").typedReturnValue->asNumber()==80000);
    assert(call(separate,"External").typedReturnValue->asNumber()==40000);
    call(first,"RoundValues"); assert(root->member("Tag",{}).asText()=="2:4" && root->member("exposed",{}).asNumber()==4);
    assert(call(first,"OverflowSmall").typedReturnValue->asNumber()==32767 && root->member("Tag",{}).asText()=="6");
    call(first,"Persist"); QSqlQuery sql(db); assert(sql.exec("SELECT value FROM numeric_audit") && sql.next() && sql.value(0).toInt()==4); sql.finish();
    root->assign("exposed",VbaValue::numeric(2147483647));
    for(const auto& invalid:std::vector<std::pair<VbaValue,int>>{{VbaValue::numeric(2147483648.0),6},{VbaValue{},94},{VbaValue::string("bad"),13}}) {
        int actual=0; try { root->assign("exposed",invalid.first); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; }
        assert(actual==invalid.second && root->member("exposed",{}).asNumber()==2147483647);
    }
    root->assign("exposed",VbaValue::numeric(-2147483648.0)); assert(root->member("exposed",{}).asNumber()==-2147483648.0);
    int actual=0; try { root->member("small",{}); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; } assert(actual==2465);
    assert(makePreviewFormObject(second)->member("exposed",{}).asNumber()==0);
    assert(first.close()); actual=0; try { root->member("exposed",{}); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; } assert(actual==2450);
    AccessFormPreview reopened("numeric reopened","Begin Form\nEnd",code,db);
    assert(call(reopened,"Increment").typedReturnValue->asNumber()==1 && call(reopened,"External").typedReturnValue->asNumber()==120000);
    linaccess::resetVbaProjectState(db);
    assert(call(second,"External").typedReturnValue->asNumber()==40000 && call(separate,"External").typedReturnValue->asNumber()==80000);
    assert(second.close() && separate.close() && reopened.close());
}

void testNumericProcedureValues() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","numeric-procedure-values"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE numeric_local_audit(small INTEGER,wide INTEGER)")); assert(sql.exec("CREATE TABLE numeric_local_events(kind TEXT)"));
    const QByteArray code=R"VBA(Function SaveValues(ByVal input As Integer) As Long
Dim small As Integer
Dim wide As Long
small=input+0.5
Call AdvanceInteger(small)
wide=40000+small
CurrentDb.Execute "INSERT INTO numeric_local_audit VALUES("&CStr(small)&","&CStr(wide)&")"
SaveValues=wide+0.5
End Function
Sub AdvanceInteger(ByRef value As Integer)
value=value+1
End Sub
Sub ChangeLong(ByRef value As Long)
CurrentDb.Execute "INSERT INTO numeric_local_events VALUES('entered')"
value=value+1
End Sub
Function RejectReference() As String
Dim narrow As Integer
narrow=3
On Error Resume Next
Call ChangeLong(narrow)
RejectReference=CStr(Err.Number)&":"&CStr(narrow)
End Function
Function TemporaryReference() As Integer
Dim narrow As Integer
narrow=3
Call ChangeLong((narrow))
TemporaryReference=narrow
End Function
Function ReturnOverflow() As Integer
ReturnOverflow=32768
End Function
Sub BadLoop()
Dim i As Integer
For i=1 To 2.5
CurrentDb.Execute "INSERT INTO numeric_local_events VALUES('bad loop')"
Next i
End Sub
Sub BadEnumerator(items As Object)
Dim item As Long
CurrentDb.Execute "INSERT INTO numeric_local_events VALUES('bad enum')"
For Each item In items
Next item
End Sub
)VBA";
    AccessFormPreview form("numeric procedure","Begin Form\nBegin Section\nBegin TextBox\nName =\"display\"\nEnd\nEnd\nEnd",code,db);
    std::vector<VbaValue> args{VbaValue::numeric(2.5)};
    auto result=form.executeProcedure("SaveValues",args); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==40004 && args[0].asNumber()==2.5);
    assert(sql.exec("SELECT small,wide FROM numeric_local_audit") && sql.next() && sql.value(0).toInt()==3 && sql.value(1).toInt()==40003); sql.finish();
    for(const auto& invalid:std::vector<std::pair<VbaValue,int>>{{VbaValue{},94},{VbaValue::numeric(32768),6},{VbaValue::string("bad"),13}}) {
        args={invalid.first}; result=form.executeProcedure("SaveValues",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==invalid.second);
    }
    assert(sql.exec("SELECT COUNT(*) FROM numeric_local_audit") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    args.clear(); result=form.executeProcedure("RejectReference",args); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="13:3");
    assert(sql.exec("SELECT COUNT(*) FROM numeric_local_events") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    result=form.executeProcedure("TemporaryReference",args); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==3);
    result=form.executeProcedure("ReturnOverflow",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==6);
    result=form.executeProcedure("BadLoop",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==3251);
    args={makePreviewFormObject(form)->member("Controls",{})}; result=form.executeProcedure("BadEnumerator",args); assert(result.status==linaccess::VbaExecutionStatus::Unsupported && result.errorNumber==3251);
    assert(sql.exec("SELECT COUNT(*) FROM numeric_local_events") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    assert(form.close());
}

void testScalarProcedureValues() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","scalar-procedure-values"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE scalar_local_audit(flag INTEGER,label TEXT)")); assert(sql.exec("CREATE TABLE scalar_local_events(kind TEXT)"));
    const QByteArray code=R"VBA(Function PersistTyped(ByVal flag As Boolean,ByVal label As String) As String
Dim localFlag As Boolean
Dim localText As String
Dim q As Object
localFlag=flag
localText=label
Set q=CurrentDb.CreateQueryDef("","PARAMETERS pFlag BOOLEAN,pLabel TEXT; INSERT INTO scalar_local_audit VALUES([pFlag],[pLabel])")
q.Parameters("pFlag").Value=localFlag
q.Parameters("pLabel").Value=localText
q.Execute
PersistTyped=CStr(localFlag)&":"&localText
End Function
Sub FlipBoolean(ByRef value As Boolean)
CurrentDb.Execute "INSERT INTO scalar_local_events VALUES('entered')"
value=Not value
End Sub
Function RejectReference() As String
Dim label As String
label="False"
On Error Resume Next
Call FlipBoolean(label)
RejectReference=CStr(Err.Number)&":"&label
End Function
Function TemporaryReference() As String
Dim label As String
label="False"
Call FlipBoolean((label))
TemporaryReference=label
End Function
Function MatchingReference() As Boolean
Dim flag As Boolean
flag=0
Call FlipBoolean(flag)
MatchingReference=flag
End Function
Sub BadCounter()
Dim flag As Boolean
CurrentDb.Execute "INSERT INTO scalar_local_events VALUES('bad counter')"
For flag=0 To 1
Next flag
End Sub
)VBA";
    AccessFormPreview form("scalar procedure","Begin Form\nEnd",code,db);
    std::vector<VbaValue> args{VbaValue::string("False"),VbaValue::numeric(42)};
    auto result=form.executeProcedure("PersistTyped",args);
    assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="False:42" && args[0].kind==VbaValue::Kind::Text && args[1].kind==VbaValue::Kind::Number);
    assert(sql.exec("SELECT flag,label FROM scalar_local_audit") && sql.next() && sql.value(0).toInt()==0 && sql.value(1).toString()=="42"); sql.finish();
    const std::string label="ö 'alpha' 😀"; args={VbaValue::numeric(1),VbaValue::string(label)};
    result=form.executeProcedure("PersistTyped",args); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="True:"+label);
    assert(sql.exec("SELECT flag,label FROM scalar_local_audit ORDER BY rowid DESC") && sql.next() && sql.value(0).toInt()==-1 && sql.value(1).toString()==QString::fromStdString(label)); sql.finish();
    for(const auto& bad:std::vector<std::pair<VbaValue,int>>{{VbaValue{},94},{VbaValue::errorValue(11),13}}) {
        args={VbaValue::logical(true),bad.first}; result=form.executeProcedure("PersistTyped",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==bad.second);
    }
    assert(sql.exec("SELECT COUNT(*) FROM scalar_local_audit") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    args.clear(); result=form.executeProcedure("RejectReference",args); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="13:False");
    assert(sql.exec("SELECT COUNT(*) FROM scalar_local_events") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    result=form.executeProcedure("TemporaryReference",args); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="False");
    result=form.executeProcedure("MatchingReference",args); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Boolean && result.typedReturnValue->boolean);
    result=form.executeProcedure("BadCounter",args); assert(result.status==linaccess::VbaExecutionStatus::Unsupported && result.errorNumber==3251);
    assert(sql.exec("SELECT COUNT(*) FROM scalar_local_events") && sql.next() && sql.value(0).toInt()==2); sql.finish(); assert(form.close());
}
