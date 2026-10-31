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
        const auto helper=QByteArray("Private hits As Variant\nPublic Function SharedNext() As Variant\nSharedNext=Advance()\nEnd Function\nPrivate Function Advance() As Variant\nhits=hits+1\nAdvance=hits\nEnd Function\nPublic Function ReadStandard() As Variant\nReadStandard=")+QByteArray::fromStdString(selected)+"\nEnd Function\nPublic Function IllegalMe() As Variant\nIllegalMe=Me.Tag\nEnd Function\n";
        add("SHARED HELPER",helper); return db;
    };
    auto db=open("original-module-state"); auto otherDb=open("separate-module-state");
    const auto source=QByteArray("Private counter As Variant\nPublic exposed As String\nFunction External() As Variant\nExternal=SharedNext()\nEnd Function\nFunction StandardValue() As Variant\nStandardValue=ReadStandard()\nEnd Function\nFunction PrivateCall() As Variant\nPrivateCall=Advance()\nEnd Function\nFunction BadMe() As Variant\nBadMe=IllegalMe()\nEnd Function\nSub Explicit()\nMe.counter=9\nMe.exposed=\"public value\"\nMe.Tag=CStr(Me.counter)\nEnd Sub\nSub Bang()\nMe!exposed=1\nEnd Sub\nSub WriteGlobal()\n")+QByteArray::fromStdString(selected)+"=\"session\"\nEnd Sub\nFunction ReadGlobal() As Variant\nReadGlobal="+QByteArray::fromStdString(selected)+"\nEnd Function\nFunction Increment() As Variant\ncounter=counter+1\nIncrement=counter\nEnd Function\n";
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
}
