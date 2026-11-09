#include "vba_array_return_fixture.hpp"
#include "vba_array.hpp"
#include "native_standard_module_host.hpp"
#include "jet_sql_functions.hpp"
#include "access_query.hpp"
#include <QDialog>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlQuery>
#include <QSqlError>
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_property_fixture.hpp"
#include "vba_parameterless_fixture.hpp"
#include "vba_default_property_fixture.hpp"
#include "vba_byref_property_fixture.hpp"
#include "vba_optional_property_fixture.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlDatabase>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
void exerciseNativeProperties(QSqlDatabase db){using namespace linaccess;
    const QByteArray layout="Begin Form\nBegin Section\nEnd\nEnd\n";
    AccessFormPreview form("Property technical fixture",layout,QByteArray::fromStdString(nativePropertyFixture),db);form.show();QApplication::processEvents();
    auto object=makePreviewFormObject(form);std::vector<VbaValue> arguments;
    auto result=form.executeProcedure("FormProbe",arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form property: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==20);
    assert(object->member("ReadCount",{}).number==1);
    object->assign("Stored",VbaValue::longInteger(17));assert(object->member("ReadCount",{}).number==1);assert(object->member("Stored",{}).number==17);
    object->assignIndexedMember("Indexed",{VbaValue::longInteger(3)},VbaValue::longInteger(31));assert(object->member("Indexed",{VbaValue::longInteger(3)}).number==31);
    const auto retained=object->objectMember("ObjectValue",{}).object;assert(retained && retained->member("Item",{VbaValue::string("entry")}).number==8);
    object->assignObject("ObjectValue",VbaValue::reference(nullptr));assert(!object->objectMember("ObjectValue",{}).object);
    bool failed=false;try{(void)object->member("Failure",{});}catch(const VbaRuntimeError& e){failed=e.number==11;}assert(failed);
    failed=false;try{object->assign("ReadCount",VbaValue::longInteger(99));}catch(const VbaRuntimeError& e){failed=e.number==383;}assert(failed);
    assert(!nativeVbaPropertySignature(nativePropertyFixture,"Secret",VbaPropertyAccess::Get,true));
    AccessFormPreview other("Separate property instance",layout,QByteArray::fromStdString(nativePropertyFixture),db);other.show();QApplication::processEvents();assert(makePreviewFormObject(other)->member("Stored",{}).number==0);assert(other.close());
    assert(form.close());failed=false;try{(void)object->member("Stored",{});}catch(const VbaRuntimeError& e){failed=e.number==2450;}assert(failed);

    AccessFormPreview bare("Parameterless declarations",layout,QByteArray::fromStdString(nativeParameterlessFixture),db);bare.show();QApplication::processEvents();arguments.clear();
    for(int call=1;call<=2;++call){result=bare.executeProcedure("FormProbe",arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call);}
    auto bareObject=makePreviewFormObject(bare);assert(bareObject->member("Value",{}).number==2 && bareObject->member("EmptyValue",{}).kind==VbaValue::Kind::Empty);assert(bare.close());

    const auto widgetSource=QByteArray::fromStdString(nativeDefaultPropertyFixture)+R"VBA(
Public Property Get Editor As Object
getterReads=getterReads+1
Set Editor=Me.Actual
End Property
Public Property Get EditorAt(ByVal index As Long) As Object
getterReads=getterReads+1
Set EditorAt=Me.Controls(index)
End Property
Public Function WidgetProbe As Long
Me.Editor=17
Editor=18
EditorAt(0)=19
WidgetProbe=Actual
End Function
Public Property Get Closing As Object
Set Closing=Me.Actual
DoCmd.Close
End Property
)VBA";
    const QByteArray widgetLayout=R"LAYOUT(Begin Form
Begin Section
Begin TextBox
Name ="Actual"
End
End
End
)LAYOUT";
    AccessFormPreview defaults("Default property forwarding",widgetLayout,widgetSource,db);defaults.show();QApplication::processEvents();auto* actual=defaults.findChild<QLineEdit*>("access-control-Actual");assert(actual);auto defaultObject=makePreviewFormObject(defaults);
    result=defaults.executeProcedure("DefaultProbe",arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form default property: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==26);
    assert(defaultObject->member("Bag",{VbaValue::string("value")}).number==7);defaultObject->assignIndexedMember("Bag",{VbaValue::string("value")},VbaValue::longInteger(9));assert(defaultObject->member("Bag",{VbaValue::string("value")}).number==9);
    const auto before=defaultObject->member("ReadCount",{}).number;result=defaults.executeProcedure("WidgetProbe",arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Widget default: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==19 && actual->text()=="19" && defaultObject->member("ReadCount",{}).number==before+3);
    defaultObject->assign("Editor",VbaValue::longInteger(21));defaultObject->assignIndexedMember("EditorAt",{VbaValue::longInteger(0)},VbaValue::longInteger(22));assert(actual->text()=="22" && defaultObject->member("ReadCount",{}).number==before+5);
    failed=false;try{defaultObject->assign("Closing",VbaValue::longInteger(99));}catch(const VbaRuntimeError& e){failed=e.number==2450;}assert(failed && actual->text()=="22" && !defaults.isVisible());

    const auto byrefSource=QByteArray::fromStdString(nativeByRefPropertyFixture)+R"VBA(
Public Property Get EditorRef(ByRef index As Long) As Object
index=index+1
Set EditorRef=Me.Actual
End Property
Public Function WidgetRefProbe As Long
Dim index As Long
index=0
EditorRef(index)=23
WidgetRefProbe=Actual+index
End Function
)VBA";
    AccessFormPreview byref("Owned ByRef property indices",widgetLayout,byrefSource,db);byref.show();QApplication::processEvents();auto* referenceText=byref.findChild<QLineEdit*>("access-control-Actual");assert(referenceText);
    for(const auto& probe:std::vector<std::pair<std::string,int>>{{"LetProbe",277},{"GetProbe",22},{"ArrayProbe",330},{"ObjectProbe",319},{"ErrorProbe",211},{"WidgetRefProbe",24}}){arguments.clear();result=byref.executeProcedure(QString::fromStdString(probe.first),arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form ByRef property "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}assert(referenceText->text()=="23");
    failed=false;try{(void)makePreviewFormObject(byref)->member("Counter",{VbaValue::longInteger(1)});}catch(const VbaRuntimeError& e){failed=e.number==3251;}assert(failed && byref.close());

    const auto optionalSource=QByteArray::fromStdString(nativeOptionalPropertyFixture)+R"VBA(
Public Function QualifiedProbe As Long
Me.Cell(,8)=7
QualifiedProbe=Written+Me.Cell(,8)+Me.Cell(col:=8)
End Function
Public Function QualifiedManyProbe As Long
Me.Many(1,2,3)=7
QualifiedManyProbe=Written+Me.Many(1,2,3)
End Function
Public Function NamedManyProbe As Long
On Error Resume Next
NamedManyProbe=Me.Many(values:=7)
NamedManyProbe=Err.Number
End Function
)VBA";
    AccessFormPreview optional("Optional property indices",layout,optionalSource,db);optional.show();QApplication::processEvents();
    for(const auto& probe:std::vector<std::pair<std::string,int>>{{"DefaultsProbe",487},{"HoleProbe",543},{"EmptyCallProbe",457},{"OwnedProbe",3067},{"ManyProbe",673},{"EmptyManyProbe",7},{"OptionalObjectProbe",439},{"ManyObjectProbe",249},{"GuardedValuesProbe",10251},{"QualifiedProbe",543},{"QualifiedManyProbe",673},{"NamedManyProbe",3251}}){arguments.clear();result=optional.executeProcedure(QString::fromStdString(probe.first),arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form optional property "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    auto optionalObject=makePreviewFormObject(optional);optionalObject->assign("Cell",VbaValue::longInteger(7));assert(optionalObject->member("Written",{}).number==457);assert(optionalObject->member("Cell",{VbaValue::missing(),VbaValue::longInteger(8)}).number==28);optionalObject->assignIndexedMember("Many",{VbaValue::longInteger(1),VbaValue::longInteger(2),VbaValue::longInteger(3)},VbaValue::longInteger(7));assert(optionalObject->member("Written",{}).number==367);assert(optional.close());
    AccessFormPreview arrays("Typed array return fixture",layout,QByteArray::fromStdString(nativeArrayReturnFixture),db);arrays.show();QApplication::processEvents();
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"LinArrayReadProbe",297},{"LinArrayCopyProbe",1299},{"LinArrayQueueProbe",14},{ "LinArrayEachProbe",3901},{"LinArrayEachObjects",19},{"LinArrayEachCopy",3999},{"LinArrayEachNested",9}}){arguments.clear();result=arrays.executeProcedure(probe.first,arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form array "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    auto arrayObject=makePreviewFormObject(arrays);const auto getCount=arrayObject->member("LinArrayGetCount",{}).number;
    assert(arrayObject->member("LinArrayReturnMatrix",{VbaValue::longInteger(3),VbaValue::longInteger(0)}).number==30 && arrayObject->member("LinArrayGetCount",{}).number==getCount+1);
    const auto retainedMatrix=arrayObject->member("LinArrayReturnMatrix",{});assert(arrays.close());assert(vbaArrayElement(retainedMatrix,{VbaValue::longInteger(2),VbaValue::longInteger(-1)}).number==21);
    QSqlQuery arraySql(db);arraySql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('module','TypedArrayReturnFixture',?)");arraySql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromStdString(nativeArrayReturnFixture)}}).toJson(QJsonDocument::Compact)));assert(arraySql.exec());arraySql.finish();installJetSqlFunctions(db);
    const auto sqlCount=[&](){assert(arraySql.exec("SELECT LinArrayCallCount()") && arraySql.next());const auto count=arraySql.value(0).toInt();arraySql.finish();return count;};assert(sqlCount()==0);
    assert(!arraySql.exec("SELECT LinArrayReturnLongs()") && arraySql.lastError().text().contains("3251"));arraySql.finish();assert(sqlCount()==0);
    const auto scalar=executeAccessQueryPlan(db,compileAccessQuery("SELECT LinArrayCallCount() AS NativeCount"),{});assert(scalar.rows.size()==1 && scalar.rows[0][0].toInt()==0);
    QDialog moduleOwner;NativeStandardModuleHost moduleHost(&moduleOwner,db,[]{});result=executeVbaProcedure(nativeArrayReturnFixture,"LinArrayReturnLongs",moduleHost);assert(result.status==VbaExecutionStatus::Completed && vbaArrayElementType(*result.typedReturnValue)=="long" && sqlCount()==1);
    assert(arraySql.exec("DELETE FROM main._la_objects WHERE kind='module' AND name='TypedArrayReturnFixture'"));arraySql.finish();installJetSqlFunctions(db);

}
