#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_property_fixture.hpp"
#include <QApplication>
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
}
