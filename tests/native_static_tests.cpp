#include "vba_def_type_fixture.hpp"
#include <iostream>
#include "access_form_preview.hpp"
#include "native_standard_module_host.hpp"
#include "vba_static_fixture.hpp"
#include <QApplication>
#include <QLineEdit>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
void exerciseNativeStaticProcedures(QSqlDatabase db){using namespace linaccess;
    const QByteArray layout="Begin Form\nBegin Section\nEnd\nEnd\n";const auto source=QByteArray::fromStdString(nativeStaticFixture);
    AccessFormPreview first("Static instance one",layout,source,db),second("Static instance two",layout,source,db);first.show();second.show();QApplication::processEvents();std::vector<VbaValue> args;
    auto result=first.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");result=first.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="2:4:6:2:A2");result=second.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");
    result=first.executeProcedure("Implicit",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:1:1:3:1");result=first.executeProcedure("Implicit",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="2:2:2:6:2");assert(first.close() && second.close());
    QDialog moduleEditor;
    for(int call=1;call<=2;++call){NativeStandardModuleHost host(&moduleEditor,db,[]{});result=executeVbaProcedure(nativeStaticFixture,"Trailing",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call);}
    AccessFormPreview reopened("Static new instance",layout,source,db);reopened.show();QApplication::processEvents();result=reopened.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");assert(reopened.close());

    const QByteArray implicitLayout="Begin Form\nBegin Section\nBegin TextBox\nName =\"Actual\"\nEnd\nEnd\nEnd\n";
    const QByteArray implicitCode="Public Function ImplicitProbe As Long\ncount=count+1\nActual=count\nImplicitProbe=Actual+count\nEnd Function";
    AccessFormPreview ordinary("Transient locals",implicitLayout,implicitCode,db);ordinary.show();QApplication::processEvents();auto* actual=ordinary.findChild<QLineEdit*>("access-control-Actual");assert(actual);
    for(int call=0;call<2;++call){result=ordinary.executeProcedure("ImplicitProbe",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==2 && actual->text()=="1");}assert(ordinary.close());
    AccessFormPreview typed("Def typed instance",layout,QByteArray::fromStdString(nativeDefTypeFixture),db);typed.show();QApplication::processEvents();
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"LValue",7},{"ArrayProbe",819533},{"ImplicitArrayProbe",819509},{"MismatchProbe",1300},{"OverrideProbe",805},{"OverflowProbe",602},{"LStatic",1},{"LStatic",2},{"LOrdinary",1},{"LOrdinary",1}}){args.clear();result=typed.executeProcedure(probe.first,args);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form Def "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    args.clear();result=typed.executeProcedure("ParameterProbe",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="value:0:2");args.clear();result=typed.executeProcedure("PropertyProbe",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="cell");assert(typed.close());

}
