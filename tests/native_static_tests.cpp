#include "access_form_preview.hpp"
#include "native_standard_module_host.hpp"
#include "vba_static_fixture.hpp"
#include <QApplication>
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
}
