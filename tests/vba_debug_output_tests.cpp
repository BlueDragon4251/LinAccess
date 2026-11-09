#include "vba_debug_output.hpp"
#include "vba_runtime.hpp"
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
namespace {
struct Host final:linaccess::VbaHost {
    int messages{};
    std::optional<std::string> field(std::string_view)override{return {};}
    void setField(std::string_view,std::optional<std::string>)override{assert(false);}
    void saveRecord()override{assert(false);}
    void message(std::string_view)override{++messages;}
};
}
void testVbaDebugOutput(){using namespace linaccess;Host host;clearVbaDebugOutput();
    const std::string code=R"VBA(Private Function Once(ByRef n As Long) As Long
n=n+1
Once=n
End Function
Public Function Probe() As Long
Dim n As Long
Debug.Print "start";1;-2;True;False;Null;Empty;CVErr(11)
Debug.Print "A","B"
Debug.Print "x";Spc(2);"y";Tab(8);"z";
Debug.Print "end"
Debug.Print Once(n);Once(n)
Debug.Print (1)=1
Probe=n
End Function)VBA";
    auto result=executeVbaProcedure(code,"Probe",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->number==2 && host.messages==0);
    const auto snapshot=vbaDebugOutputSnapshot();const auto expected=std::string("start 1  -2 TrueFalseNullError 11\nA")+std::string(13,' ')+"B\nx  y   zend\n 1  2 \nTrue\n";assert(snapshot.text==expected && snapshot.column==1 && !snapshot.truncated);assert(vbaDebugOutputSnapshot().text==snapshot.text && vbaDebugOutputRevision()==snapshot.revision);
    clearVbaDebugOutput();result=executeVbaProcedure("Function Recover() As Long\nOn Error Resume Next\nDebug.Print \"before\";CStr(Null);\"unreached\"\nDebug.Print \"after\"\nRecover=Err.Number\nEnd Function","Recover",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==94 && vbaDebugOutputSnapshot().text=="beforeafter\n");
    clearVbaDebugOutput();result=executeVbaProcedure("Sub Bad()\nMsgBox \"unreached\"\nDebug.Print #1, \"file\"\nEnd Sub","Bad",host);assert(result.status==VbaExecutionStatus::Unsupported && host.messages==0 && vbaDebugOutputSnapshot().text.empty());
    result=executeVbaProcedure("Sub Position()\nDebug.? \"a\";Tab;\"b\";\nDebug.Print Tab(1);Spc(-1);\"c\"\nEnd Sub","Position",host);assert(result.status==VbaExecutionStatus::Completed && vbaDebugOutputSnapshot().text==std::string("a")+std::string(13,' ')+"b\nc\n");
    clearVbaDebugOutput();appendVbaDebugOutput("ä"+std::string(1024*1024-1,'x'));auto bounded=vbaDebugOutputSnapshot();assert(bounded.truncated && bounded.text.size()<=1024*1024 && vbaStringCodeUnits(bounded.text)==bounded.text.size() && bounded.column==1024*1024+1);appendVbaDebugOutput("\n");assert(vbaDebugOutputSnapshot().text.size()>1000000 && vbaDebugOutputColumn()==1);
    clearVbaDebugOutput();std::string manyLines;for(int i=0;i<9000;++i)manyLines+="x\xe2\x80\xa9";appendVbaDebugOutput(manyLines);bounded=vbaDebugOutputSnapshot();assert(bounded.truncated && vbaStringCodeUnits(bounded.text)<=8192*2);
    const auto revision=vbaDebugOutputRevision();clearVbaDebugOutput();assert(vbaDebugOutputSnapshot().text.empty() && !vbaDebugOutputSnapshot().truncated && vbaDebugOutputColumn()==1 && vbaDebugOutputRevision()!=revision);
}
