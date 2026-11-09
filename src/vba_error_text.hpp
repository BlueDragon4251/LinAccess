#pragma once
#include <string>
#include <utility>
namespace linaccess {
class VbaHost;
int lastRaisedVbaError(VbaHost& host);
// Bounded invariant native messages. This is not a Windows locale catalog.
inline std::string vbaErrorText(int code) {
    if(code==0) return {};
    static constexpr std::pair<int,const char*> descriptions[]{
        {3,"Return without GoSub"},{5,"Invalid procedure call or argument"},
        {6,"Overflow"},{7,"Out of memory"},{9,"Subscript out of range"},
        {10,"This array is fixed or temporarily locked"},{11,"Division by zero"},
        {13,"Type mismatch"},{14,"Out of string space"},{18,"User interrupt occurred"},
        {20,"Resume without error"},{28,"Out of stack space"},{35,"Sub or Function not defined"},
        {52,"Bad file name or number"},{53,"File not found"},{55,"File already open"},
        {58,"File already exists"},{61,"Disk full"},{62,"Input past end of file"},
        {70,"Permission denied"},{75,"Path/File access error"},{76,"Path not found"},
        {91,"Object variable or With block variable not set"},{93,"Invalid pattern string"},
        {94,"Invalid use of Null"},{424,"Object required"},
        {438,"Object doesn't support this property or method"},{448,"Named argument not found"},
        {449,"Argument not optional"},{450,"Wrong number of arguments or invalid property assignment"},
        {457,"This key is already associated with an element of this collection"}};
    for(const auto& [number,text]:descriptions) if(number==code) return text;
    return "Application-defined or object-defined error.";
}
}
