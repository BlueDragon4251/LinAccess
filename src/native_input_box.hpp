#pragma once
#include "vba_runtime.hpp"
class QWidget;
namespace linaccess {
std::string nativeInputBox(QWidget* owner,const VbaInputBoxArguments& arguments);
}
