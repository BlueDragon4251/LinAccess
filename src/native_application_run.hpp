#pragma once
#include "vba_expression.hpp"
class AccessFormPreview;
linaccess::VbaValue runNativeApplicationProcedure(AccessFormPreview& preview,const std::vector<linaccess::VbaValue>& arguments);

void runNativeMacroFunction(AccessFormPreview& preview,const std::string& expression);
