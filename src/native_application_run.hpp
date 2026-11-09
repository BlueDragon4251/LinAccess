#pragma once
#include "vba_expression.hpp"
class AccessFormPreview;
linaccess::VbaValue runNativeApplicationProcedure(AccessFormPreview& preview,const std::vector<linaccess::VbaValue>& arguments);

void runNativeMacroFunction(AccessFormPreview& preview,const std::string& expression);

linaccess::VbaValue invokeNativePreviewCommand(AccessFormPreview& preview,std::string_view name,const std::vector<linaccess::VbaValue>& arguments);

// Share the actual Qt execution lifetime with standalone module/macro frames.
void enterNativeUiExecution();
void leaveNativeUiExecution() noexcept;
