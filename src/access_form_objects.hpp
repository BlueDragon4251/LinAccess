#pragma once
#include "vba_expression.hpp"
class AccessFormPreview;
// Non-owning GUI-thread lookup of the active native document.
AccessFormPreview* activePreviewDocument();
std::shared_ptr<linaccess::VbaObject> makePreviewFormObject(AccessFormPreview& preview);
std::shared_ptr<linaccess::VbaObject> makePreviewFormsObject(AccessFormPreview& preview);

std::shared_ptr<linaccess::VbaObject> makePreviewApplicationObject(AccessFormPreview& preview);

std::shared_ptr<linaccess::VbaObject> makePreviewReportsObject(AccessFormPreview& preview);

class QSqlDatabase;
void connectPreviewSqlObjects(QSqlDatabase db);

linaccess::VbaValue dispatchPreviewDocumentCommand(AccessFormPreview& preview,std::string_view command,const std::vector<linaccess::VbaValue>& arguments);
