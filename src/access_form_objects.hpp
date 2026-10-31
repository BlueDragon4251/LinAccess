#pragma once
#include "vba_expression.hpp"
class AccessFormPreview;
std::shared_ptr<linaccess::VbaObject> makePreviewFormObject(AccessFormPreview& preview);
std::shared_ptr<linaccess::VbaObject> makePreviewFormsObject(AccessFormPreview& preview);
