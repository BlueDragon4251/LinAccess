#pragma once
#include "vba_expression.hpp"
#include <QString>
#include <functional>
class QWidget;
namespace linaccess {
void nativeMacroMessageBox(QWidget*,const std::vector<VbaValue>&,
    const std::function<VbaValue(const QString&)>&,const std::function<void()>&);
}
