#pragma once
#include <QString>
#include <functional>
class QWidget;
namespace linaccess {
// In-process Qt keyboard adapter; never injects into another application.
void sendNativeKeys(QWidget* scope,const QString& keys,bool wait,
                    std::function<void(int)> asynchronousFailure={},
                    std::function<QWidget*()> focusResolver={});
}
