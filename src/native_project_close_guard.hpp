#pragma once
#include <QString>
namespace linaccess {
// Native UI participants can refuse project/application close before any form
// unloads. Empty connection means application-wide; no app dependency in forms.
class NativeProjectCloseGuard {
public:
    virtual ~NativeProjectCloseGuard()=default;
    virtual bool allowProjectClose(const QString& connection)=0;
};
}
