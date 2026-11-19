#pragma once
#include <QString>
#include <functional>
#include <optional>
namespace linaccess {
class NativeProjectCloseGuard {
public:
    using Completion=std::function<void(bool)>;
    virtual ~NativeProjectCloseGuard()=default;
    // A disengaged result vetoes close. An accepted participant receives the
    // final outcome after every participant and runtime form has been checked.
    // Empty connection means application-wide.
    virtual std::optional<Completion> prepareProjectClose(const QString& connection)=0;
};
}
