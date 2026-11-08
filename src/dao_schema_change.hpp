#pragma once
#include "vba_expression.hpp"
#include <utility>
namespace linaccess {
// The schema mutation committed; a subsequent connection-restoration failure
// must never be presented as an uncommitted, safely repeatable operation.
class DaoStoredSchemaChange final : public VbaRuntimeError {
public:
    explicit DaoStoredSchemaChange(std::string message):VbaRuntimeError(3251,std::move(message)) {}
};
}
