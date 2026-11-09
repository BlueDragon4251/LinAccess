#pragma once
#include "vba_expression.hpp"
#include <optional>
namespace linaccess {
struct VbaDebugOutputSnapshot {std::string text;std::uint64_t revision{};std::size_t column{1};bool truncated{};};
void clearVbaDebugOutput();
void appendVbaDebugOutput(std::string_view text);
std::uint64_t vbaDebugOutputRevision() noexcept;
std::size_t vbaDebugOutputColumn() noexcept;
VbaDebugOutputSnapshot vbaDebugOutputSnapshot();
struct VbaDebugPrintItem {
    enum class Kind { Expression,Spaces,Tab,Zone } kind{Kind::Expression};
    std::string expression;char delimiter{};
};
std::optional<std::vector<VbaDebugPrintItem>> parseVbaDebugPrint(std::string_view text,
    const std::function<bool(std::string_view)>& validExpression);
void executeVbaDebugPrint(const std::vector<VbaDebugPrintItem>& items,
    const std::function<VbaValue(std::string_view)>& evaluate);
}
