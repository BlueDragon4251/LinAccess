#pragma once

#include <string>
#include <string_view>
#include <vector>

namespace linaccess {
struct CsvField {
    std::string value;
    bool quoted{};
};
// Retains whether each value was quoted: required to distinguish an unquoted
// MDBTools NULL marker from a literal string containing that same marker.
std::vector<std::vector<CsvField>> parseCsvDetailed(std::string_view data);
// CSV parser: escaped double-quotes and embedded newlines are supported.
// Returns empty vector for empty input.
std::vector<std::vector<std::string>> parseCsv(std::string_view data);
}
