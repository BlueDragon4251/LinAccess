#include "csv.hpp"

#include <stdexcept>

namespace linaccess {
std::vector<std::vector<CsvField>> parseCsvDetailed(std::string_view input) {
    std::vector<std::vector<CsvField>> rows;
    if (input.empty()) return rows;
    std::vector<CsvField> row;
    std::string field;
    bool quoted = false;
    bool fieldQuoted = false;
    bool started = false;
    for (std::size_t i = 0; i < input.size(); ++i) {
        const char ch = input[i];
        started = true;
        if (ch == '"') {
            if (!quoted && field.empty()) fieldQuoted = true;
            if (quoted && i + 1 < input.size() && input[i + 1] == '"') {
                field.push_back('"');
                ++i;
            } else {
                quoted = !quoted;
            }
        } else if (ch == ',' && !quoted) {
            row.push_back({field, fieldQuoted});
            field.clear();
            fieldQuoted = false;
        } else if ((ch == '\r' || ch == '\n') && !quoted) {
            if (ch == '\r' && i + 1 < input.size() && input[i + 1] == '\n') ++i;
            row.push_back({field, fieldQuoted});
            rows.push_back(row);
            row.clear();
            field.clear();
            fieldQuoted = false;
            started = false;
        } else {
            field.push_back(ch);
        }
    }
    if (quoted) throw std::runtime_error("CSV contains an unterminated quoted field");
    if (started) {
        row.push_back({field, fieldQuoted});
        rows.push_back(row);
    }
    return rows;
}

std::vector<std::vector<std::string>> parseCsv(std::string_view input) {
    const auto detailed = parseCsvDetailed(input);
    std::vector<std::vector<std::string>> rows;
    rows.reserve(detailed.size());
    for (const auto& cells : detailed) {
        std::vector<std::string> row;
        row.reserve(cells.size());
        for (const auto& cell : cells) row.push_back(cell.value);
        rows.push_back(std::move(row));
    }
    return rows;
}
}
