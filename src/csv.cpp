#include "csv.hpp"

#include <stdexcept>

namespace linaccess {
namespace {
std::vector<std::vector<CsvField>> parseDelimited(std::string_view input,char delimiter,bool strict,std::size_t cellLimit=0) {
    std::vector<std::vector<CsvField>> rows;
    if (input.empty()) return rows;
    std::vector<CsvField> row;
    std::string field;
    bool quoted = false;
    bool fieldQuoted = false;
    bool started = false;
    bool closedQuote = false;
    std::size_t cells=0;
    const auto appendField=[&] {
        if(cellLimit && ++cells>cellLimit) throw std::length_error("Delimited input exceeds the cell limit");
        row.push_back({field,fieldQuoted});
    };
    for (std::size_t i = 0; i < input.size(); ++i) {
        const char ch = input[i];
        started = true;
        if(strict && closedQuote && ch!=delimiter && ch!='\r' && ch!='\n')
            throw std::runtime_error("Delimited input has characters after a quoted field");
        if (ch == '"') {
            if(strict && !quoted && !field.empty()) throw std::runtime_error("Delimited input has an unescaped quote");
            if (!quoted && field.empty()) fieldQuoted = true;
            if (quoted && i + 1 < input.size() && input[i + 1] == '"') {
                field.push_back('"');
                ++i;
            } else {
                quoted = !quoted;
                if(strict && !quoted) closedQuote=true;
            }
        } else if (ch == delimiter && !quoted) {
            appendField();
            field.clear();
            fieldQuoted = false;
            closedQuote = false;
        } else if ((ch == '\r' || ch == '\n') && !quoted) {
            if (ch == '\r' && i + 1 < input.size() && input[i + 1] == '\n') ++i;
            appendField();
            rows.push_back(row);
            row.clear();
            field.clear();
            fieldQuoted = false;
            closedQuote = false;
            started = false;
        } else {
            field.push_back(ch);
        }
    }
    if (quoted) throw std::runtime_error("CSV contains an unterminated quoted field");
    if (started) {
        appendField();
        rows.push_back(row);
    }
    return rows;
}
}
std::vector<std::vector<CsvField>> parseCsvDetailed(std::string_view input) { return parseDelimited(input,',',false); }
std::vector<std::vector<CsvField>> parseTsvDetailed(std::string_view input) { return parseDelimited(input,'\t',true,1000000); }

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
