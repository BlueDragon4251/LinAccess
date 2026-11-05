#pragma once
#include "vba_calendar.hpp"
#include <chrono>
#include <ctime>
#include <filesystem>
#include <iomanip>
#include <locale>
#include <sstream>
namespace linaccess {
// Shared local-time conversion for native FileDateTime and FSO file metadata.
inline VbaValue nativeFileDate(std::filesystem::file_time_type modified) {
    const auto system=std::chrono::floor<std::chrono::seconds>(std::filesystem::file_time_type::clock::to_sys(modified));
    const auto stamp=std::chrono::system_clock::to_time_t(system);
    std::tm local{};
    if(!localtime_r(&stamp,&local) || local.tm_year+1900<100 || local.tm_year+1900>9999)
        throw VbaRuntimeError(6,"File timestamp exceeds VBA Date");
    std::ostringstream formatted; formatted.imbue(std::locale::classic());
    formatted<<std::put_time(&local,"%Y-%m-%d %H:%M:%S");
    return vbaDateValue(VbaValue::string(formatted.str()));
}
}
