#pragma once
#include "vba_calendar.hpp"
#include <chrono>
#include <ctime>
#include <filesystem>
#include <iomanip>
#include <locale>
#include <sstream>
#include <cerrno>
#include <fcntl.h>
#include <sys/stat.h>
namespace linaccess {
// Shared local-time conversion for native FileDateTime and FSO file metadata.
inline VbaValue nativeFileDateFromUnix(std::int64_t seconds) {
    const auto stamp=static_cast<std::time_t>(seconds);
    if(static_cast<std::int64_t>(stamp)!=seconds) throw VbaRuntimeError(6,"File timestamp exceeds native time range");
    std::tm local{};
    if(!localtime_r(&stamp,&local) || local.tm_year+1900<100 || local.tm_year+1900>9999)
        throw VbaRuntimeError(6,"File timestamp exceeds VBA Date");
    std::ostringstream formatted; formatted.imbue(std::locale::classic());
    formatted<<std::put_time(&local,"%Y-%m-%d %H:%M:%S");
    return vbaDateValue(VbaValue::string(formatted.str()));
}
inline VbaValue nativeFileDate(std::filesystem::file_time_type modified) {
    const auto system=std::chrono::floor<std::chrono::seconds>(std::filesystem::file_time_type::clock::to_sys(modified));
    return nativeFileDateFromUnix(system.time_since_epoch().count());
}
// Linux reports actual birth/access times; unavailable birth time is never
// replaced by inode change time or the file's modification time.
inline VbaValue nativeFileTimestamp(const std::filesystem::path& path,bool created,bool folder) {
    struct statx metadata{};
    const auto wanted=created ? STATX_BTIME : STATX_ATIME;
    if(::statx(AT_FDCWD,path.c_str(),AT_SYMLINK_NOFOLLOW|AT_NO_AUTOMOUNT,STATX_TYPE|wanted,&metadata)!=0) {
        if(errno==ENOSYS || errno==EOPNOTSUPP) throw VbaRuntimeError(3251,"Native filesystem timestamp is unavailable");
        if(errno==EACCES || errno==EPERM) throw VbaRuntimeError(70,"Filesystem timestamp permission denied");
        if(errno==ENOENT || errno==ENOTDIR) throw VbaRuntimeError(folder ? 76 : 53,"Filesystem object unavailable");
        throw VbaRuntimeError(75,"Native filesystem timestamp failed");
    }
    if(!(metadata.stx_mask&STATX_TYPE)) throw VbaRuntimeError(3251,"Filesystem timestamp object kind is unavailable");
    if(S_ISLNK(metadata.stx_mode)) throw VbaRuntimeError(3251,"Filesystem link timestamp semantics require reference validation");
    if(folder ? !S_ISDIR(metadata.stx_mode) : !S_ISREG(metadata.stx_mode)) throw VbaRuntimeError(3251,"Filesystem timestamp object kind changed");
    if(!(metadata.stx_mask&wanted)) throw VbaRuntimeError(3251,"Requested filesystem timestamp is unavailable");
    return nativeFileDateFromUnix(created ? metadata.stx_btime.tv_sec : metadata.stx_atime.tv_sec);
}
}
