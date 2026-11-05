#pragma once
#include "vba_expression.hpp"
#include "native_file_handles.hpp"
#include <filesystem>
#include <cerrno>
#include <unistd.h>
#if defined(__linux__)
#include <sys/syscall.h>
#include <fcntl.h>
#include <linux/fs.h>
#endif
namespace linaccess {
inline std::filesystem::path nativeFilePath(const VbaValue& value) {
    const auto text=coerceVbaModuleValue(value,"string").asText();
    if(text.find('\0')!=std::string::npos || text.find_first_of("*?")!=std::string::npos)
        throw VbaRuntimeError(52,"Invalid native file pathname");
    if(text.find('\\')!=std::string::npos || text.starts_with("//") || (text.size()>1 && text[1]==':'))
        throw VbaRuntimeError(3251,"Windows file path mapping is unavailable");
    return std::filesystem::path(text);
}
inline void nativeFileCopy(const VbaValue& sourceValue,const VbaValue& destinationValue,bool overwrite=true) {
    namespace fs=std::filesystem;
    const auto source=nativeFilePath(sourceValue),destination=nativeFilePath(destinationValue);
    if(source.empty()) throw VbaRuntimeError(53,"FileCopy source not found");
    if(destination.empty()) throw VbaRuntimeError(52,"Invalid FileCopy destination");
    const auto fail=[&](const std::error_code& reason,bool sourceOperation) {
        if(reason==std::errc::file_exists && !overwrite) throw VbaRuntimeError(58,"Copy destination already exists");
        if(reason==std::errc::permission_denied || reason==std::errc::operation_not_permitted)
            throw VbaRuntimeError(70,"FileCopy permission denied");
        if(reason==std::errc::invalid_argument || reason==std::errc::filename_too_long)
            throw VbaRuntimeError(52,"Invalid FileCopy pathname");
        if(reason==std::errc::no_such_file_or_directory || reason==std::errc::not_a_directory) {
            std::error_code parentError;
            const auto parent=source.has_parent_path() ? source.parent_path() : fs::path(".");
            const bool exists=fs::is_directory(parent,parentError);
            if(parentError==std::errc::permission_denied) throw VbaRuntimeError(70,"FileCopy permission denied");
            throw VbaRuntimeError(sourceOperation && exists ? 53 : 76,"FileCopy file or path not found");
        }
        throw VbaRuntimeError(75,"Native FileCopy failed");
    };
    std::error_code problem;
    const auto sourceStatus=fs::symlink_status(source,problem);
    if(problem) fail(problem,true);
    if(fs::is_symlink(sourceStatus)) throw VbaRuntimeError(3251,"FileCopy symbolic-link source unsupported");
    if(!fs::is_regular_file(sourceStatus)) throw VbaRuntimeError(75,"FileCopy requires a regular source file");
    const auto destinationStatus=fs::symlink_status(destination,problem);
    if(problem && problem!=std::errc::no_such_file_or_directory) fail(problem,false);
    problem.clear();
    if(fs::is_symlink(destinationStatus)) throw VbaRuntimeError(3251,"FileCopy symbolic-link destination unsupported");
    if(fs::exists(destinationStatus)) {
        if(!fs::is_regular_file(destinationStatus)) throw VbaRuntimeError(75,"FileCopy destination must be a file");
        if(!overwrite) throw VbaRuntimeError(58,"Copy destination already exists");
        const bool same=fs::equivalent(source,destination,problem);
        if(problem) fail(problem,false);
        if(same) throw VbaRuntimeError(3251,"FileCopy identical-file behavior requires reference validation");
    }
    nativeGuardFileStreams(destination);
    const bool copied=fs::copy_file(source,destination,overwrite ? fs::copy_options::overwrite_existing : fs::copy_options::none,problem);
    if(problem) fail(problem,false);
    if(!copied) throw VbaRuntimeError(75,"Native FileCopy did not copy a file");
}
inline void nativeFileKill(const VbaValue& value,bool force=false) {
    namespace fs=std::filesystem;
    const auto text=coerceVbaModuleValue(value,"string").asText();
    if(text.find('\0')!=std::string::npos) throw VbaRuntimeError(52,"Invalid Kill pathname");
    if(text.find_first_of("*?")!=std::string::npos) throw VbaRuntimeError(3251,"Kill wildcard semantics require reference validation");
    const auto path=nativeFilePath(VbaValue::string(text));
    if(path.empty()) throw VbaRuntimeError(53,"Kill file not found");
    const auto fail=[&](const std::error_code& reason) {
        if(reason==std::errc::permission_denied || reason==std::errc::operation_not_permitted)
            throw VbaRuntimeError(70,"Kill permission denied");
        if(reason==std::errc::invalid_argument || reason==std::errc::filename_too_long)
            throw VbaRuntimeError(52,"Invalid Kill pathname");
        if(reason==std::errc::no_such_file_or_directory || reason==std::errc::not_a_directory) {
            std::error_code parentError;
            const auto parent=path.has_parent_path() ? path.parent_path() : fs::path(".");
            const bool exists=fs::is_directory(parent,parentError);
            if(parentError==std::errc::permission_denied) throw VbaRuntimeError(70,"Kill permission denied");
            throw VbaRuntimeError(exists ? 53 : 76,"Kill file or path not found");
        }
        throw VbaRuntimeError(75,"Native Kill operation failed");
    };
    std::error_code problem;
    const auto status=fs::symlink_status(path,problem);
    if(problem) fail(problem);
    if(fs::is_symlink(status)) throw VbaRuntimeError(3251,"Kill link semantics require reference validation");
    if(!fs::is_regular_file(status)) throw VbaRuntimeError(75,"Kill requires a regular file");
    const auto writable=fs::perms::owner_write | fs::perms::group_write | fs::perms::others_write;
    if(!force && (status.permissions() & writable)==fs::perms::none) throw VbaRuntimeError(70,"Kill read-only file");
    // unlink cannot remove a directory, including one replacing the checked path.
    nativeGuardFileStreams(path);
    if(::unlink(path.c_str())!=0) fail(std::error_code(errno,std::generic_category()));
}
inline void nativeFileRename(const VbaValue& sourceValue,const VbaValue& destinationValue) {
    namespace fs=std::filesystem;
    const auto source=nativeFilePath(sourceValue),destination=nativeFilePath(destinationValue);
    if(source.empty()) throw VbaRuntimeError(53,"Name source not found");
    if(destination.empty()) throw VbaRuntimeError(52,"Invalid Name destination");
    const auto fail=[&](const std::error_code& reason,bool sourceOperation) {
        if(reason==std::errc::file_exists) throw VbaRuntimeError(58,"Name destination already exists");
        if(reason==std::errc::permission_denied || reason==std::errc::operation_not_permitted)
            throw VbaRuntimeError(70,"Name permission denied");
        if(reason==std::errc::filename_too_long) throw VbaRuntimeError(52,"Invalid Name pathname");
        if(reason==std::errc::no_such_file_or_directory || reason==std::errc::not_a_directory) {
            std::error_code parentError;
            const auto parent=source.has_parent_path() ? source.parent_path() : fs::path(".");
            const bool exists=fs::is_directory(parent,parentError);
            if(parentError==std::errc::permission_denied) throw VbaRuntimeError(70,"Name permission denied");
            throw VbaRuntimeError(sourceOperation && exists ? 53 : 76,"Name file or path not found");
        }
        if(reason==std::errc::cross_device_link || reason==std::errc::function_not_supported || reason==std::errc::operation_not_supported || reason==std::errc::invalid_argument)
            throw VbaRuntimeError(3251,"Atomic native Name operation unavailable");
        throw VbaRuntimeError(75,"Native Name operation failed");
    };
    std::error_code problem;
    const auto status=fs::symlink_status(source,problem);
    if(problem) fail(problem,true);
    if(!fs::is_regular_file(status) && !fs::is_directory(status))
        throw VbaRuntimeError(3251,"Name requires a regular file or directory, without terminal symlinks");
    if(fs::is_directory(status)) {
        const auto working=fs::current_path(problem); if(problem) fail(problem,false);
        const auto relative=fs::relative(working,source,problem); if(problem) fail(problem,true);
        if(relative.empty() || *relative.begin()!=fs::path(".."))
            throw VbaRuntimeError(3251,"Renaming an active working-directory ancestor requires reference validation");
    }
    const auto target=fs::symlink_status(destination,problem);
    if(problem && problem!=std::errc::no_such_file_or_directory) fail(problem,false);
    if(fs::exists(target)) throw VbaRuntimeError(58,"Name destination already exists");
#if defined(__linux__) && defined(SYS_renameat2)
    nativeGuardFileStreams(source);
    if(::syscall(SYS_renameat2,AT_FDCWD,source.c_str(),AT_FDCWD,destination.c_str(),RENAME_NOREPLACE)!=0)
        fail(std::error_code(errno,std::generic_category()),false);
#else
    throw VbaRuntimeError(3251,"Atomic native Name operation unavailable");
#endif
}

inline bool nativeFileFlag(const std::vector<VbaValue>& args,std::size_t index,bool fallback) {
    if(index>=args.size() || args[index].kind==VbaValue::Kind::Missing) return fallback;
    return coerceVbaModuleValue(scalarVbaValue(args[index]),"boolean").boolean;
}
inline void nativeRegularFile(const std::filesystem::path& path) {
    std::error_code error;
    const auto status=std::filesystem::symlink_status(path,error);
    if(error==std::errc::permission_denied) throw VbaRuntimeError(70,"File permission denied");
    if(error==std::errc::no_such_file_or_directory || error==std::errc::not_a_directory) throw VbaRuntimeError(53,"File unavailable");
    if(error) throw VbaRuntimeError(75,"File metadata unavailable");
    if(!std::filesystem::is_regular_file(status)) throw VbaRuntimeError(3251,"Operation requires a regular file without terminal links");
}
inline std::filesystem::path nativeFileDestination(const std::filesystem::path& source,const VbaValue& value) {
    const auto text=coerceVbaModuleValue(scalarVbaValue(value),"string").asText();
    auto destination=nativeFilePath(VbaValue::string(text));
    if(destination.empty()) throw VbaRuntimeError(52,"Invalid destination pathname");
    std::error_code error;
    if(text.back()=='/') {
        destination=destination.lexically_normal();
        if(destination.filename().empty() && destination!=destination.root_path()) destination=destination.parent_path();
        const auto status=std::filesystem::symlink_status(destination,error);
        if(error==std::errc::permission_denied) throw VbaRuntimeError(70,"Destination permission denied");
        if(!error && std::filesystem::is_symlink(status)) throw VbaRuntimeError(3251,"Destination folder link semantics require reference validation");
        if(error || !std::filesystem::is_directory(status)) throw VbaRuntimeError(76,"Destination folder unavailable");
        destination/=source.filename();
    }
    destination=std::filesystem::absolute(destination,error);
    if(error) throw VbaRuntimeError(75,"Destination path unavailable");
    return destination.lexically_normal();
}
inline void nativeCopyDestinationWritable(const std::filesystem::path& path) {
    std::error_code error;
    const auto status=std::filesystem::symlink_status(path,error);
    if(error==std::errc::no_such_file_or_directory || error==std::errc::not_a_directory) return;
    if(error==std::errc::permission_denied) throw VbaRuntimeError(70,"Copy permission denied");
    if(error) throw VbaRuntimeError(75,"Copy destination unavailable");
    const auto writable=std::filesystem::perms::owner_write | std::filesystem::perms::group_write | std::filesystem::perms::others_write;
    if(std::filesystem::is_regular_file(status) && (status.permissions() & writable)==std::filesystem::perms::none)
        throw VbaRuntimeError(70,"Copy destination is read-only");
}
}
