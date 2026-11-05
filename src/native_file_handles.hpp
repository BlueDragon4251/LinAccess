#pragma once
#include "vba_expression.hpp"
#include <algorithm>
#include <filesystem>
#include <memory>
#include <mutex>
#include <vector>
#include <sys/stat.h>
namespace linaccess {
struct NativeFileStreamToken { dev_t device; ino_t inode; bool writing; };
inline std::mutex nativeFileStreamMutex;
inline std::vector<std::weak_ptr<NativeFileStreamToken>> nativeFileStreams;
inline std::shared_ptr<NativeFileStreamToken> nativeRegisterFileStream(const struct stat& status,bool writing=false) {
    auto token=std::make_shared<NativeFileStreamToken>(NativeFileStreamToken{status.st_dev,status.st_ino,writing});
    std::lock_guard lock(nativeFileStreamMutex);
    std::erase_if(nativeFileStreams,[](const auto& entry) { return entry.expired(); });
    for(const auto& entry:nativeFileStreams) if(const auto existing=entry.lock();existing && existing->device==status.st_dev && existing->inode==status.st_ino && (writing || existing->writing))
        throw VbaRuntimeError(3251,"Concurrent native text writing requires sharing validation");
    nativeFileStreams.push_back(token);
    return token;
}
inline void nativeGuardFileStreams(const std::filesystem::path& path) {
    struct stat status{};
    if(::stat(path.c_str(),&status)!=0) return;
    std::lock_guard lock(nativeFileStreamMutex);
    for(const auto& entry:nativeFileStreams) if(const auto token=entry.lock();token && token->device==status.st_dev && token->inode==status.st_ino)
        throw VbaRuntimeError(3251,"Native file mutation while a text stream is open requires sharing validation");
}
}
