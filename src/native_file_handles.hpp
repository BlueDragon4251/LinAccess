#pragma once
#include "vba_expression.hpp"
#include <algorithm>
#include <filesystem>
#include <memory>
#include <mutex>
#include <vector>
#include <sys/stat.h>
namespace linaccess {
struct NativeFileReadToken { dev_t device; ino_t inode; };
inline std::mutex nativeFileReadMutex;
inline std::vector<std::weak_ptr<NativeFileReadToken>> nativeFileReads;
inline std::shared_ptr<NativeFileReadToken> nativeRegisterFileRead(const struct stat& status) {
    auto token=std::make_shared<NativeFileReadToken>(NativeFileReadToken{status.st_dev,status.st_ino});
    std::lock_guard lock(nativeFileReadMutex);
    std::erase_if(nativeFileReads,[](const auto& entry) { return entry.expired(); });
    nativeFileReads.push_back(token);
    return token;
}
inline void nativeGuardFileRead(const std::filesystem::path& path) {
    struct stat status{};
    if(::stat(path.c_str(),&status)!=0) return;
    std::lock_guard lock(nativeFileReadMutex);
    for(const auto& entry:nativeFileReads) if(const auto token=entry.lock();token && token->device==status.st_dev && token->inode==status.st_ino)
        throw VbaRuntimeError(3251,"Native file mutation while a text stream is open requires sharing validation");
}
}
