#pragma once
#include "vba_expression.hpp"
#include <algorithm>
#include <cctype>
#include <filesystem>

namespace linaccess {
// Stateless local filesystem adapter. No COM activation, host or database handle.
class VbaFileSystemObject final : public VbaObject {
    static std::string key(std::string_view name) {
        std::string result(name);
        std::transform(result.begin(),result.end(),result.begin(),[](unsigned char ch) { return std::tolower(ch); });
        return result;
    }
    static std::filesystem::path path(const VbaValue& value,bool pattern=false) {
        const auto text=coerceVbaModuleValue(scalarVbaValue(value),"string").asText();
        if(text.find('\0')!=std::string::npos) throw VbaRuntimeError(52,"Invalid filesystem pathname");
        if(text.find('\\')!=std::string::npos || text.starts_with("//") || (text.size()>1 && text[1]==':'))
            throw VbaRuntimeError(3251,"Windows filesystem path mapping is unavailable");
        if(!pattern && text.find_first_of("*?")!=std::string::npos)
            throw VbaRuntimeError(3251,"Filesystem wildcard existence semantics require reference validation");
        return std::filesystem::path(text);
    }
    static void failure(const std::error_code& reason) {
        if(reason==std::errc::permission_denied || reason==std::errc::operation_not_permitted)
            throw VbaRuntimeError(70,"Filesystem permission denied");
        if(reason==std::errc::filename_too_long || reason==std::errc::invalid_argument)
            throw VbaRuntimeError(52,"Invalid filesystem pathname");
        if(reason==std::errc::no_such_file_or_directory || reason==std::errc::not_a_directory)
            throw VbaRuntimeError(76,"Filesystem path unavailable");
        throw VbaRuntimeError(75,"Native filesystem operation failed");
    }
public:
    std::string_view className() const override { return "FileSystemObject"; }
    bool isNativeType(std::string_view type) const override { return type=="scripting.filesystemobject"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto method=key(name);
        if(method=="fileexists") return VbaArgumentSignature{{"FileSpec"},1};
        if(method=="folderexists") return VbaArgumentSignature{{"FolderSpec"},1};
        if(method=="buildpath") return VbaArgumentSignature{{"Path","Name"},2};
        if(method=="getabsolutepathname") return VbaArgumentSignature{{"PathSpec"},1};
        return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto method=key(name);
        if(method!="fileexists" && method!="folderexists" && method!="buildpath" && method!="getabsolutepathname")
            throw VbaRuntimeError(3251,"Native FileSystemObject member is not implemented");
        if(args.size()!=(method=="buildpath" ? 2u : 1u) || std::any_of(args.begin(),args.end(),[](const auto& value) { return value.kind==VbaValue::Kind::Missing; }))
            throw VbaRuntimeError(450,"Invalid FileSystemObject argument count");
        const auto source=path(args[0],method=="buildpath");
        if(method=="buildpath") {
            const auto name=path(args[1],true);
            if(name.has_root_path()) throw VbaRuntimeError(3251,"Rooted BuildPath names require reference validation");
            return VbaValue::string((source/name).string());
        }
        std::error_code problem;
        if(method=="getabsolutepathname") {
            auto result=(source.empty() ? std::filesystem::current_path(problem) : std::filesystem::absolute(source,problem));
            if(problem) failure(problem);
            result=result.lexically_normal();
            if(result.filename().empty() && result!=result.root_path()) result=result.parent_path();
            return VbaValue::string(result.string());
        }
        if(source.empty()) return VbaValue::logical(false);
        const auto status=std::filesystem::status(source,problem);
        if(problem==std::errc::no_such_file_or_directory || problem==std::errc::not_a_directory) return VbaValue::logical(false);
        if(problem) failure(problem);
        return VbaValue::logical(method=="fileexists" ? std::filesystem::is_regular_file(status) : std::filesystem::is_directory(status));
    }
    void assign(std::string_view,const VbaValue&) override { throw VbaRuntimeError(438,"FileSystemObject property is not writable"); }
};
}
