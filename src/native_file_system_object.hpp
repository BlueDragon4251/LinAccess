#pragma once
#include "vba_expression.hpp"
#include "native_file_time.hpp"
#include <algorithm>
#include <cctype>
#include <filesystem>

namespace linaccess {
// Read-only native file object. Metadata is read again rather than cached.
class VbaFileObject final : public VbaObject {
    std::filesystem::path path_;
public:
    explicit VbaFileObject(std::filesystem::path path) : path_(std::move(path)) {}
    std::string_view className() const override { return "File"; }
    bool isNativeType(std::string_view type) const override { return type=="scripting.file"; }
    bool hasScalarDefault() const override { return true; }
    VbaValue scalarValue() override { return VbaValue::string(path_.string()); }
    VbaValue defaultMember(const std::vector<VbaValue>& args,bool reference) override {
        if(!args.empty()) throw VbaRuntimeError(450,"File default property takes no arguments");
        if(reference) throw VbaRuntimeError(13,"File default property is scalar");
        return scalarValue();
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        std::string method(name);
        std::transform(method.begin(),method.end(),method.begin(),[](unsigned char ch) { return std::tolower(ch); });
        if(method!="path" && method!="name" && method!="size" && method!="datelastmodified") throw VbaRuntimeError(3251,"Native file member is not implemented");
        if(!args.empty()) throw VbaRuntimeError(450,"File property takes no arguments");
        if(method=="path") return VbaValue::string(path_.string());
        if(method=="name") return VbaValue::string(path_.filename().string());
        std::error_code error;
        const auto status=std::filesystem::symlink_status(path_,error);
        if(error==std::errc::no_such_file_or_directory || error==std::errc::not_a_directory) throw VbaRuntimeError(53,"File unavailable");
        if(error==std::errc::permission_denied) throw VbaRuntimeError(70,"Filesystem permission denied");
        if(error) throw VbaRuntimeError(75,"Native file metadata unavailable");
        if(!std::filesystem::is_regular_file(status)) throw VbaRuntimeError(3251,"File metadata requires a regular file");
        if(method=="datelastmodified") {
            const auto modified=std::filesystem::last_write_time(path_,error);
            if(error==std::errc::permission_denied) throw VbaRuntimeError(70,"Filesystem permission denied");
            if(error) throw VbaRuntimeError(75,"Native file timestamp unavailable");
            return nativeFileDate(modified);
        }
        const auto size=std::filesystem::file_size(path_,error);
        if(error==std::errc::permission_denied) throw VbaRuntimeError(70,"Filesystem permission denied");
        if(error) throw VbaRuntimeError(75,"Native file size unavailable");
        if(size>9007199254740991ULL) throw VbaRuntimeError(6,"File size exceeds exact numeric range");
        if(size<=2147483647ULL) return VbaValue::longInteger(static_cast<std::int32_t>(size));
        return VbaValue::numeric(static_cast<double>(size));
    }
    void assign(std::string_view,const VbaValue&) override { throw VbaRuntimeError(3251,"Native file property mutation is not implemented"); }
};
// Stateless local filesystem adapter. No COM activation, host or database handle.
class VbaFileSystemObject final : public VbaObject {
    static std::string key(std::string_view name) {
        std::string result(name);
        std::transform(result.begin(),result.end(),result.begin(),[](unsigned char ch) { return std::tolower(ch); });
        return result;
    }
    static std::string pathText(const VbaValue& value) {
        auto text=coerceVbaModuleValue(scalarVbaValue(value),"string").asText();
        if(text.find('\0')!=std::string::npos) throw VbaRuntimeError(52,"Invalid filesystem pathname");
        return text;
    }
    static std::filesystem::path path(const VbaValue& value,bool pattern=false) {
        const auto text=pathText(value);
        if(text.find('\\')!=std::string::npos || text.starts_with("//") || (text.size()>1 && text[1]==':'))
            throw VbaRuntimeError(3251,"Windows filesystem path mapping is unavailable");
        if(!pattern && text.find_first_of("*?")!=std::string::npos)
            throw VbaRuntimeError(3251,"Filesystem wildcard existence semantics require reference validation");
        return std::filesystem::path(text);
    }
    static VbaValue pathPart(const std::string& method,const VbaValue& value) {
        const auto text=pathText(value);
        const auto separator=[](char ch) { return ch=='/' || ch=='\\'; };
        const bool drive=text.size()>=2 && std::isalpha(static_cast<unsigned char>(text[0])) && text[1]==':';
        if(text.empty() || (drive && text.size()==2)) return VbaValue::string("");
        if(separator(text.back())) throw VbaRuntimeError(3251,"Trailing-separator path component semantics require reference validation");
        if(text.size()>1 && separator(text[0]) && separator(text[1])) {
            int components=0; bool inside=false;
            for(const auto ch:text) { if(separator(ch)) inside=false; else if(!inside) { ++components; inside=true; } }
            if(components<3) throw VbaRuntimeError(3251,"UNC-root component semantics require reference validation");
        }
        const auto split=text.find_last_of("/\\");
        const auto leaf=text.substr(split==std::string::npos ? (drive ? 2 : 0) : split+1);
        if(leaf=="." || leaf=="..") throw VbaRuntimeError(3251,"Dot-directory component semantics require reference validation");
        if(method=="getfilename") return VbaValue::string(leaf);
        if(method=="getparentfoldername") {
            if(drive && split==std::string::npos) throw VbaRuntimeError(3251,"Drive-relative parent semantics require reference validation");
            if(split==std::string::npos) return VbaValue::string("");
            const auto length=(split==0 || (drive && split==2)) ? split+1 : split;
            return VbaValue::string(text.substr(0,length));
        }
        if(!leaf.empty() && leaf.front()=='.') throw VbaRuntimeError(3251,"Leading-dot extension semantics require reference validation");
        const auto dot=leaf.find_last_of('.');
        if(method=="getbasename") return VbaValue::string(dot==std::string::npos ? leaf : leaf.substr(0,dot));
        return VbaValue::string(dot==std::string::npos ? "" : leaf.substr(dot+1));
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
        if(method=="getfile") return VbaArgumentSignature{{"FileSpec"},1};
        if(method=="fileexists") return VbaArgumentSignature{{"FileSpec"},1};
        if(method=="folderexists") return VbaArgumentSignature{{"FolderSpec"},1};
        if(method=="buildpath") return VbaArgumentSignature{{"Path","Name"},2};
        if(method=="getabsolutepathname") return VbaArgumentSignature{{"PathSpec"},1};
        if(method=="getfilename") return VbaArgumentSignature{{"PathSpec"},1};
        if(method=="getbasename" || method=="getextensionname" || method=="getparentfoldername") return VbaArgumentSignature{{"Path"},1};
        return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto method=key(name);
        const bool lexical=method=="getfilename" || method=="getbasename" || method=="getextensionname" || method=="getparentfoldername";
        if(!lexical && method!="getfile" && method!="fileexists" && method!="folderexists" && method!="buildpath" && method!="getabsolutepathname")
            throw VbaRuntimeError(3251,"Native FileSystemObject member is not implemented");
        if(args.size()!=(method=="buildpath" ? 2u : 1u) || std::any_of(args.begin(),args.end(),[](const auto& value) { return value.kind==VbaValue::Kind::Missing; }))
            throw VbaRuntimeError(450,"Invalid FileSystemObject argument count");
        if(lexical) return pathPart(method,args[0]);
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
        if(method=="getfile") {
            if(source.empty()) throw VbaRuntimeError(53,"File unavailable");
            auto absolute=std::filesystem::absolute(source,problem);
            if(problem) failure(problem);
            absolute=absolute.lexically_normal();
            const auto status=std::filesystem::symlink_status(absolute,problem);
            if(problem==std::errc::no_such_file_or_directory || problem==std::errc::not_a_directory) throw VbaRuntimeError(53,"File unavailable");
            if(problem) failure(problem);
            if(std::filesystem::is_symlink(status)) throw VbaRuntimeError(3251,"GetFile link semantics require reference validation");
            if(!std::filesystem::is_regular_file(status)) throw VbaRuntimeError(53,"File unavailable");
            return VbaValue::reference(std::make_shared<VbaFileObject>(std::move(absolute)));
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
