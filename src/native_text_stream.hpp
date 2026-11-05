#pragma once
#include "native_file_operations.hpp"
#include "native_file_handles.hpp"
#include <cstdio>
#include <limits>
namespace linaccess {
// A real owned read handle. No host or database reference is retained.
class VbaTextStream final : public VbaObject {
    std::FILE* file_{};
    std::shared_ptr<NativeFileReadToken> token_;
    std::int32_t line_{1},column_{1};
    bool previousCr_{};
    static std::string key(std::string_view name) {
        std::string result(name);
        std::transform(result.begin(),result.end(),result.begin(),[](unsigned char ch) { return std::tolower(ch); });
        return result;
    }
    int peek() {
        const auto ch=std::fgetc(file_);
        if(ch==EOF) { if(std::ferror(file_)) throw VbaRuntimeError(75,"Native text read failed"); return EOF; }
        if(std::ungetc(ch,file_)==EOF) throw VbaRuntimeError(75,"Native text lookahead failed");
        return ch;
    }
    int take() {
        const auto ch=std::fgetc(file_);
        if(ch==EOF) { if(std::ferror(file_)) throw VbaRuntimeError(75,"Native text read failed"); return EOF; }
        if(ch>127) throw VbaRuntimeError(3251,"Non-ASCII text requires an explicit validated encoding adapter");
        if(line_==std::numeric_limits<std::int32_t>::max() || column_==std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"Text stream position overflow");
        if(ch=='\r') { ++line_; column_=1; previousCr_=true; }
        else if(ch=='\n') { if(!previousCr_) ++line_; column_=1; previousCr_=false; }
        else { ++column_; previousCr_=false; }
        return ch;
    }
public:
    VbaTextStream(std::FILE* file,std::shared_ptr<NativeFileReadToken> token) : file_(file),token_(std::move(token)) {}
    ~VbaTextStream() override { if(file_) std::fclose(file_); }
    std::string_view className() const override { return "TextStream"; }
    bool isNativeType(std::string_view type) const override { return type=="scripting.textstream"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(key(name)=="read") return VbaArgumentSignature{{"Characters"},1};
        return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto method=key(name);
        if(method!="close" && method!="read" && method!="readline" && method!="readall" && method!="atendofstream" && method!="line" && method!="column")
            throw VbaRuntimeError(3251,"Native TextStream member is not implemented");
        if(args.size()!=(method=="read" ? 1u : 0u)) throw VbaRuntimeError(450,"Invalid TextStream argument count");
        if(!file_) throw VbaRuntimeError(3251,"Closed TextStream error contract requires reference validation");
        if(method=="close") {
            auto* closing=file_; file_=nullptr;
            const auto result=std::fclose(closing); token_.reset();
            if(result!=0) throw VbaRuntimeError(75,"TextStream close failed");
            return {};
        }
        if(method=="atendofstream") return VbaValue::logical(peek()==EOF);
        if(method=="line") return VbaValue::longInteger(line_);
        if(method=="column") return VbaValue::longInteger(column_);
        std::size_t count=16u*1024u*1024u;
        if(method=="read") {
            const auto value=coerceVbaModuleValue(scalarVbaValue(args[0]),"long");
            if(value.number<=0 || value.number>1000000) throw VbaRuntimeError(3251,"TextStream Read count outside supported range");
            count=static_cast<std::size_t>(value.number);
            struct stat status{}; const auto position=::ftello(file_);
            if(position<0 || ::fstat(::fileno(file_),&status)!=0) throw VbaRuntimeError(75,"Text stream extent unavailable");
            if(status.st_size<position || static_cast<std::uintmax_t>(status.st_size-position)<count)
                throw VbaRuntimeError(3251,"Read beyond remaining characters requires reference validation");
        }
        const auto position=::ftello(file_); const auto savedLine=line_,savedColumn=column_; const auto savedCr=previousCr_;
        if(position<0) throw VbaRuntimeError(75,"Text stream position unavailable");
        std::string result;
        try {
            if(method=="readline") {
                if(peek()==EOF) throw VbaRuntimeError(62,"Input past end of text stream");
                for(;;) {
                    const auto ch=take();
                    if(ch==EOF) break;
                    if(ch=='\r') { if(peek()=='\n') (void)take(); break; }
                    if(ch=='\n') break;
                    if(result.size()>=1000000) throw VbaRuntimeError(3251,"Native text line limit exceeded");
                    result.push_back(static_cast<char>(ch));
                }
            } else {
                while(result.size()<count) { const auto ch=take(); if(ch==EOF) break; result.push_back(static_cast<char>(ch)); }
                if(method=="readall" && result.size()==count && peek()!=EOF) throw VbaRuntimeError(3251,"Native ReadAll size limit exceeded");
            }
        } catch(...) {
            const bool restored=::fseeko(file_,position,SEEK_SET)==0;
            line_=savedLine; column_=savedColumn; previousCr_=savedCr;
            if(!restored) throw VbaRuntimeError(75,"Text stream position recovery failed");
            throw;
        }
        return VbaValue::string(std::move(result));
    }
    void assign(std::string_view,const VbaValue&) override { throw VbaRuntimeError(438,"TextStream property is not writable"); }
};
inline VbaValue nativeOpenTextStream(const VbaValue& input) {
    const auto path=nativeFilePath(scalarVbaValue(input));
    if(path.empty()) throw VbaRuntimeError(53,"Text file unavailable");
    const auto descriptor=::open(path.c_str(),O_RDONLY|O_NONBLOCK|O_NOFOLLOW|O_CLOEXEC);
    if(descriptor<0) {
        if(errno==EACCES || errno==EPERM) throw VbaRuntimeError(70,"Text file permission denied");
        if(errno==ENOENT || errno==ENOTDIR) throw VbaRuntimeError(53,"Text file unavailable");
        if(errno==ELOOP) throw VbaRuntimeError(3251,"Text stream link semantics require reference validation");
        throw VbaRuntimeError(75,"Native text file open failed");
    }
    struct stat status{};
    if(::fstat(descriptor,&status)!=0 || !S_ISREG(status.st_mode)) { ::close(descriptor); throw VbaRuntimeError(3251,"TextStream requires a regular native file"); }
    auto* file=::fdopen(descriptor,"rb");
    if(!file) { ::close(descriptor); throw VbaRuntimeError(75,"Native text stream open failed"); }
    try { return VbaValue::reference(std::make_shared<VbaTextStream>(file,nativeRegisterFileRead(status))); }
    catch(...) { std::fclose(file); throw; }
}
inline void nativeTextReadOptions(const std::vector<VbaValue>& args,std::size_t modeIndex,std::size_t formatIndex) {
    const auto number=[&](std::size_t index,double fallback) { return index>=args.size() || args[index].kind==VbaValue::Kind::Missing ? fallback : coerceVbaModuleValue(scalarVbaValue(args[index]),"long").number; };
    if(number(modeIndex,1)!=1 || number(formatIndex,0)!=0) throw VbaRuntimeError(3251,"Only native ASCII text reading is currently implemented");
}
}
