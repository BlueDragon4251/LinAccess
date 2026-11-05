#pragma once
#include "native_file_operations.hpp"
#include "native_file_handles.hpp"
#include <cstdio>
#include <limits>
namespace linaccess {
// A real owned text handle. No host or database reference is retained.
class VbaTextStream final : public VbaObject {
    std::FILE* file_{};
    std::shared_ptr<NativeFileStreamToken> token_;
    std::int32_t line_{1},column_{1};
    bool previousCr_{};
    int mode_{1};
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
    VbaTextStream(std::FILE* file,std::shared_ptr<NativeFileStreamToken> token,int mode=1) : file_(file),token_(std::move(token)),mode_(mode) {}
    ~VbaTextStream() override { if(file_) std::fclose(file_); }
    std::string_view className() const override { return "TextStream"; }
    bool isNativeType(std::string_view type) const override { return type=="scripting.textstream"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto method=key(name);
        if(method=="read") return VbaArgumentSignature{{"Characters"},1};
        if(method=="write") return VbaArgumentSignature{{"Text"},1};
        if(method=="writeline") return VbaArgumentSignature{{"Text"},0};
        if(method=="writeblanklines") return VbaArgumentSignature{{"Lines"},1};
        return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto method=key(name);
        if(method!="close" && method!="read" && method!="readline" && method!="readall" && method!="atendofstream" && method!="line" && method!="column" && method!="write" && method!="writeline" && method!="writeblanklines")
            throw VbaRuntimeError(3251,"Native TextStream member is not implemented");
        const bool output=method=="write" || method=="writeline" || method=="writeblanklines";
        if(output) {
            if(args.size()>1 || (method!="writeline" && (args.empty() || args[0].kind==VbaValue::Kind::Missing))) throw VbaRuntimeError(450,"Invalid TextStream write argument count");
        } else if(args.size()!=(method=="read" ? 1u : 0u)) throw VbaRuntimeError(450,"Invalid TextStream argument count");
        if(!file_) throw VbaRuntimeError(3251,"Closed TextStream error contract requires reference validation");
        if(method=="close") {
            auto* closing=file_; file_=nullptr;
            const auto result=std::fclose(closing); token_.reset();
            if(result!=0) throw VbaRuntimeError(75,"TextStream close failed");
            return {};
        }
        if(output) {
            if(mode_==1) throw VbaRuntimeError(3251,"TextStream is not open for writing");
            std::string text;
            if(method=="writeblanklines") {
                const auto count=coerceVbaModuleValue(scalarVbaValue(args[0]),"long").number;
                if(count<=0 || count>1000000) throw VbaRuntimeError(3251,"WriteBlankLines count outside supported range");
                text.reserve(static_cast<std::size_t>(count)*2);
                for(std::size_t i=0;i<static_cast<std::size_t>(count);++i) text+="\r\n";
            } else {
                if(!args.empty() && args[0].kind!=VbaValue::Kind::Missing) text=coerceVbaModuleValue(scalarVbaValue(args[0]),"string").asText();
                if(text.size()>16u*1024u*1024u) throw VbaRuntimeError(3251,"Native text write limit exceeded");
                if(std::any_of(text.begin(),text.end(),[](unsigned char ch) { return ch>127; })) throw VbaRuntimeError(3251,"Non-ASCII output requires a validated encoding adapter");
                if(method=="writeline") text+="\r\n";
            }
            if(std::fwrite(text.data(),1,text.size(),file_)!=text.size() || std::fflush(file_)!=0) throw VbaRuntimeError(75,"Native text write failed");
            return {};
        }
        if(mode_!=1) throw VbaRuntimeError(3251,"TextStream input properties require reading mode");
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
struct NativeTextFileCloser {
    void operator()(std::FILE* file) const noexcept { std::fclose(file); }
};
struct NativeTextDescriptor {
    int value{-1};
    ~NativeTextDescriptor() { if(value>=0) ::close(value); }
};
inline VbaValue nativeOpenTextStream(const VbaValue& input,int mode=1,bool create=false,bool exclusive=false) {
    if(mode!=1 && mode!=2 && mode!=8) throw VbaRuntimeError(3251,"Unsupported native text mode");
    const auto path=nativeFilePath(scalarVbaValue(input));
    if(path.empty()) throw VbaRuntimeError(53,"Text file unavailable");
    int flags=O_NONBLOCK|O_NOFOLLOW|O_CLOEXEC|(mode==1 ? O_RDONLY : O_WRONLY);
    if(create) flags|=O_CREAT;
    if(exclusive) flags|=O_EXCL;
    if(mode==8) flags|=O_APPEND;
    NativeTextDescriptor descriptor{::open(path.c_str(),flags,0666)};
    if(descriptor.value<0) {
        if(errno==EEXIST) throw VbaRuntimeError(58,"Text file already exists");
        if(errno==EACCES || errno==EPERM) throw VbaRuntimeError(70,"Text file permission denied");
        if(errno==ENOENT || errno==ENOTDIR) throw VbaRuntimeError(53,"Text file unavailable");
        if(errno==ELOOP) throw VbaRuntimeError(3251,"Text stream link semantics require reference validation");
        throw VbaRuntimeError(75,"Native text file open failed");
    }
    struct stat status{};
    if(::fstat(descriptor.value,&status)!=0 || !S_ISREG(status.st_mode)) throw VbaRuntimeError(3251,"TextStream requires a regular native file");
    if(mode!=1 && !(status.st_mode & (S_IWUSR|S_IWGRP|S_IWOTH))) throw VbaRuntimeError(70,"Text file is read-only");
    auto token=nativeRegisterFileStream(status,mode!=1);
    // Register/check sharing before truncation, so refusal preserves existing bytes.
    if(mode==2 && ::ftruncate(descriptor.value,0)!=0) throw VbaRuntimeError(75,"Native text truncation failed");
    auto* file=::fdopen(descriptor.value,mode==1 ? "rb" : mode==8 ? "ab" : "wb");
    if(!file) throw VbaRuntimeError(75,"Native text stream open failed");
    descriptor.value=-1;
    std::unique_ptr<std::FILE,NativeTextFileCloser> owned(file);
    auto stream=std::make_shared<VbaTextStream>(file,std::move(token),mode);
    owned.release();
    return VbaValue::reference(std::move(stream));
}
inline int nativeTextOptions(const std::vector<VbaValue>& args,std::size_t modeIndex,std::size_t formatIndex) {
    const auto number=[&](std::size_t index,double fallback) { return index>=args.size() || args[index].kind==VbaValue::Kind::Missing ? fallback : coerceVbaModuleValue(scalarVbaValue(args[index]),"long").number; };
    const auto mode=number(modeIndex,1);
    if((mode!=1 && mode!=2 && mode!=8) || number(formatIndex,0)!=0) throw VbaRuntimeError(3251,"Only native ASCII text modes are currently implemented");
    return static_cast<int>(mode);
}
inline VbaValue nativeCreateTextStream(const VbaValue& path,const std::vector<VbaValue>& args) {
    const bool overwrite=nativeFileFlag(args,1,true);
    if(nativeFileFlag(args,2,false)) throw VbaRuntimeError(3251,"Unicode text output requires a validated encoding adapter");
    return nativeOpenTextStream(path,2,true,!overwrite);
}
}
