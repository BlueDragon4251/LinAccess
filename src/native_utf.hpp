#pragma once
#include "vba_expression.hpp"
#include <cstdint>
#include <string>
#include <string_view>
namespace linaccess {
inline std::u16string nativeUtf16(std::string_view input) {
    std::u16string output;
    for (std::size_t i=0;i<input.size();) {
        const auto first=static_cast<unsigned char>(input[i++]);
        std::uint32_t point=first,minimum=0; unsigned continuation=0;
        if (first<128) {}
        else if (first>=0xc2 && first<=0xdf) { point=first&31; continuation=1; minimum=0x80; }
        else if (first>=0xe0 && first<=0xef) { point=first&15; continuation=2; minimum=0x800; }
        else if (first>=0xf0 && first<=0xf4) { point=first&7; continuation=3; minimum=0x10000; }
        else throw VbaRuntimeError(13,"Invalid UTF-8 runtime string");
        if (continuation>input.size()-i) throw VbaRuntimeError(13,"Truncated UTF-8 runtime string");
        for (unsigned n=0;n<continuation;++n) {
            const auto byte=static_cast<unsigned char>(input[i++]);
            if ((byte&0xc0)!=0x80) throw VbaRuntimeError(13,"Invalid UTF-8 continuation");
            point=(point<<6)|(byte&63);
        }
        if (point<minimum || point>0x10ffff || (point>=0xd800 && point<=0xdfff)) throw VbaRuntimeError(13,"Invalid Unicode scalar");
        if (point<=0xffff) output.push_back(static_cast<char16_t>(point));
        else { point-=0x10000; output.push_back(static_cast<char16_t>(0xd800+(point>>10))); output.push_back(static_cast<char16_t>(0xdc00+(point&1023))); }
    }
    return output;
}
inline std::string nativeUtf8(std::u16string_view input) {
    std::string output;
    for (std::size_t i=0;i<input.size();++i) {
        std::uint32_t point=input[i];
        if (point>=0xd800 && point<=0xdbff) {
            if (i+1>=input.size() || input[i+1]<0xdc00 || input[i+1]>0xdfff) throw VbaRuntimeError(3251,"Isolated UTF-16 surrogate unsupported");
            point=0x10000+((point-0xd800)<<10)+(input[++i]-0xdc00);
        } else if (point>=0xdc00 && point<=0xdfff) throw VbaRuntimeError(3251,"Isolated UTF-16 surrogate unsupported");
        if (point<0x80) output.push_back(static_cast<char>(point));
        else if (point<0x800) { output.push_back(static_cast<char>(0xc0|(point>>6))); output.push_back(static_cast<char>(0x80|(point&63))); }
        else if (point<0x10000) { output.push_back(static_cast<char>(0xe0|(point>>12))); output.push_back(static_cast<char>(0x80|((point>>6)&63))); output.push_back(static_cast<char>(0x80|(point&63))); }
        else { output.push_back(static_cast<char>(0xf0|(point>>18))); output.push_back(static_cast<char>(0x80|((point>>12)&63))); output.push_back(static_cast<char>(0x80|((point>>6)&63))); output.push_back(static_cast<char>(0x80|(point&63))); }
    }
    return output;
}

}
