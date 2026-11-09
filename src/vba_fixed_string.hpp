#pragma once
#include "vba_runtime.hpp"
#include <charconv>
#include <cstddef>
namespace linaccess {
// Internal canonical declaration, distinct from the public String Variant type.
inline std::optional<std::size_t> vbaFixedStringWidth(std::string_view type) {
    if(!type.starts_with("string*")) return {};
    const auto digits=type.substr(7);unsigned width=0;
    const auto converted=std::from_chars(digits.data(),digits.data()+digits.size(),width);
    if(converted.ec!=std::errc{} || converted.ptr!=digits.data()+digits.size() || width<1 || width>65526)
        throw VbaRuntimeError(5,"Fixed String length must be1..65526");
    return width;
}
}
