#pragma once
#include <functional>
#include <string_view>
namespace linaccess {
// Shares the native VBA Like parser/matcher; caller supplies UTF-16 collation.
// Text ranges/surrogates keep the same explicit native restrictions.
bool matchVbaLikePattern(std::u16string_view input,std::u16string_view pattern,bool textMode,
    const std::function<int(char16_t,char16_t)>& compare);
}
