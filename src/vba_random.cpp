#include "vba_random.hpp"
#include "vba_expression.hpp"
#include <bit>
#include <chrono>
#include <ctime>
namespace linaccess {
// VBMath algorithm reference and MIT notice: third_party/visualbasic-random.
VbaValue nextVbaRandom(VbaHost& host,const VbaValue& argument) {
    const auto value=scalarVbaValue(argument);
    const float number=value.kind==VbaValue::Kind::Missing ? 1.0f
        : static_cast<float>(coerceVbaModuleValue(value,"single").asNumber());
    const auto state=host.randomState();if(!state) throw VbaRuntimeError(3251,"VBA random state unavailable");
    if(number!=0.0f) {
        if(number<0.0f) {const auto bits=std::bit_cast<std::uint32_t>(number);state->seed=(bits+(bits>>24))&0xffffffu;}
        state->seed=static_cast<std::uint32_t>((static_cast<std::uint64_t>(state->seed)*0x43fd43fdu+0xc39ec3u)&0xffffffu);
    }
    return VbaValue::single(static_cast<float>(state->seed)/16777216.0f);
}
void randomizeVbaGenerator(VbaHost& host,const VbaValue& argument) {
    const auto value=scalarVbaValue(argument);
    std::uint32_t bits;
    if(value.kind==VbaValue::Kind::Missing || value.kind==VbaValue::Kind::Empty) {
        const auto sampled=std::chrono::system_clock::now();const auto seconds=std::chrono::floor<std::chrono::seconds>(sampled);
        const auto epoch=std::chrono::system_clock::to_time_t(seconds);std::tm local{};
        if(!::localtime_r(&epoch,&local)) throw VbaRuntimeError(3251,"VBA random timer unavailable");
        const float timer=static_cast<float>(local.tm_hour*3600+local.tm_min*60+local.tm_sec+std::chrono::duration<double>(sampled-seconds).count());
        bits=std::bit_cast<std::uint32_t>(timer);
    } else {
        const double number=coerceVbaModuleValue(value,"double").asNumber();bits=static_cast<std::uint32_t>(std::bit_cast<std::uint64_t>(number)>>32);
    }
    const auto mixed=((bits&0xffffu)^(bits>>16))&0xffffu;
    const auto state=host.randomState();if(!state) throw VbaRuntimeError(3251,"VBA random state unavailable");
    state->seed=(state->seed&0xffu)|(mixed<<8);
}
}
