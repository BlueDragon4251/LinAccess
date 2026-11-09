#include "vba_debug_output.hpp"
#include "native_utf.hpp"
#include <algorithm>
#include <cctype>
#include <deque>
#include <limits>
namespace linaccess {
namespace {
constexpr std::size_t byteLimit=1024*1024,lineLimit=8192;
struct Chunk {std::string text;std::size_t lines{};};
struct Output {std::deque<Chunk> chunks;std::size_t bytes{},lines{},column{1};std::uint64_t revision{};bool truncated{};};
thread_local Output output;
std::string trim(std::string_view input) {while(!input.empty() && std::isspace(static_cast<unsigned char>(input.front())))input.remove_prefix(1);while(!input.empty() && std::isspace(static_cast<unsigned char>(input.back())))input.remove_suffix(1);return std::string(input);}
std::string lower(std::string text){for(auto& c:text)c=static_cast<char>(std::tolower(static_cast<unsigned char>(c)));return text;}
std::size_t lineBreak(std::string_view text,std::size_t i){if(text[i]=='\r')return i+1<text.size() && text[i+1]=='\n' ? 2 : 1;if(text[i]=='\n')return 1;if(i+2<text.size() && static_cast<unsigned char>(text[i])==0xe2 && static_cast<unsigned char>(text[i+1])==0x80 && (static_cast<unsigned char>(text[i+2])==0xa8 || static_cast<unsigned char>(text[i+2])==0xa9))return 3;return 0;}
std::size_t lines(std::string_view text){std::size_t count=0;for(std::size_t i=0;i<text.size();){const auto length=lineBreak(text,i);if(length)++count;i+=length ? length : 1;}return count;}
}
void clearVbaDebugOutput(){const auto revision=output.revision+1;output=Output{};output.revision=revision;}
std::uint64_t vbaDebugOutputRevision() noexcept{return output.revision;}
std::size_t vbaDebugOutputColumn() noexcept{return output.column;}
void appendVbaDebugOutput(std::string_view text){
    if(text.empty())return;const auto units=nativeUtf16(text);auto column=output.column;
    for(const auto unit:units){if(unit==u'\r' || unit==u'\n')column=1;else{if(column==std::numeric_limits<std::size_t>::max())throw VbaRuntimeError(6,"Debug output column overflow");++column;}}
    auto retained=text;std::size_t start=retained.size()>byteLimit ? retained.size()-byteLimit : 0;
    while(start<retained.size() && (static_cast<unsigned char>(retained[start])&0xc0)==0x80)++start;
    retained.remove_prefix(start);if(start)output.truncated=true;
    auto lineCount=lines(retained);
    if(lineCount>lineLimit){std::size_t excess=lineCount-lineLimit,offset=0;while(excess && offset<retained.size()){const auto length=lineBreak(retained,offset);if(length)--excess;offset+=length ? length : 1;}retained.remove_prefix(offset);lineCount=lines(retained);output.truncated=true;}
    for(std::size_t offset=0;offset<retained.size();){auto end=std::min(offset+4096,retained.size());while(end<retained.size() && end>offset && (static_cast<unsigned char>(retained[end])&0xc0)==0x80)--end;const auto piece=retained.substr(offset,end-offset);const auto pieceLines=lines(piece);
        while(!output.chunks.empty() && (output.bytes+piece.size()>byteLimit || output.lines+pieceLines>lineLimit)){output.bytes-=output.chunks.front().text.size();output.lines-=output.chunks.front().lines;output.chunks.pop_front();output.truncated=true;}
        if(!output.chunks.empty() && output.chunks.back().text.size()+piece.size()<=4096){output.chunks.back().text.append(piece);output.chunks.back().lines+=pieceLines;}else output.chunks.push_back({std::string(piece),pieceLines});output.bytes+=piece.size();output.lines+=pieceLines;offset=end;}
    output.column=column;++output.revision;
}
VbaDebugOutputSnapshot vbaDebugOutputSnapshot(){VbaDebugOutputSnapshot snapshot;snapshot.text.reserve(output.bytes);for(const auto& chunk:output.chunks)snapshot.text+=chunk.text;snapshot.revision=output.revision;snapshot.column=output.column;snapshot.truncated=output.truncated;return snapshot;}
std::optional<std::vector<VbaDebugPrintItem>> parseVbaDebugPrint(std::string_view input,const std::function<bool(std::string_view)>& validExpression){
    const auto text=trim(input);if(text.size()>65536)return {};std::vector<VbaDebugPrintItem> items;if(text.empty())return items;
    const auto add=[&](std::string_view raw,char delimiter){if(items.size()>=256)return false;auto clause=trim(raw);VbaDebugPrintItem item;item.delimiter=delimiter;const auto open=clause.find('(');const auto function=lower(trim(clause.substr(0,open)));
        if(lower(clause)=="tab")item.kind=VbaDebugPrintItem::Kind::Zone;
        else if((function=="spc" || function=="tab") && open!=std::string::npos && clause.back()==')'){item.kind=function=="spc" ? VbaDebugPrintItem::Kind::Spaces : VbaDebugPrintItem::Kind::Tab;item.expression=trim(std::string_view(clause).substr(open+1,clause.size()-open-2));if(item.expression.empty() || !validExpression(item.expression))return false;}
        else{item.expression=std::move(clause);if(!item.expression.empty() && !validExpression(item.expression))return false;}
        items.push_back(std::move(item));return true;};
    std::size_t start=0;int depth=0;bool quoted=false,bracket=false,date=false;
    for(std::size_t i=0;i<text.size();++i){const auto c=text[i];if(quoted){if(c=='"'){if(i+1<text.size() && text[i+1]=='"')++i;else quoted=false;}continue;}if(bracket){if(c==']'){if(i+1<text.size() && text[i+1]==']')++i;else bracket=false;}continue;}if(date){if(c=='#')date=false;continue;}
        if(c=='"')quoted=true;else if(c=='[')bracket=true;else if(c=='#')date=true;else if(c=='('){if(++depth>64)return {};}else if(c==')'){if(--depth<0)return {};}else if(depth==0 && (c==';' || c==',')){if(!add(std::string_view(text).substr(start,i-start),c))return {};start=i+1;}}
    if(quoted || bracket || date || depth)return {};if(start<text.size() && !add(std::string_view(text).substr(start),0))return {};return items;
}
void executeVbaDebugPrint(const std::vector<VbaDebugPrintItem>& items,const std::function<VbaValue(std::string_view)>& evaluate){
    std::size_t bytes=0;const auto emit=[&](std::string_view text){if(text.size()>byteLimit-bytes)throw VbaRuntimeError(3035,"Debug.Print statement output limit exceeded");appendVbaDebugOutput(text);bytes+=text.size();};
    const auto zone=[&]{emit(std::string(14-(vbaDebugOutputColumn()-1)%14,' '));};
    for(const auto& item:items){if(item.kind==VbaDebugPrintItem::Kind::Zone)zone();
        else if(item.kind==VbaDebugPrintItem::Kind::Spaces || item.kind==VbaDebugPrintItem::Kind::Tab){const auto value=coerceVbaModuleValue(scalarVbaValue(evaluate(item.expression)),"integer");const auto count=static_cast<int>(value.number);if(item.kind==VbaDebugPrintItem::Kind::Spaces)emit(std::string(static_cast<std::size_t>(std::max(0,count)),' '));else{const auto target=static_cast<std::size_t>(std::max(1,count));if(target<=vbaDebugOutputColumn())emit("\n");emit(std::string(target-vbaDebugOutputColumn(),' '));}}
        else if(!item.expression.empty()){const auto value=scalarVbaValue(evaluate(item.expression));if(value.kind==VbaValue::Kind::Null)emit("Null");else if(value.kind==VbaValue::Kind::Boolean)emit(value.boolean ? "True" : "False");else if(value.kind==VbaValue::Kind::Error)emit("Error "+std::to_string(value.errorNumber));else{const auto text=coerceVbaModuleValue(value,"string").text;const bool number=value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Currency || value.kind==VbaValue::Kind::Decimal;if(number)emit(" ");emit(text);if(number)emit(" ");}}
        if(item.delimiter==',')zone();}
    if(items.empty() || !items.back().delimiter)emit("\n");
}
}
