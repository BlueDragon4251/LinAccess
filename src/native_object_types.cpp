#include "native_object_types.hpp"
#include <utility>

namespace linaccess {
namespace {
bool primitiveType(std::string_view type) {
    switch(type.size()) {
    case 4:return type=="byte" || type=="date" || type=="long";
    case 6:return type=="string" || type=="single" || type=="double";
    case 7:return type=="boolean" || type=="integer" || type=="variant" || type=="decimal";
    case 8:return type=="currency";
    default:return false;
    }
}
}
// Closed native class vocabulary. Project classes, reference ordering and COM
// interfaces require a real binder, not a guessed match against TypeName.
std::string_view nativeAccessObjectType(std::string_view type) {
    if(primitiveType(type)) return {};
    static constexpr std::pair<std::string_view,std::string_view> types[]{
        {"form","access.form"},{"report","access.report"},{"grouplevel","access.grouplevel"},{"control","access.control"},{"section","access.section"},
        {"application","access.application"},{"forms","access.forms"},{"reports","access.reports"},{"controls","access.controls"},
        {"textbox","access.textbox"},{"combobox","access.combobox"},{"listbox","access.listbox"},
        {"checkbox","access.checkbox"},{"optionbutton","access.optionbutton"},{"togglebutton","access.togglebutton"},
        {"commandbutton","access.commandbutton"},{"label","access.label"},{"rectangle","access.rectangle"},
        {"line","access.line"},{"subform","access.subform"},{"optiongroup","access.optiongroup"},
        {"tabcontrol","access.tabcontrol"},{"page","access.page"},
        {"filedialog","office.filedialog"},{"filedialogfilters","office.filedialogfilters"},
        {"filedialogfilter","office.filedialogfilter"},{"filedialogselecteditems","office.filedialogselecteditems"}};
    for(const auto& [shortName,qualified]:types) if(type==shortName || type==qualified) return qualified;
    return {};
}

std::string_view nativeScriptingObjectType(std::string_view type) {
    if(!type.starts_with("scripting.")) return {};
    if(type=="scripting.dictionary") return "scripting.dictionary";
    if(type=="scripting.textstream") return "scripting.textstream";
    if(type=="scripting.folders") return "scripting.folders";
    if(type=="scripting.folder") return "scripting.folder";
    if(type=="scripting.files") return "scripting.files";
    if(type=="scripting.file") return "scripting.file";
    if(type=="scripting.filesystemobject") return "scripting.filesystemobject";
    return {};
}

std::string_view nativeObjectType(std::string_view type) {
    if(primitiveType(type)) return {};
    if(const auto access=nativeAccessObjectType(type);!access.empty()) return access;
    if(const auto scripting=nativeScriptingObjectType(type);!scripting.empty()) return scripting;
    if(type=="collection" || type=="vba.collection") return "vba.collection";
    if(type=="errobject" || type=="vba.errobject") return "vba.errobject";
    static constexpr std::string_view types[]{"object",
        "dao.database","dao.databases","dao.workspace","dao.workspaces","dao.dbengine","dao.recordset",
        "dao.field","dao.fields","dao.querydef","dao.querydefs","dao.parameter","dao.parameters","dao.tabledef","dao.tabledefs","dao.relation","dao.relations","dao.index","dao.indexes","dao.property","dao.properties"};
    for(const auto name:types) if(type==name) return name;
    return {};
}
}
