#pragma once
#include <array>
#include <string_view>
#include <utility>

namespace linaccess {
// Closed native class vocabulary. Project classes, reference ordering and COM
// interfaces require a real binder, not a guessed match against TypeName.
inline std::string_view nativeAccessObjectType(std::string_view type) {
    static constexpr std::pair<std::string_view,std::string_view> types[]{
        {"form","access.form"},{"report","access.report"},{"control","access.control"},
        {"application","access.application"},{"forms","access.forms"},{"reports","access.reports"},{"controls","access.controls"},
        {"textbox","access.textbox"},{"combobox","access.combobox"},{"listbox","access.listbox"},
        {"checkbox","access.checkbox"},{"optionbutton","access.optionbutton"},{"togglebutton","access.togglebutton"},
        {"commandbutton","access.commandbutton"},{"label","access.label"},{"rectangle","access.rectangle"},
        {"line","access.line"},{"subform","access.subform"},{"optiongroup","access.optiongroup"},
        {"tabcontrol","access.tabcontrol"},{"page","access.page"}};
    for(const auto& [shortName,qualified]:types) if(type==shortName || type==qualified) return qualified;
    return {};
}

inline std::string_view nativeObjectType(std::string_view type) {
    if(const auto access=nativeAccessObjectType(type);!access.empty()) return access;
    if(type=="collection" || type=="vba.collection") return "vba.collection";
    if(type=="errobject" || type=="vba.errobject") return "vba.errobject";
    static constexpr std::string_view types[]{"object","scripting.dictionary",
        "dao.database","dao.databases","dao.workspace","dao.workspaces","dao.dbengine","dao.recordset",
        "dao.field","dao.fields","dao.querydef","dao.parameter","dao.parameters","dao.tabledef","dao.tabledefs"};
    for(const auto name:types) if(type==name) return name;
    return {};
}
}
