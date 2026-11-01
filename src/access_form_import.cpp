#include "access_form_import.hpp"
#include <stdexcept>
#include <limits>

#include <QRegularExpression>
#include <QSet>
#include <QHash>

namespace linaccess {
namespace {
QString propertyValue(const QString& raw);
QStringList saveAsTextLines(const QString& source);
}
AccessFormDraft parseAccessFormText(const QString& source) {
    AccessFormDraft draft;
    const auto logicalSource=saveAsTextLines(source).join('\n');
    // The export consists of many nested Begin/End blocks. For the simple
    // bound-form import we need only top-level RecordSource and field bindings.
    // SaveAsText splits long strings into adjacent quoted fragments.
    const QRegularExpression record(R"re(^\s*RecordSource\s*=\s*(.+)$)re",
        QRegularExpression::MultilineOption | QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression control(R"re(^\s*ControlSource\s*=\s*(.+)$)re",
        QRegularExpression::MultilineOption | QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression event(R"re(^\s*(OnOpen|OnLoad|AfterUpdate|OnClick|OnCurrent)\s*=\s*"\[Event Procedure\]")re",
        QRegularExpression::MultilineOption | QRegularExpression::CaseInsensitiveOption);
    const auto table = record.match(logicalSource);
    if (table.hasMatch()) draft.recordSource = propertyValue(table.captured(1)).trimmed();

    QSet<QString> seen;
    auto fields = control.globalMatch(logicalSource);
    while (fields.hasNext()) {
        const auto name = propertyValue(fields.next().captured(1)).trimmed();
        if (!name.isEmpty() && !seen.contains(name.toCaseFolded())) {
            seen.insert(name.toCaseFolded());
            draft.controlSources << name;
        }
    }
    auto events = event.globalMatch(logicalSource);
    while (events.hasNext()) {
        const auto name = events.next().captured(1);
        if (!draft.eventNames.contains(name, Qt::CaseInsensitive)) draft.eventNames << name;
    }
    return draft;
}

namespace {
QString propertyValue(const QString& raw) {
    const auto value = raw.trimmed();
    if (!value.startsWith('"')) return value;
    QString result;
    for (qsizetype index = 1; index < value.size(); ++index) {
        if (value[index] == '\\' && index+1<value.size() && value[index+1]=='"') {
            result += '"';
            ++index;
        } else if (value[index] == '"') {
            if (index + 1 < value.size() && value[index + 1] == '"') {
                result += '"';
                ++index;
            } else break;
        } else result += value[index];
    }
    return result;
}

QStringList saveAsTextLines(const QString& source) {
    const auto physical=source.split('\n');
    QStringList logical;
    static const QRegularExpression property(R"(^\s*[A-Za-z][A-Za-z0-9_]*\s*=\s*(".*)$)");
    for (qsizetype i=0;i<physical.size();++i) {
        auto line=physical[i];
        const auto match=property.match(line);
        if (match.hasMatch() && i+1<physical.size() && physical[i+1].trimmed().startsWith('"')) {
            auto value=propertyValue(match.captured(1));
            while (i+1<physical.size() && physical[i+1].trimmed().startsWith('"')) value+=propertyValue(physical[++i]);
            value.replace('"',"\"\"");
            line=line.left(match.capturedStart(1))+'"'+value+'"';
        }
        logical.append(line);
    }
    return logical;
}
bool accessBoolean(const QString& value,bool defaultValue) {
    if (value.compare("NotDefault",Qt::CaseInsensitive)==0) return !defaultValue;
    if (value.compare("True",Qt::CaseInsensitive)==0 || value.compare("Yes",Qt::CaseInsensitive)==0) return true;
    if (value.compare("False",Qt::CaseInsensitive)==0 || value.compare("No",Qt::CaseInsensitive)==0) return false;
    bool valid=false; const auto number=value.toLongLong(&valid);
    if (!valid) throw std::runtime_error("Invalid imported Boolean property");
    return number!=0;
}
bool isControl(const QString& type) {
    static const QSet<QString> types = {
        "TextBox", "Label", "CommandButton", "ComboBox", "ListBox",
        "CheckBox", "OptionButton", "ToggleButton", "Subform",
        "Rectangle", "Line", "Image", "Tab", "OptionGroup",
        "BoundObjectFrame", "UnboundObjectFrame", "CustomControl"
    };
    return types.contains(type);
}
}

AccessFormLayout parseAccessFormLayout(const QString& source) {
    AccessFormLayout form;
    struct Block { QString type; int control{-1}; QString section; };
    QVector<Block> stack;
    QHash<QString, int> sections;
    static const QRegularExpression begin(R"(^\s*Begin(?:\s+(\w+))?\s*$)",
        QRegularExpression::CaseInsensitiveOption);
    static const QRegularExpression property(R"(^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$)");
    const auto lines = saveAsTextLines(source);
    for (const auto& rawLine : lines) {
        const auto line = rawLine.trimmed();
        const auto start = begin.match(line);
        if (start.hasMatch()) {
            const auto type = start.captured(1);
            if (stack.isEmpty() && type.compare("Report",Qt::CaseInsensitive)==0) form.kind = AccessDocumentKind::Report;
            QString section;
            for (auto it = stack.crbegin(); it != stack.crend(); ++it)
                if (!it->section.isEmpty()) { section = it->section; break; }
            if (type == "Section" || type == "FormHeader" || type == "FormFooter"
                || type == "PageHeader" || type == "PageFooter") section = type;
            int control = -1;
            // The first unnamed Begin inside Form defines the Access defaults
            // for control types; these are not actual instantiated controls.
            if (isControl(type) && !section.isEmpty()) {
                control = static_cast<int>(form.controls.size());
                AccessVisualControl item;
                item.type = type;
                for (auto it=stack.crbegin();it!=stack.crend();++it)
                    if (it->type=="OptionGroup" && it->control>=0) { item.optionGroupIndex=it->control; break; }
                item.section = section;
                form.controls.push_back(item);
            }
            stack.push_back({type, control, section});
            continue;
        }
        if (line.compare("End", Qt::CaseInsensitive) == 0) {
            if (!stack.empty()) stack.pop_back();
            continue;
        }
        const auto assignment = property.match(line);
        if (!assignment.hasMatch() || stack.empty()) continue;
        const auto key = assignment.captured(1);
        const auto value = propertyValue(assignment.captured(2));
        const auto current = stack.back();
        // GUID/NameMap/embedded-image values use `Property = Begin` followed
        // by hex bytes and `End`. That End is *not* the enclosing control End.
        static const QRegularExpression binaryBegin(R"(^(?:\{[0-9, \t-]+\}\s*)?Begin$)",
            QRegularExpression::CaseInsensitiveOption);
        if (binaryBegin.match(assignment.captured(2).trimmed()).hasMatch()) {
            stack.push_back({"BinaryProperty", -1, current.section});
            continue;
        }
        const auto integer = [&] { bool ok = false; const int result = value.toInt(&ok); return ok ? result : 0; };
        if (current.control >= 0) {
            auto& control = form.controls[current.control];
            if (key == "Name") control.name = value;
            else if (key == "EventProcPrefix") control.eventProcPrefix = value;
            else if (key == "Caption") control.caption = value;
            else if (key == "Tag") control.tag = value;
            else if (key == "ControlSource") control.controlSource = value;
            else if (key == "DefaultValue") control.defaultValue = value;
            else if (key == "InputMask") control.inputMask = value;
            else if (key == "OptionValue") control.optionValue = integer();
            else if (key == "RowSource") control.rowSource = value;
            else if (key == "RowSourceType") control.rowSourceType = value;
            else if (key == "BoundColumn") control.boundColumn = integer();
            else if (key == "ColumnCount") control.columnCount = integer();
            else if (key == "ColumnWidths") control.columnWidths = value;
            else if (key == "SourceObject") control.sourceObject = value;
            else if (key == "LinkMasterFields") control.linkMasterFields = value;
            else if (key == "LinkChildFields") control.linkChildFields = value;
            else if (key == "OnClick") control.onClick = value;
            else if (key == "OnDblClick") control.onDblClick = value;
            else if (key == "OnKeyDown") control.onKeyDown = value;
            else if (key == "OnKeyPress") control.onKeyPress = value;
            else if (key == "OnChange") control.onChange = value;
            else if (key == "OnKeyUp") control.onKeyUp = value;
            else if (key == "Class") control.className = value;
            else if (key == "OLEClass") control.oleClass = value;
            else if (key == "AfterUpdate") control.afterUpdate = value;
            else if (key == "BeforeUpdate") control.beforeUpdate = value;
            else if (key == "Left") control.left = integer();
            else if (key == "Top") control.top = integer();
            else if (key == "Width") control.width = integer();
            else if (key == "Height") control.height = integer();
            else if (key == "Visible") control.visible = accessBoolean(value,true);
            else if (key == "Enabled") control.enabled = accessBoolean(value,true);
            else if (key == "Locked") control.locked = accessBoolean(value,false);
            else if (key == "TabStop") control.tabStop = accessBoolean(value,true);
            else if (key == "TabIndex") {
                bool valid=false; const auto index=value.toInt(&valid);
                if (!valid || index<0 || index>32767) throw std::runtime_error("Invalid imported TabIndex");
                control.tabIndex=index;
            }
        } else if (current.type == "Form" || current.type == "Report") {
            static const QSet<QString> events{"beforeinsert","afterinsert","ondelete","beforedelconfirm","afterdelconfirm","onopen","onload","oncurrent","ondirty","beforeupdate","afterupdate","onunload","onclose","onactivate","ondeactivate","onnodata","onerror","onpage","onkeydown","onkeypress","onkeyup","ontimer"};
            if (events.contains(key.toCaseFolded())) form.events.insert(key.toCaseFolded(),value);
            if (key == "AllowAdditions") form.allowAdditions = accessBoolean(value,true);
            if (key == "AllowDeletions") form.allowDeletions = accessBoolean(value,true);
            if (key == "Caption") form.caption = value;
            else if (key == "Tag") form.tag = value;
            else if (key == "RecordSource") form.recordSource = value;
            else if (key == "Filter") form.filter = value;
            else if (key == "TimerInterval") {
                bool valid=false; const auto interval=value.toLongLong(&valid);
                if (!valid || interval<0 || interval>std::numeric_limits<int>::max()) throw std::runtime_error("Invalid imported TimerInterval");
                form.timerInterval=static_cast<int>(interval);
            }
            else if (key == "KeyPreview") form.keyPreview = accessBoolean(value,false);
            else if (key == "Modal") form.modal = accessBoolean(value,false);
            else if (key == "FilterOnLoad") form.filterOnLoad = accessBoolean(value,false);
            else if (key == "Width") form.width = integer();
        } else if (current.type == "Section" || current.type == "FormHeader"
                   || current.type == "FormFooter") {
            if (key == "Height") sections[current.type] = integer();
        }
    }
    const int header = sections.value("FormHeader");
    const int detail = sections.value("Section");
    const int footer = sections.value("FormFooter");
    for (auto& control : form.controls) {
        if (control.section == "Section") control.top += header;
        else if (control.section == "FormFooter") control.top += header + detail;
    }
    form.height = qMax(2000, header + detail + footer);
    return form;
}
}
