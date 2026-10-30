#include "access_form_import.hpp"

#include <QRegularExpression>
#include <QSet>
#include <QHash>

namespace linaccess {
AccessFormDraft parseAccessFormText(const QString& source) {
    AccessFormDraft draft;
    // The export consists of many nested Begin/End blocks. For the simple
    // bound-form import we need only top-level RecordSource and field bindings.
    // Text after the first embedded quote is deliberately not evaluated.
    const QRegularExpression record(R"re(^\s*RecordSource\s*=\s*"([^"]*)")re",
        QRegularExpression::MultilineOption | QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression control(R"re(^\s*ControlSource\s*=\s*"([^"]*)")re",
        QRegularExpression::MultilineOption | QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression event(R"re(^\s*(OnOpen|OnLoad|AfterUpdate|OnClick|OnCurrent)\s*=\s*"\[Event Procedure\]")re",
        QRegularExpression::MultilineOption | QRegularExpression::CaseInsensitiveOption);
    const auto table = record.match(source);
    if (table.hasMatch()) draft.recordSource = table.captured(1).trimmed();

    QSet<QString> seen;
    auto fields = control.globalMatch(source);
    while (fields.hasNext()) {
        const auto name = fields.next().captured(1).trimmed();
        if (!name.isEmpty() && !seen.contains(name.toCaseFolded())) {
            seen.insert(name.toCaseFolded());
            draft.controlSources << name;
        }
    }
    auto events = event.globalMatch(source);
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
        if (value[index] == '"') {
            if (index + 1 < value.size() && value[index + 1] == '"') {
                result += '"';
                ++index;
            } else break;
        } else result += value[index];
    }
    return result;
}

bool isControl(const QString& type) {
    static const QSet<QString> types = {
        "TextBox", "Label", "CommandButton", "ComboBox", "ListBox",
        "CheckBox", "OptionButton", "ToggleButton", "Subform",
        "Rectangle", "Line", "Image", "Tab", "OptionGroup",
        "BoundObjectFrame", "UnboundObjectFrame"
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
    const auto lines = source.split('\n');
    for (const auto& rawLine : lines) {
        const auto line = rawLine.trimmed();
        const auto start = begin.match(line);
        if (start.hasMatch()) {
            const auto type = start.captured(1);
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
        if (value.compare("Begin", Qt::CaseInsensitive) == 0) {
            stack.push_back({"BinaryProperty", -1, current.section});
            continue;
        }
        const auto integer = [&] { bool ok = false; const int result = value.toInt(&ok); return ok ? result : 0; };
        if (current.control >= 0) {
            auto& control = form.controls[current.control];
            if (key == "Name") control.name = value;
            else if (key == "Caption") control.caption = value;
            else if (key == "ControlSource") control.controlSource = value;
            else if (key == "RowSource") control.rowSource = value;
            else if (key == "SourceObject") control.sourceObject = value;
            else if (key == "OnClick") control.onClick = value;
            else if (key == "AfterUpdate") control.afterUpdate = value;
            else if (key == "Left") control.left = integer();
            else if (key == "Top") control.top = integer();
            else if (key == "Width") control.width = integer();
            else if (key == "Height") control.height = integer();
            else if (key == "Visible") control.visible = value != "0" && value != "False";
        } else if (current.type == "Form") {
            if (key == "Caption") form.caption = value;
            else if (key == "RecordSource") form.recordSource = value;
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
