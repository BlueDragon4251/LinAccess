#include "access_form_objects.hpp"
#include "access_form_preview.hpp"
#include "vba_qt_value.hpp"
#include <QApplication>
#include <QPointer>
#include <QWidget>
#include <algorithm>

std::shared_ptr<linaccess::VbaObject> makePreviewControlsObject(AccessFormPreview&);
std::shared_ptr<linaccess::VbaObject> makePreviewControlObject(AccessFormPreview&,const QString&);

namespace {
QString text(std::string_view value) { return QString::fromUtf8(value.data(),value.size()); }
std::string bytes(const QString& value) { return value.toUtf8().toStdString(); }
qsizetype ordinal(const linaccess::VbaValue& value,qsizetype count) {
    const auto index = linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
    if (index < 0 || index >= count) throw linaccess::VbaRuntimeError(2465,"Collection item unavailable");
    return index;
}
}
class PreviewObjectBase {
protected:
    explicit PreviewObjectBase(AccessFormPreview& preview) : preview_(&preview) {}
    AccessFormPreview& checked() const {
        if (!preview_ || preview_->closed_ || preview_->openCancelled_)
            throw linaccess::VbaRuntimeError(2450,"Form is closed or unavailable");
        return *preview_;
    }
    QPointer<AccessFormPreview> preview_;
};
class PreviewControlObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    PreviewControlObject(AccessFormPreview& preview,QString name) : PreviewObjectBase(preview),name_(std::move(name)) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked();
        auto* widget = control();
        if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected control arguments");
        const auto key = text(name).toCaseFolded();
        if (key == "name") return linaccess::VbaValue::string(bytes(name_));
        if (key == "value") { requireValue(); return preview.getFieldValue(bytes(name_)); }
        if (key == "visible") return linaccess::VbaValue::logical(!widget->isHidden());
        if (key == "enabled") return linaccess::VbaValue::logical(widget->isEnabled());
        if (key == "controltype") {
            static const QMap<QString,int> types{{"Label",100},{"Rectangle",101},{"Line",102},
                {"CommandButton",104},{"OptionButton",105},{"CheckBox",106},{"TextBox",109},
                {"ListBox",110},{"ComboBox",111},{"Subform",112},{"ToggleButton",122}};
            for (const auto& item : preview.layout_.controls)
                if (item.name.compare(name_,Qt::CaseInsensitive)==0 && types.contains(item.type))
                    return linaccess::VbaValue::numeric(types.value(item.type));
            throw linaccess::VbaRuntimeError(438,"Control type unsupported");
        }
        if (key == "setfocus") {
            if (widget->isHidden() || !widget->isEnabled()) throw linaccess::VbaRuntimeError(2110,"Control cannot receive focus");
            widget->setFocus(); return {};
        }
        if (key == "requery") { preview.requeryControl(bytes(name_)); return {}; }
        throw linaccess::VbaRuntimeError(438,"Control member unsupported");
    }
    void assign(std::string_view name,const linaccess::VbaValue& value) override {
        auto& preview = checked(); auto* widget = control();
        const auto key = text(name).toCaseFolded();
        if (key == "value") { requireValue(); preview.setFieldValue(bytes(name_),value); return; }
        if ((key == "visible" || key == "enabled") && value.kind == linaccess::VbaValue::Kind::Null)
            throw linaccess::VbaRuntimeError(94,"Invalid Null property assignment");
        if (key == "visible") { widget->setVisible(value.truthy()); return; }
        if (key == "enabled") { widget->setEnabled(value.truthy()); return; }
        throw linaccess::VbaRuntimeError(438,"Control property is unavailable or read-only");
    }
private:
    void requireValue() const {
        for (const auto& control : checked().layout_.controls)
            if (control.name.compare(name_,Qt::CaseInsensitive)==0 &&
                (control.type=="TextBox" || control.type=="ComboBox" || control.type=="ListBox" ||
                 control.type=="CheckBox" || control.type=="OptionButton" || control.type=="ToggleButton")) return;
        throw linaccess::VbaRuntimeError(438,"Control has no supported Value property");
    }
    QWidget* control() const {
        auto* widget = checked().controls_.value(name_.toCaseFolded());
        if (!widget) throw linaccess::VbaRuntimeError(2465,"Control unavailable");
        return widget;
    }
    QString name_;
};
class PreviewControlsObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewControlsObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked(); const auto names = controlNames(); const auto key = text(name).toCaseFolded();
        if (key == "count" && args.empty()) return linaccess::VbaValue::numeric(names.size());
        QString target;
        if (key == "item" && args.size()==1) {
            target = args[0].kind == linaccess::VbaValue::Kind::Number
                ? names[ordinal(args[0],names.size())] : text(args[0].asText());
        } else if (args.empty()) target = text(name);
        else throw linaccess::VbaRuntimeError(450,"Unexpected Controls arguments");
        for (const auto& item : names) if (item.compare(target,Qt::CaseInsensitive)==0)
            return linaccess::VbaValue::reference(makePreviewControlObject(preview,item));
        throw linaccess::VbaRuntimeError(2465,"Control unavailable");
    }
    linaccess::VbaValue item(std::string_view name) override { return member("Item",{linaccess::VbaValue::string(std::string(name))}); }
    std::vector<linaccess::VbaValue> enumerate() override {
        auto& preview = checked(); std::vector<linaccess::VbaValue> values;
        for (const auto& name : controlNames())
            values.push_back(linaccess::VbaValue::reference(makePreviewControlObject(preview,name)));
        return values;
    }
private:
    QStringList controlNames() const {
        const auto& preview = checked(); QStringList names;
        for (const auto& control : preview.layout_.controls)
            if (!control.name.isEmpty() && preview.controls_.contains(control.name.toCaseFolded()) && !names.contains(control.name,Qt::CaseInsensitive))
                names.append(control.name);
        return names;
    }
};
class PreviewFormObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewFormObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked(); const auto key = text(name).toCaseFolded();
        if (key == "controls") {
            auto controls = makePreviewControlsObject(preview);
            if (args.empty()) return linaccess::VbaValue::reference(controls);
            return controls->member("Item",args);
        }
        if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected form arguments");
        if (key == "name") return linaccess::VbaValue::string(bytes(preview.formName_));
        if (key == "recordsource") return linaccess::VbaValue::string(bytes(preview.layout_.recordSource));
        if (key == "dirty") return linaccess::VbaValue::logical(preview.dirty_);
        if (key == "newrecord") return linaccess::VbaValue::logical(preview.currentRecordIndex_==0);
        if (key == "requery") {
            if (!preview.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Requery cancelled");
            preview.loadRecords(); return {};
        }
        if (key == "setfocus") { preview.activateWindow(); preview.setFocus(); return {}; }
        return PreviewControlObject(preview,text(name)).member("Value",{});
    }
    linaccess::VbaValue item(std::string_view name) override {
        return PreviewControlObject(checked(),text(name)).member("Value",{});
    }
    void assignItem(std::string_view name,const linaccess::VbaValue& value) override {
        PreviewControlObject(checked(),text(name)).assign("Value",value);
    }
    void assign(std::string_view name,const linaccess::VbaValue& value) override {
        auto& preview = checked(); const auto key = text(name).toCaseFolded();
        if (key == "dirty") {
            if (!value.truthy()) {
                if (!preview.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Record save cancelled");
            } else {
                if (!preview.directTable_) throw linaccess::VbaRuntimeError(2448,"Record source is read-only");
                preview.markDirty({});
            }
            return;
        }
        if (key == "name" || key == "recordsource" || key == "newrecord")
            throw linaccess::VbaRuntimeError(438,"Form property is read-only");
        PreviewControlObject(preview,text(name)).assign("Value",value);
    }
};
class PreviewFormsObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewFormsObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        const auto forms = openForms(); const auto key = text(name).toCaseFolded();
        if (key == "count" && args.empty()) return linaccess::VbaValue::numeric(forms.size());
        QString target;
        if (key == "item" && args.size()==1) {
            if (args[0].kind == linaccess::VbaValue::Kind::Number)
                return linaccess::VbaValue::reference(makePreviewFormObject(*forms[ordinal(args[0],forms.size())]));
            target = text(args[0].asText());
        } else if (args.empty()) target = text(name);
        else throw linaccess::VbaRuntimeError(450,"Unexpected Forms arguments");
        for (auto* form : forms) if (form->formName_.compare(target,Qt::CaseInsensitive)==0)
            return linaccess::VbaValue::reference(makePreviewFormObject(*form));
        throw linaccess::VbaRuntimeError(2450,"Form is not open");
    }
    linaccess::VbaValue item(std::string_view name) override { return member("Item",{linaccess::VbaValue::string(std::string(name))}); }
    std::vector<linaccess::VbaValue> enumerate() override {
        std::vector<linaccess::VbaValue> values;
        for (auto* form : openForms()) values.push_back(linaccess::VbaValue::reference(makePreviewFormObject(*form)));
        return values;
    }
private:
    QVector<AccessFormPreview*> openForms() const {
        const auto& owner = checked(); QVector<AccessFormPreview*> forms;
        for (auto* widget : QApplication::topLevelWidgets())
            if (auto* form = dynamic_cast<AccessFormPreview*>(widget); form && !form->closed_ && !form->openCancelled_
                && form->db_.connectionName()==owner.db_.connectionName()) forms.append(form);
        std::sort(forms.begin(),forms.end(),[](auto* left,auto* right) {
            return left->property("la_form_open_order").toLongLong() < right->property("la_form_open_order").toLongLong();
        });
        return forms;
    }
};

std::shared_ptr<linaccess::VbaObject> makePreviewFormObject(AccessFormPreview& preview) {
    if (auto object = preview.formObject_.lock()) return object;
    auto object = std::make_shared<PreviewFormObject>(preview);
    preview.formObject_ = object; return object;
}
std::shared_ptr<linaccess::VbaObject> makePreviewFormsObject(AccessFormPreview& preview) {
    if (auto object = preview.formsObject_.lock()) return object;
    auto object = std::make_shared<PreviewFormsObject>(preview);
    preview.formsObject_ = object; return object;
}
std::shared_ptr<linaccess::VbaObject> makePreviewControlsObject(AccessFormPreview& preview) {
    if (auto object = preview.controlsObject_.lock()) return object;
    auto object = std::make_shared<PreviewControlsObject>(preview);
    preview.controlsObject_ = object; return object;
}
std::shared_ptr<linaccess::VbaObject> makePreviewControlObject(AccessFormPreview& preview,const QString& name) {
    auto& cached = preview.controlObjects_[name.toCaseFolded()];
    if (auto object = cached.lock()) return object;
    auto object = std::make_shared<PreviewControlObject>(preview,name);
    cached = object; return object;
}
