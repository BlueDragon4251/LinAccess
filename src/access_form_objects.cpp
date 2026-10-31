#include "access_form_objects.hpp"
#include "jet_sql_functions.hpp"
#include <QThread>
#include "access_form_preview.hpp"
#include "vba_qt_value.hpp"
#include <QApplication>
#include <QAbstractButton>
#include <QAction>
#include <QLabel>
#include <QComboBox>
#include <QCalendarWidget>
#include <QToolBar>
#include <QMainWindow>
#include <QMenuBar>
#include <QMenu>
#include <QWidgetAction>
#include <QPointer>
#include <QWidget>
#include <algorithm>
#include <limits>

std::shared_ptr<linaccess::VbaObject> makePreviewControlsObject(AccessFormPreview&);
std::shared_ptr<linaccess::VbaObject> makePreviewControlObject(AccessFormPreview&,const QString&);

std::shared_ptr<linaccess::VbaObject> makePreviewCommandBarObject(AccessFormPreview&,QWidget*);
std::shared_ptr<linaccess::VbaObject> makePreviewCommandControlsObject(AccessFormPreview&,QWidget*);

std::shared_ptr<linaccess::VbaObject> makePreviewCommandControlObject(AccessFormPreview&,QAction*);

namespace {
QList<QAction*> commandActions(QWidget* bar) {
    QList<QAction*> actions;
    for (auto* action : bar->actions())
        if (!action->isSeparator() && !qobject_cast<QWidgetAction*>(action)) actions.append(action);
    return actions;
}
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
    linaccess::VbaValue scalarValue() override { return member("Value",{}); }
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked();
        auto* widget = control();
        const auto key = text(name).toCaseFolded();
        if (key=="column" || key=="spalte") return preview.controlColumnValue(name_,args);
        if (!args.empty() && key!="form" && key!="formular") throw linaccess::VbaRuntimeError(450,"Unexpected control arguments");
        if (key=="form" || key=="formular") {
            if (auto* child=preview.subforms_.value(name_.toCaseFolded()).data()) {
                if (args.empty()) return linaccess::VbaValue::reference(makePreviewFormObject(*child));
                return makePreviewFormObject(*child)->defaultMember(args,false);
            }
            if (preview.subformErrors_.contains(name_.toCaseFolded()))
                throw linaccess::VbaRuntimeError(preview.subformErrors_.value(name_.toCaseFolded()),"Subform unavailable");
            throw linaccess::VbaRuntimeError(438,"Control has no Form property");
        }
        if (key=="parent") return linaccess::VbaValue::reference(makePreviewFormObject(preview));
        if (auto* calendar=qobject_cast<QCalendarWidget*>(widget);calendar && (key=="nextmonth" || key=="previousmonth")) {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Calendar navigation takes no arguments");
            if (key=="nextmonth") calendar->showNextMonth(); else calendar->showPreviousMonth(); return {};
        }
        if (key == "name") return linaccess::VbaValue::string(bytes(name_));
        if (key == "tag") return linaccess::VbaValue::string(bytes(widget->property("la_access_tag").toString()));
        if (key == "caption") return linaccess::VbaValue::string(bytes(caption()));
        if (key == "value") { requireValue(); return preview.getFieldValue(bytes(name_)); }
        if (key == "visible") return linaccess::VbaValue::logical(!widget->isHidden());
        if (key == "enabled") return linaccess::VbaValue::logical(widget->isEnabled());
        if (key == "controltype") {
            static const QMap<QString,int> types{{"Label",100},{"Rectangle",101},{"Line",102},
                {"OptionGroup",103},{"CommandButton",104},{"OptionButton",105},{"CheckBox",106},{"TextBox",109},
                {"ListBox",110},{"ComboBox",111},{"Subform",112},{"ToggleButton",122}};
            for (const auto& item : preview.layout_.controls)
                if (item.name.compare(name_,Qt::CaseInsensitive)==0 && types.contains(item.type))
                    return linaccess::VbaValue::numeric(types.value(item.type));
            if (qobject_cast<QCalendarWidget*>(widget)) return linaccess::VbaValue::numeric(119);
            throw linaccess::VbaRuntimeError(438,"Control type unsupported");
        }
        if (key=="dropdown") {
            auto* combo=qobject_cast<QComboBox*>(widget);
            const auto definition=std::find_if(preview.layout_.controls.cbegin(),preview.layout_.controls.cend(),[&](const auto& control) {
                return control.name.compare(name_,Qt::CaseInsensitive)==0 && control.type=="ComboBox";
            });
            if (!combo || definition==preview.layout_.controls.cend()) throw linaccess::VbaRuntimeError(438,"Dropdown requires a combo box");
            if (!combo->isVisible() || !combo->isEnabled()) throw linaccess::VbaRuntimeError(2110,"Combo box cannot open");
            if (!combo->property("la_rows_loaded").toBool()) preview.requeryControl(bytes(name_));
            combo->setFocus(); combo->showPopup(); return {};
        }
        if (key == "setfocus") {
            if (widget->isHidden() || !widget->isEnabled()) throw linaccess::VbaRuntimeError(2110,"Control cannot receive focus");
            widget->setFocus(); return {};
        }
        if (key == "requery") { preview.requeryControl(bytes(name_)); return {}; }
        throw linaccess::VbaRuntimeError(438,"Control member unsupported");
    }
    linaccess::VbaValue objectMember(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        if ((text(name).compare("Form",Qt::CaseInsensitive)==0 || text(name).compare("Formular",Qt::CaseInsensitive)==0) && !args.empty()) {
            const auto child=member(name,{}); return child.object->defaultMember(args,true);
        }
        return member(name,args);
    }
    void assign(std::string_view name,const linaccess::VbaValue& value) override {
        auto& preview = checked(); auto* widget = control();
        const auto key = text(name).toCaseFolded();
        if (key=="tag") {
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null Tag");
            control()->setProperty("la_access_tag",text(value.asText())); return;
        }
        if (key == "value") { requireValue(); preview.setFieldValue(bytes(name_),value); return; }
        if (key == "caption") {
            (void)caption();
            if (value.kind == linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null caption");
            const auto captionText = text(value.asText());
            if (auto* label = qobject_cast<QLabel*>(widget)) label->setText(captionText);
            else qobject_cast<QAbstractButton*>(widget)->setText(captionText);
            return;
        }
        if ((key == "visible" || key == "enabled") && value.kind == linaccess::VbaValue::Kind::Null)
            throw linaccess::VbaRuntimeError(94,"Invalid Null property assignment");
        if (key == "visible") { widget->setVisible(value.truthy()); return; }
        if (key == "enabled") { widget->setEnabled(value.truthy()); return; }
        throw linaccess::VbaRuntimeError(438,"Control property is unavailable or read-only");
    }
private:
    QString caption() const {
        auto* widget = control();
        for (const auto& control : checked().layout_.controls) {
            if (control.name.compare(name_,Qt::CaseInsensitive)!=0) continue;
            if (control.type=="Label") return qobject_cast<QLabel*>(widget)->text();
            if (control.type=="CommandButton" || control.type=="CheckBox" || control.type=="OptionButton" || control.type=="ToggleButton")
                return qobject_cast<QAbstractButton*>(widget)->text();
        }
        throw linaccess::VbaRuntimeError(438,"Control has no supported Caption property");
    }
    void requireValue() const {
        if (qobject_cast<QCalendarWidget*>(control())) return;
        for (const auto& control : checked().layout_.controls)
            if (control.name.compare(name_,Qt::CaseInsensitive)==0 &&
                (control.type=="OptionGroup" || control.type=="TextBox" || control.type=="ComboBox" || control.type=="ListBox" ||
                 control.type=="CheckBox" || control.type=="OptionButton" || control.type=="ToggleButton")) return;
        throw linaccess::VbaRuntimeError(438,"Control has no supported Value property");
    }
    QWidget* control() const {
        auto* widget = checked().controls_.value(name_.toCaseFolded()).data();
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
            if (!control.name.isEmpty() && preview.controls_.value(control.name.toCaseFolded()) && !names.contains(control.name,Qt::CaseInsensitive))
                names.append(control.name);
        return names;
    }
};
class PreviewFormObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewFormObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue defaultMember(const std::vector<linaccess::VbaValue>& args,bool reference) override {
        if (args.size()!=1) throw linaccess::VbaRuntimeError(450,"Form default member requires one control");
        const auto control=makePreviewControlsObject(checked())->member("Item",args);
        return reference ? control : control.object->member("Value",{});
    }
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked(); const auto key = text(name).toCaseFolded();
        if (key == "controls") {
            auto controls = makePreviewControlsObject(preview);
            if (args.empty()) return linaccess::VbaValue::reference(controls);
            return controls->member("Item",args);
        }
        if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected form arguments");
        if (key == "name") return linaccess::VbaValue::string(bytes(preview.formName_));
        if (key == "tag") return linaccess::VbaValue::string(bytes(preview.property("la_access_tag").toString()));
        if (key == "caption") return linaccess::VbaValue::string(bytes(preview.layout_.caption));
        if (key == "recordsource") return linaccess::VbaValue::string(bytes(preview.layout_.recordSource));
        if (key=="filter") return linaccess::VbaValue::string(bytes(preview.filter_));
        if (key=="filteron") return linaccess::VbaValue::logical(preview.filterOn_);
        if (preview.layout_.kind==linaccess::AccessDocumentKind::Report && (key=="dirty" || key=="newrecord"))
            throw linaccess::VbaRuntimeError(438,"Report member unsupported");
        if (key == "dirty") return linaccess::VbaValue::logical(preview.dirty_);
        if (key == "newrecord") return linaccess::VbaValue::logical(preview.currentRecordIndex_==0);
        if (key == "requery") {
            if (!preview.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Requery cancelled");
            preview.loadRecords();
            if (!preview.layout_.recordSource.trimmed().isEmpty() && !preview.dataSourceAvailable_)
                throw linaccess::VbaRuntimeError(preview.dataSourceErrorNumber_ ? preview.dataSourceErrorNumber_ : 3146,"Form record source unavailable after Requery");
            if (preview.opened_ && preview.initialCurrentDispatched_ && !preview.layout_.recordSource.trimmed().isEmpty()) (void)preview.formEvent("OnCurrent","Form_Current");
            return {};
        }
        if (key == "setfocus") { preview.activateWindow(); preview.setFocus(); return {}; }
        if (auto value=preview.getModuleMember(name,true)) return *value;
        return preview.getFieldValue(name);
    }
    linaccess::VbaValue item(std::string_view name) override {
        return checked().getFieldValue(name);
    }
    linaccess::VbaValue objectItem(std::string_view name) override {
        return makePreviewControlsObject(checked())->item(name);
    }
    linaccess::VbaValue objectMember(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        const auto key=text(name).toCaseFolded();
        if (!args.empty() || key=="name" || key=="caption" || key=="tag" || key=="recordsource" || key=="dirty" || key=="newrecord" || key=="controls" || key=="requery" || key=="setfocus" || key=="filter" || key=="filteron")
            return member(name,args);
        if (auto value=checked().getModuleMember(name,true)) return *value;
        return objectItem(name);
    }
    void assignItem(std::string_view name,const linaccess::VbaValue& value) override {
        PreviewControlObject(checked(),text(name)).assign("Value",value);
    }
    void assign(std::string_view name,const linaccess::VbaValue& value) override {
        auto& preview = checked(); const auto key = text(name).toCaseFolded();
        if (key=="filter" || key=="filteron") {
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null filter property");
            preview.setFilterState(key=="filter" ? text(value.asText()) : preview.filter_,key=="filteron" ? value.truthy() : preview.filterOn_); return;
        }
        if (key=="tag") {
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null Tag");
            preview.setProperty("la_access_tag",text(value.asText())); return;
        }
        if (key == "caption") {
            if (value.kind == linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null caption");
            preview.layout_.caption = text(value.asText());
            preview.setWindowTitle((preview.layout_.caption.isEmpty() ? preview.formName_ : preview.layout_.caption) + " — LinAccess");
            return;
        }
        if (preview.layout_.kind==linaccess::AccessDocumentKind::Report && key=="dirty")
            throw linaccess::VbaRuntimeError(438,"Report member unsupported");
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
        if (preview.setModuleMember(name,value,true)) return;
        PreviewControlObject(preview,text(name)).assign("Value",value);
    }
};
class PreviewFormsObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewFormsObject(AccessFormPreview& preview,bool reports=false) : PreviewObjectBase(preview),reports_(reports) {}
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
        throw linaccess::VbaRuntimeError(reports_ ? 2451 : 2450,"Requested document is not open");
    }
    linaccess::VbaValue item(std::string_view name) override { return member("Item",{linaccess::VbaValue::string(std::string(name))}); }
    std::vector<linaccess::VbaValue> enumerate() override {
        std::vector<linaccess::VbaValue> values;
        for (auto* form : openForms()) values.push_back(linaccess::VbaValue::reference(makePreviewFormObject(*form)));
        return values;
    }
private:
    bool reports_{};
    QVector<AccessFormPreview*> openForms() const {
        const auto& owner = checked(); QVector<AccessFormPreview*> forms;
        for (auto* widget : QApplication::topLevelWidgets())
            if (auto* form = dynamic_cast<AccessFormPreview*>(widget); form && !form->closed_ && !form->openCancelled_
                && form->db_.connectionName()==owner.db_.connectionName()
                && (form->layout_.kind==linaccess::AccessDocumentKind::Report)==reports_) forms.append(form);
        std::sort(forms.begin(),forms.end(),[](auto* left,auto* right) {
            return left->property("la_form_open_order").toLongLong() < right->property("la_form_open_order").toLongLong();
        });
        return forms;
    }
};


class PreviewCommandControlObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    PreviewCommandControlObject(AccessFormPreview& preview,QAction* action) : PreviewObjectBase(preview),action_(action) {}
    bool alive() const { return !action_.isNull(); }
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto* action = checkedAction();
        const auto key = text(name).toCaseFolded();
        if (key=="controls") {
            if (!action->menu()) throw linaccess::VbaRuntimeError(438,"Command button has no Controls collection");
            auto controls=makePreviewCommandControlsObject(checked(),action->menu());
            return args.empty() ? linaccess::VbaValue::reference(controls) : controls->member("Item",args);
        }
        if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected command control arguments");
        if (key=="caption") return linaccess::VbaValue::string(bytes(action->text()));
        if (key=="visible") return linaccess::VbaValue::logical(action->isVisible());
        if (key=="enabled") return linaccess::VbaValue::logical(action->isEnabled());
        if (key=="type") return linaccess::VbaValue::numeric(action->menu() ? 10 : 1); // msoControlPopup / msoControlButton
        throw linaccess::VbaRuntimeError(438,"Command control member unsupported");
    }
    void assign(std::string_view name,const linaccess::VbaValue& value) override {
        auto* action = checkedAction(); const auto key = text(name).toCaseFolded();
        if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null command control property");
        if (key=="caption") { action->setText(text(value.asText())); return; }
        if (key=="visible") { action->setVisible(value.truthy()); return; }
        if (key=="enabled") { action->setEnabled(value.truthy()); return; }
        throw linaccess::VbaRuntimeError(438,"Command control property unsupported");
    }
private:
    QAction* checkedAction() const {
        (void)checked();
        if (!action_) throw linaccess::VbaRuntimeError(91,"Command control unavailable");
        return action_;
    }
    QPointer<QAction> action_;
};
class PreviewCommandControlsObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    PreviewCommandControlsObject(AccessFormPreview& preview,QWidget* bar) : PreviewObjectBase(preview),bar_(bar) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked();
        if (!bar_) throw linaccess::VbaRuntimeError(91,"Command bar unavailable");
        const auto actions = commandActions(bar_);
        const auto key = text(name).toCaseFolded();
        if (key=="count" && args.empty()) return linaccess::VbaValue::numeric(actions.size());
        if (key=="item" && args.size()==1) {
            if (args[0].kind==linaccess::VbaValue::Kind::Number) {
                const auto index = linaccess::sqlFieldValue(args[0],"INTEGER").toLongLong();
                if (index<1 || index>actions.size()) throw linaccess::VbaRuntimeError(5,"Command control index out of range");
                return linaccess::VbaValue::reference(makePreviewCommandControlObject(preview,actions[index-1]));
            }
            const auto caption = text(args[0].asText());
            for (auto* action : actions) if (action->text().compare(caption,Qt::CaseInsensitive)==0)
                return linaccess::VbaValue::reference(makePreviewCommandControlObject(preview,action));
            throw linaccess::VbaRuntimeError(5,"Command control unavailable");
        }
        throw linaccess::VbaRuntimeError(438,"Command controls member unsupported");
    }
    linaccess::VbaValue item(std::string_view name) override { return member("Item",{linaccess::VbaValue::string(std::string(name))}); }
    std::vector<linaccess::VbaValue> enumerate() override {
        auto& preview = checked(); std::vector<linaccess::VbaValue> values;
        if (!bar_) throw linaccess::VbaRuntimeError(91,"Command bar unavailable");
        for (auto* action : commandActions(bar_))
            values.push_back(linaccess::VbaValue::reference(makePreviewCommandControlObject(preview,action)));
        return values;
    }
private:
    QPointer<QWidget> bar_;
};
std::shared_ptr<linaccess::VbaObject> makePreviewCommandControlsObject(AccessFormPreview& preview,QWidget* container) {
    return std::make_shared<PreviewCommandControlsObject>(preview,container);
}
std::shared_ptr<linaccess::VbaObject> makePreviewCommandControlObject(AccessFormPreview& preview,QAction* action) {
    auto& cached = preview.commandControlObjects_[action];
    if (auto object = cached.lock(); object && std::static_pointer_cast<PreviewCommandControlObject>(object)->alive()) return object;
    auto object = std::make_shared<PreviewCommandControlObject>(preview,action);
    cached = object; return object;
}
class PreviewCommandBarObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    PreviewCommandBarObject(AccessFormPreview& preview,QWidget* bar,QString name) : PreviewObjectBase(preview),bar_(bar),name_(std::move(name)) {}
    bool alive() const { return !bar_.isNull(); }
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto* bar = toolbar();
        if (text(name).compare("Controls",Qt::CaseInsensitive)==0) {
            auto controls = std::make_shared<PreviewCommandControlsObject>(checked(),bar);
            return args.empty() ? linaccess::VbaValue::reference(controls) : controls->member("Item",args);
        }
        if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected command bar arguments");
        const auto key = text(name).toCaseFolded();
        if (key=="name") return linaccess::VbaValue::string(bytes(name_));
        if (key=="enabled") return linaccess::VbaValue::logical(bar->isEnabled());
        if (key=="visible") return linaccess::VbaValue::logical(!bar->isHidden());
        throw linaccess::VbaRuntimeError(438,"Command bar member unsupported");
    }
    void assign(std::string_view name,const linaccess::VbaValue& value) override {
        auto* bar = toolbar(); const auto key = text(name).toCaseFolded();
        if (key!="enabled" && key!="visible") throw linaccess::VbaRuntimeError(438,"Command bar property unsupported");
        if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null command bar property");
        if (key=="enabled") bar->setEnabled(value.truthy());
        else bar->setVisible(value.truthy());
    }
private:
    QWidget* toolbar() const {
        (void)checked();
        if (!bar_) throw linaccess::VbaRuntimeError(91,"Command bar unavailable");
        return bar_;
    }
    QPointer<QWidget> bar_;
    QString name_;
};
class PreviewCommandBarsObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewCommandBarsObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview=checked(); const auto bars=toolbars(); const auto key=text(name).toCaseFolded();
        if (key=="activemenubar" && args.empty()) {
            for (auto* bar : bars) if (qobject_cast<QMenuBar*>(bar))
                return linaccess::VbaValue::reference(makePreviewCommandBarObject(preview,bar));
            throw linaccess::VbaRuntimeError(5,"Active menu bar unavailable");
        }
        if (key=="count" && args.empty()) return linaccess::VbaValue::numeric(bars.size());
        QString target;
        if (key=="item" && args.size()==1) {
            if (args[0].kind==linaccess::VbaValue::Kind::Number) {
                const auto index=linaccess::sqlFieldValue(args[0],"INTEGER").toLongLong();
                if (index<1 || index>bars.size()) throw linaccess::VbaRuntimeError(5,"Command bar index out of range");
                return linaccess::VbaValue::reference(makePreviewCommandBarObject(preview,bars[index-1]));
            }
            target=text(args[0].asText());
        } else if (args.empty()) target=text(name);
        else throw linaccess::VbaRuntimeError(450,"Unexpected command bars arguments");
        for (auto* bar : bars) {
            const auto barName=bar==preview.printToolbar_ ? QString("Print Preview") : qobject_cast<QMenuBar*>(bar) ? QString("Menu Bar") : bar->windowTitle();
            if (barName.compare(target,Qt::CaseInsensitive)==0)
                return linaccess::VbaValue::reference(makePreviewCommandBarObject(preview,bar));
        }
        throw linaccess::VbaRuntimeError(5,"Command bar unavailable");
    }
    linaccess::VbaValue item(std::string_view name) override { return member("Item",{linaccess::VbaValue::string(std::string(name))}); }
    std::vector<linaccess::VbaValue> enumerate() override {
        auto& preview=checked(); std::vector<linaccess::VbaValue> values;
        for (auto* bar : toolbars()) values.push_back(linaccess::VbaValue::reference(makePreviewCommandBarObject(preview,bar)));
        return values;
    }
private:
    QList<QWidget*> toolbars() const {
        const auto& preview=checked(); QList<QWidget*> bars;
        if (preview.printToolbar_) bars.append(preview.printToolbar_);
        for (auto* parent=preview.parentWidget();parent;parent=parent->parentWidget()) {
            if (auto* window=qobject_cast<QMainWindow*>(parent)) {
                for (auto* bar : window->findChildren<QToolBar*>(QString(),Qt::FindDirectChildrenOnly))
                    if (!bars.contains(bar)) bars.append(bar);
                if (auto* menu=qobject_cast<QMenuBar*>(window->menuWidget());menu && !bars.contains(menu)) bars.append(menu);
                break;
            }
        }
        return bars;
    }
};
class PreviewScreenObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewScreenObject(AccessFormPreview& preview):PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& owner=checked(); const auto key=text(name).toCaseFolded();
        if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected Screen arguments");
        if (key=="activeform") {
            for (auto* widget=QApplication::activeWindow();widget;widget=widget->parentWidget()) {
                auto* form=dynamic_cast<AccessFormPreview*>(widget);
                if (form && !form->closed_ && !form->openCancelled_ && form->db_.connectionName()==owner.db_.connectionName()
                    && form->layout_.kind==linaccess::AccessDocumentKind::Form)
                    return linaccess::VbaValue::reference(makePreviewFormObject(*form));
            }
            throw linaccess::VbaRuntimeError(2450,"Active form unavailable");
        }
        if (key=="activecontrol") {
            for (auto* focused=QApplication::focusWidget();focused;focused=focused->parentWidget())
                for (auto* widget:QApplication::allWidgets()) {
                    auto* form=dynamic_cast<AccessFormPreview*>(widget);
                    if (!form || form->closed_ || form->openCancelled_ || form->db_.connectionName()!=owner.db_.connectionName()) continue;
                    for (auto it=form->controls_.cbegin();it!=form->controls_.cend();++it)
                        if (it.value()==focused) return linaccess::VbaValue::reference(makePreviewControlObject(*form,it.key()));
                }
            throw linaccess::VbaRuntimeError(2474,"Active control unavailable");
        }
        throw linaccess::VbaRuntimeError(438,"Screen member unsupported");
    }
};
class PreviewDoCmdObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewDoCmdObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview=checked(); const auto key=text(name).toCaseFolded();
        if (key=="openform") {
            const auto request=linaccess::parseVbaOpenFormArguments(args);
            preview.openOtherForm(text(request.name),text(request.where)); return {};
        }
        if (key!="close") throw linaccess::VbaRuntimeError(438,"DoCmd member unsupported");
        if (args.size()>3) throw linaccess::VbaRuntimeError(450,"Unexpected Close arguments");
        for (const auto& arg:args) if (arg.kind==linaccess::VbaValue::Kind::Null)
            throw linaccess::VbaRuntimeError(94,"Invalid Null Close argument");
        const auto number=[](const linaccess::VbaValue& value) {
            const auto integer=linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
            if (integer<std::numeric_limits<std::int32_t>::min() || integer>std::numeric_limits<std::int32_t>::max())
                throw linaccess::VbaRuntimeError(6,"Close argument overflow");
            return static_cast<int>(integer);
        };
        const int type=args.empty() ? -1 : number(args[0]);
        const auto objectName=args.size()>1 ? std::optional<QString>(text(args[1].asText())) : std::nullopt;
        const int save=args.size()>2 ? number(args[2]) : 0;
        preview.closeOtherDocument(type,objectName,save); return {};
    }
};
class PreviewApplicationObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewApplicationObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked(); const auto key = text(name).toCaseFolded();
        if (key=="screen" && args.empty()) {
            if (auto object=preview.screenObject_.lock()) return linaccess::VbaValue::reference(object);
            auto object=std::make_shared<PreviewScreenObject>(preview); preview.screenObject_=object; return linaccess::VbaValue::reference(object);
        }
        if (key=="docmd" && args.empty()) {
            if (auto object=preview.doCmdObject_.lock()) return linaccess::VbaValue::reference(object);
            auto object=std::make_shared<PreviewDoCmdObject>(preview); preview.doCmdObject_=object;
            return linaccess::VbaValue::reference(object);
        }
        if (key=="commandbars") {
            auto bars = std::make_shared<PreviewCommandBarsObject>(preview);
            return args.empty() ? linaccess::VbaValue::reference(bars) : bars->member("Item",args);
        }
        if (key=="forms" && args.empty()) return linaccess::VbaValue::reference(makePreviewFormsObject(preview));
        if (key=="reports" && args.empty()) return linaccess::VbaValue::reference(makePreviewReportsObject(preview));
        throw linaccess::VbaRuntimeError(438,"Application member unsupported");
    }
};
std::shared_ptr<linaccess::VbaObject> makePreviewApplicationObject(AccessFormPreview& preview) {
    if (auto object = preview.applicationObject_.lock()) return object;
    auto object = std::make_shared<PreviewApplicationObject>(preview);
    preview.applicationObject_ = object; return object;
}
std::shared_ptr<linaccess::VbaObject> makePreviewCommandBarObject(AccessFormPreview& preview,QWidget* bar) {
    auto& cached=preview.commandBarObjects_[bar];
    if (auto object=cached.lock();object && std::static_pointer_cast<PreviewCommandBarObject>(object)->alive()) return object;
    auto object=std::make_shared<PreviewCommandBarObject>(preview,bar,bar==preview.printToolbar_ ? QString("Print Preview") : qobject_cast<QMenuBar*>(bar) ? QString("Menu Bar") : bar->windowTitle());
    cached=object; return object;
}

std::shared_ptr<linaccess::VbaObject> makePreviewFormObject(AccessFormPreview& preview) {
    if (auto object = preview.formObject_.lock()) return object;
    auto object = std::make_shared<PreviewFormObject>(preview);
    preview.formObject_ = object; return object;
}
std::shared_ptr<linaccess::VbaObject> makePreviewReportsObject(AccessFormPreview& preview) {
    if (auto object = preview.reportsObject_.lock()) return object;
    auto object = std::make_shared<PreviewFormsObject>(preview,true);
    preview.reportsObject_ = object; return object;
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
    auto canonical=name;
    for (const auto& control:preview.layout_.controls) if (control.name.compare(name,Qt::CaseInsensitive)==0) { canonical=control.name; break; }
    auto object = std::make_shared<PreviewControlObject>(preview,canonical);
    cached = object; return object;
}

void connectPreviewSqlObjects(QSqlDatabase db) {
    const auto connection=db.connectionName();
    const auto database=db.databaseName();
    linaccess::setJetSqlObjectResolver(db,[connection,database](std::string_view collection,std::string_view object,std::string_view field) {
        if (!QApplication::instance() || QThread::currentThread()!=QApplication::instance()->thread())
            throw linaccess::VbaRuntimeError(3251,"SQL GUI object lookup requires its GUI thread");
        const bool report=collection=="reports";
        const auto name=text(object);
        for (auto* widget : QApplication::topLevelWidgets()) {
            auto* preview=dynamic_cast<AccessFormPreview*>(widget);
            if (!preview || preview->closed_ || preview->openCancelled_ ||
                preview->db_.connectionName()!=connection || preview->db_.databaseName()!=database ||
                (preview->layout_.kind==linaccess::AccessDocumentKind::Report)!=report ||
                preview->formName_.compare(name,Qt::CaseInsensitive)!=0) continue;
            return preview->getFieldValue(field);
        }
        throw linaccess::VbaRuntimeError(report ? 2451 : 2450,"Referenced SQL document is not open");
    });
}
