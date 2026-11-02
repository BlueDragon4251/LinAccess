#include "access_form_objects.hpp"
#include "access_list_box.hpp"
#include "jet_sql_functions.hpp"
#include "dao_database.hpp"
#include <QThread>
#include "access_form_preview.hpp"
#include "vba_qt_value.hpp"
#include <QApplication>
#include <QAbstractButton>
#include <QAction>
#include <QLabel>
#include <QColor>
#include <QPalette>
#include <QLineEdit>
#include <QComboBox>
#include <QStandardItemModel>
#include <QCalendarWidget>
#include <QToolBar>
#include <QMainWindow>
#include <QMenuBar>
#include <QMenu>
#include <QWidgetAction>
#include <QPointer>
#include <QRegularExpression>
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
namespace {
int listRowIndex(const linaccess::VbaValue& argument,int count) {
    const auto scalar=linaccess::scalarVbaValue(argument);
    if(scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null list index");
    const auto index=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
    if(index<-2147483648LL || index>2147483647LL) throw linaccess::VbaRuntimeError(6,"List index Long overflow");
    if(index<0 || index>=count) throw linaccess::VbaRuntimeError(5,"List index out of range");
    return static_cast<int>(index);
}
}
class PreviewItemsSelectedObject final:public linaccess::VbaObject,private PreviewObjectBase {
public:
    PreviewItemsSelectedObject(AccessFormPreview& preview,QString name):PreviewObjectBase(preview),name_(std::move(name)) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        const auto rows=list().selectedRows(); const auto key=text(name).toCaseFolded();
        if(key=="count") {
            if(!args.empty()) throw linaccess::VbaRuntimeError(450,"Count takes no arguments");
            return linaccess::VbaValue::numeric(rows.size());
        }
        if(key=="item") {
            if(args.size()!=1) throw linaccess::VbaRuntimeError(450,"Item requires one index");
            return linaccess::VbaValue::numeric(rows[listRowIndex(args[0],rows.size())]);
        }
        throw linaccess::VbaRuntimeError(438,"Unknown ItemsSelected member");
    }
    std::vector<linaccess::VbaValue> enumerate() override {
        std::vector<linaccess::VbaValue> values;
        for(const auto row:list().selectedRows()) values.push_back(linaccess::VbaValue::numeric(row));
        return values;
    }
    void assign(std::string_view,const linaccess::VbaValue&) override { (void)checked(); throw linaccess::VbaRuntimeError(383,"ItemsSelected is read-only"); }
    void assignIndexedMember(std::string_view,const std::vector<linaccess::VbaValue>&,const linaccess::VbaValue&) override { (void)checked(); throw linaccess::VbaRuntimeError(383,"ItemsSelected is read-only"); }
private:
    AccessListBox& list() const {
        auto* list=dynamic_cast<AccessListBox*>(checked().controls_.value(name_.toCaseFolded()).data());
        if(!list) throw linaccess::VbaRuntimeError(438,"ListBox unavailable"); return *list;
    }
    QString name_;
};
class PreviewControlObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    static QLineEdit* selectionEditor(QWidget* widget,bool duringUpdate=false) {
        auto* editor=qobject_cast<QLineEdit*>(widget);
        if (auto* combo=qobject_cast<QComboBox*>(widget);combo && combo->isEditable()) editor=combo->lineEdit();
        if (!editor) throw linaccess::VbaRuntimeError(438,"Control has no text selection");
        if (!editor->hasFocus() && !duringUpdate) throw linaccess::VbaRuntimeError(2185,"Text selection requires control focus");
        return editor;
    }
    PreviewControlObject(AccessFormPreview& preview,QString name) : PreviewObjectBase(preview),name_(std::move(name)) {}
    bool hasScalarDefault() const override { return true; }
    linaccess::VbaValue scalarValue() override { return member("Value",{}); }
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked();
        auto* widget = control();
        const auto key = text(name).toCaseFolded();
        if(key=="listrows") {
            if(!args.empty()) throw linaccess::VbaRuntimeError(450,"ListRows takes no arguments");
            return linaccess::VbaValue::numeric(listCombo()->maxVisibleItems());
        }
        if(auto* list=dynamic_cast<AccessListBox*>(widget); list && (key=="selected" || key=="itemsselected" || key=="listindex" || key=="itemdata" || key=="multiselect")) {
            if(key=="multiselect") {
                if(!args.empty()) throw linaccess::VbaRuntimeError(450,"MultiSelect takes no arguments");
                return linaccess::VbaValue::numeric(list->multiSelectMode());
            }
            if(key!="itemdata" && preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(3251,"List selection requires native Form view");
            if(!widget->property("la_rows_loaded").toBool()) preview.requeryControl(bytes(name_));
            if(key=="itemsselected") {
                auto collection=std::make_shared<PreviewItemsSelectedObject>(preview,name_);
                return args.empty() ? linaccess::VbaValue::reference(collection) : collection->member("Item",args);
            }
            if(key=="listindex") {
                if(!args.empty()) throw linaccess::VbaRuntimeError(450,"ListIndex takes no arguments");
                return linaccess::VbaValue::numeric(list->currentIndex().row());
            }
            if(args.size()!=1) throw linaccess::VbaRuntimeError(450,"List property requires one row index");
            const auto row=listRowIndex(args[0],list->rowCount());
            if(key=="selected") return linaccess::VbaValue::logical(list->rowSelected(row));
            return linaccess::qtFieldValue(list->boundData(row));
        }
        if (key=="defaultvalue") {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected DefaultValue arguments");
            return linaccess::VbaValue::string(bytes(defaultDefinition().defaultValue));
        }
        if (fontProperty(key)) {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected font property arguments");
            (void)fontDefinition();
            const auto font=widget->font();
            if (key=="fontname") return linaccess::VbaValue::string(bytes(font.family()));
            if (key=="fontsize") {
                if (font.pointSizeF()<=0) throw linaccess::VbaRuntimeError(3251,"Pixel-sized native font has no point size");
                return linaccess::VbaValue::numeric(font.pointSizeF());
            }
            if (key=="fontweight") return linaccess::VbaValue::numeric(font.weight());
            return linaccess::VbaValue::logical(key=="fontbold" ? font.weight()>=700 : key=="fontitalic" ? font.italic() : font.underline());
        }

        if(key=="listindex" || key=="itemdata") {
            auto* combo=listCombo();
            if(key=="listindex") {
                if(!args.empty()) throw linaccess::VbaRuntimeError(450,"ListIndex takes no arguments");
                if(preview.layout_.kind!=linaccess::AccessDocumentKind::Form)
                    throw linaccess::VbaRuntimeError(3251,"ListIndex requires native Form view");
                if(combo->lineEdit() && combo->lineEdit()->isModified())
                    throw linaccess::VbaRuntimeError(3251,"ListIndex during an uncommitted editable draft is not implemented");
                if(!combo->property("la_rows_loaded").toBool()) preview.requeryControl(bytes(name_));
                return linaccess::VbaValue::numeric(combo->currentIndex());
            }
            if(args.size()!=1) throw linaccess::VbaRuntimeError(450,"ItemData requires one row index");
            const auto scalar=linaccess::scalarVbaValue(args[0]);
            if(scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null ItemData index");
            const auto index=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
            if(index<-2147483648LL || index>2147483647LL) throw linaccess::VbaRuntimeError(6,"ItemData index Long overflow");
            if(!combo->property("la_rows_loaded").toBool()) preview.requeryControl(bytes(name_));
            if(index<0 || index>=combo->count()) throw linaccess::VbaRuntimeError(5,"ItemData row index out of range");
            return linaccess::qtFieldValue(combo->itemData(static_cast<int>(index),Qt::UserRole));
        }
        if ((qobject_cast<QComboBox*>(widget) || dynamic_cast<AccessListBox*>(widget)) && key=="listcount") {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected ListCount arguments");
            if (!widget->property("la_rows_loaded").toBool()) preview.requeryControl(bytes(name_));
            const int headers=widget->property("la_header_rows").toInt();
            if(auto* list=dynamic_cast<AccessListBox*>(widget)) return linaccess::VbaValue::numeric(list->rowCount()+headers);
            return linaccess::VbaValue::numeric(qobject_cast<QComboBox*>(widget)->count()+headers);
        }
        if ((qobject_cast<QComboBox*>(widget) || dynamic_cast<AccessListBox*>(widget)) && (key=="rowsource" || key=="rowsourcetype" || key=="boundcolumn" || key=="columncount" || key=="columnwidths" || key=="columnheads")) {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected combo property arguments");
            for (const auto& definition : preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
                if (key=="rowsource") return linaccess::VbaValue::string(bytes(definition.rowSource));
                if (key=="rowsourcetype") return linaccess::VbaValue::string(bytes(definition.rowSourceType));
                if (key=="boundcolumn") return linaccess::VbaValue::numeric(definition.boundColumn);
                if (key=="columncount") return linaccess::VbaValue::numeric(definition.columnCount);
                if (key=="columnwidths") return linaccess::VbaValue::string(bytes(definition.columnWidths));
                if (key=="columnheads") return linaccess::VbaValue::logical(definition.columnHeads);
            }
        }
        if (key=="text") {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected Text arguments");
            return linaccess::VbaValue::string(bytes(selectionEditor(widget,preview.updatingControl_.compare(name_,Qt::CaseInsensitive)==0)->text()));
        }
        if (key=="selstart" || key=="sellength" || key=="seltext") {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected selection property arguments");
            auto* editor=selectionEditor(widget);
            if (key=="selstart") return linaccess::VbaValue::numeric(editor->hasSelectedText() ? editor->selectionStart() : editor->cursorPosition());
            if (key=="sellength") return linaccess::VbaValue::numeric(editor->selectedText().size());
            return editor->hasSelectedText() ? linaccess::VbaValue::string(bytes(editor->selectedText())) : linaccess::VbaValue{};
        }
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
        if (key == "eventprocprefix") return linaccess::VbaValue::string(bytes(preview.controlEventPrefix(name_)));
        if (key == "tag") return linaccess::VbaValue::string(bytes(widget->property("la_access_tag").toString()));
        if (key == "caption") return linaccess::VbaValue::string(bytes(caption()));
        if (key=="backcolor" || key=="backstyle") {
            const auto& definition=alignmentDefinition();
            if (key=="backstyle") return linaccess::VbaValue::numeric(definition.backStyle);
            if (definition.backColor) return linaccess::VbaValue::numeric(*definition.backColor);
            const auto color=QApplication::palette().color(QPalette::Active,definition.type=="TextBox" ? QPalette::Base : QPalette::Window);
            return linaccess::VbaValue::numeric(color.red()|(color.green()<<8)|(color.blue()<<16));
        }
        if (key=="forecolor") {
            const auto& definition=fontDefinition();
            if (definition.foreColor) return linaccess::VbaValue::numeric(*definition.foreColor);
            const auto color=widget->palette().color(QPalette::Active,qobject_cast<QLabel*>(widget) ? QPalette::WindowText : qobject_cast<QAbstractButton*>(widget) ? QPalette::ButtonText : QPalette::Text);
            return linaccess::VbaValue::numeric(color.red()|(color.green()<<8)|(color.blue()<<16));
        }
        if (key=="textalign") {
            if (!args.empty()) throw linaccess::VbaRuntimeError(450,"TextAlign takes no arguments");
            return linaccess::VbaValue::numeric(alignmentDefinition().textAlign);
        }
        if (key == "value") { requireValue(); return preview.getFieldValue(bytes(name_)); }
        if (key=="oldvalue") {
            requireValue();
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"OldValue requires a form control");
            if(auto* list=dynamic_cast<AccessListBox*>(widget); list && list->multiSelectMode()) return {};
            for (const auto& definition:preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
                auto source=definition.controlSource.trimmed();
                if (source.isEmpty() || source.startsWith('=')) throw linaccess::VbaRuntimeError(3251,"OldValue requires a bound data field");
                if (source.startsWith('[') && source.endsWith(']')) source=source.mid(1,source.size()-2);
                for (int column=0;column<preview.dataColumns_.size();++column) if (preview.dataColumns_[column].compare(source,Qt::CaseInsensitive)==0) {
                    const auto previous=preview.property("la_after_update_old_values").toMap();
                    const auto field=preview.dataColumns_[column].toCaseFolded();
                    if (previous.contains(field)) return linaccess::qtFieldValue(previous.value(field),preview.dataTypes_.value(field));
                    const auto row=preview.currentRecordIndex_-1;
                    return row>=0 && row<preview.recordRows_.size()
                        ? linaccess::qtFieldValue(preview.recordRows_[row][column],preview.dataTypes_.value(preview.dataColumns_[column].toCaseFolded())) : linaccess::VbaValue{};
                }
                throw linaccess::VbaRuntimeError(2465,"OldValue field unavailable");
            }
        }
        if (key=="undo") {
            requireValue();
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form || (!qobject_cast<QLineEdit*>(widget) && !qobject_cast<QComboBox*>(widget)))
                throw linaccess::VbaRuntimeError(3251,"Control Undo adapter requires a form text/combo editor");
            (void)preview.undoControlEdit(name_); return {};
        }
        if (key=="inputmask") {
            if (!qobject_cast<QLineEdit*>(widget)) throw linaccess::VbaRuntimeError(3251,"InputMask adapter requires a TextBox");
            for (const auto& definition:preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) return linaccess::VbaValue::string(bytes(definition.inputMask));
        }
        if (key=="locked" || key=="tabstop" || key=="tabindex") {
            for (const auto& definition:preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
                if (key=="locked") {
                    if (definition.type!="TextBox" && !dynamic_cast<AccessListBox*>(widget)) throw linaccess::VbaRuntimeError(3251,"Locked adapter is not implemented for this control");
                    return linaccess::VbaValue::logical(definition.locked);
                }
                if (qobject_cast<QLabel*>(widget)) throw linaccess::VbaRuntimeError(438,"Control has no tab properties");
                return key=="tabstop" ? linaccess::VbaValue::logical(definition.tabStop) : linaccess::VbaValue::numeric(definition.tabIndex);
            }
        }
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
            if (combo->property("la_access_input_readonly").toBool()) throw linaccess::VbaRuntimeError(2448,"Combo source is read-only");
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
        if(key=="listrows") {
            auto* combo=listCombo();
            if(preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(3251,"Runtime ListRows requires native Form view");
            const auto scalar=linaccess::scalarVbaValue(value);
            if(scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null ListRows");
            const auto rows=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
            if(rows<-32768 || rows>32767) throw linaccess::VbaRuntimeError(6,"ListRows Integer overflow");
            if(rows<1 || rows>255) throw linaccess::VbaRuntimeError(5,"ListRows must be from 1 to 255");
            for(auto& definition:preview.layout_.controls) if(definition.name.compare(name_,Qt::CaseInsensitive)==0) definition.listRows=static_cast<int>(rows);
            combo->setMaxVisibleItems(static_cast<int>(rows)); return;
        }
        if(dynamic_cast<AccessListBox*>(widget) && (key=="multiselect" || key=="listindex" || key=="selected" || key=="itemsselected" || key=="itemdata"))
            throw linaccess::VbaRuntimeError(key=="selected" ? 450 : key=="itemsselected" || key=="itemdata" ? 383 : 3251,"ListBox selection assignment requires its supported runtime API");
        if(key=="listindex" || key=="itemdata") {
            (void)listCombo();
            throw linaccess::VbaRuntimeError(383,"ComboBox list property is read-only");
        }
        if (key=="defaultvalue") {
            auto& definition=defaultDefinition();
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null DefaultValue");
            const auto expression=scalar.asText();
            if (linaccess::vbaStringCodeUnits(expression)>255) throw linaccess::VbaRuntimeError(5,"DefaultValue expression exceeds 255 code units");
            definition.defaultValue=text(expression); return;
        }
        if (fontProperty(key)) {
            auto& definition=fontDefinition();
            if (preview.layout_.kind==linaccess::AccessDocumentKind::Report)
                throw linaccess::VbaRuntimeError(3251,"Report font assignment requires an implemented Print event context");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null font property");
            auto next=definition.font;
            if (key=="fontname") {
                const auto family=text(scalar.asText());
                if (family.isEmpty()) throw linaccess::VbaRuntimeError(5,"Invalid empty font family");
                next.name=family;
            } else if (key=="fontsize" || key=="fontweight") {
                const auto number=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
                if (number< -32768 || number>32767) throw linaccess::VbaRuntimeError(6,"Font Integer overflow");
                if (key=="fontsize" ? number<1 || number>127 : number<0 || number>1000)
                    throw linaccess::VbaRuntimeError(5,"Font property outside supported range");
                if (key=="fontsize") next.size=static_cast<int>(number); else next.weight=static_cast<int>(number);
            } else {
                const bool flag=linaccess::coerceVbaModuleValue(scalar,"boolean").boolean;
                if (key=="fontbold") next.weight=flag ? 700 : 400;
                else if (key=="fontitalic") next.italic=flag;
                else next.underline=flag;
            }
            definition.font=std::move(next); preview.applyControlFont(widget,definition); return;
        }

        if (key=="forecolor") {
            auto& definition=fontDefinition();
            if (preview.layout_.kind==linaccess::AccessDocumentKind::Report)
                throw linaccess::VbaRuntimeError(3251,"Report runtime foreground color context is not implemented");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null ForeColor");
            const auto number=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
            if (number<std::numeric_limits<qint32>::min() || number>std::numeric_limits<qint32>::max()) throw linaccess::VbaRuntimeError(6,"ForeColor Long overflow");
            (void)preview.resolveOleColor(static_cast<qint32>(number));
            definition.foreColor=static_cast<qint32>(number); preview.applyForeColor(widget,definition); return;
        }
        if (key=="backcolor" || key=="backstyle") {
            auto& definition=alignmentDefinition();
            if (preview.layout_.kind==linaccess::AccessDocumentKind::Report)
                throw linaccess::VbaRuntimeError(3251,"Report runtime background context is not implemented");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null background property");
            const auto number=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
            if (key=="backstyle") {
                if (number<0 || number>255) throw linaccess::VbaRuntimeError(6,"BackStyle Byte overflow");
                if (number>1) throw linaccess::VbaRuntimeError(5,"Invalid BackStyle setting");
                definition.backStyle=static_cast<int>(number);
            } else {
                if (number<std::numeric_limits<qint32>::min() || number>std::numeric_limits<qint32>::max()) throw linaccess::VbaRuntimeError(6,"BackColor Long overflow");
                (void)preview.resolveOleColor(static_cast<qint32>(number));
                definition.backColor=static_cast<qint32>(number);
            }
            preview.applyControlBackground(widget,definition); return;
        }
        if (key=="textalign") {
            auto& definition=alignmentDefinition();
            if (preview.layout_.kind==linaccess::AccessDocumentKind::Report)
                throw linaccess::VbaRuntimeError(3251,"Report runtime alignment context is not implemented");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null TextAlign");
            const auto number=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
            if (number<0 || number>255) throw linaccess::VbaRuntimeError(6,"TextAlign Byte overflow");
            if (number>4) throw linaccess::VbaRuntimeError(5,"Invalid TextAlign setting");
            if (number==4) throw linaccess::VbaRuntimeError(3251,"Distributed text alignment is not implemented");
            definition.textAlign=static_cast<int>(number); preview.applyTextAlignment(widget,definition); return;
        }
        if (key=="oldvalue") throw linaccess::VbaRuntimeError(383,"OldValue is read-only");
        if (key=="text") {
            preview.ensureWritableControl(bytes(name_));
            auto* editor=selectionEditor(widget,preview.updatingControl_.compare(name_,Qt::CaseInsensitive)==0);
            if (editor->isReadOnly()) throw linaccess::VbaRuntimeError(2046,"Editor is read-only");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null Text assignment");
            preview.rememberControlEdit(widget);
            editor->selectAll(); editor->insert(text(scalar.asText())); return;
        }
        if (key=="selstart" || key=="sellength" || key=="seltext") {
            if (key=="seltext") preview.ensureWritableControl(bytes(name_));
            auto* editor=selectionEditor(widget,preview.updatingControl_.compare(name_,Qt::CaseInsensitive)==0);
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null selection property");
            if (key=="seltext") {
                preview.ensureWritableControl(bytes(name_));
                if (editor->isReadOnly()) throw linaccess::VbaRuntimeError(2046,"Editor is read-only");
                editor->insert(text(value.asText())); return;
            }
            const auto position=linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
            if (position<0) throw linaccess::VbaRuntimeError(5,"Invalid text selection position");
            if (position>32767) throw linaccess::VbaRuntimeError(6,"Text selection Integer overflow");
            if (key=="selstart") editor->setCursorPosition(std::min<int>(position,editor->text().size()));
            else {
                const int start=editor->hasSelectedText() ? editor->selectionStart() : editor->cursorPosition();
                editor->setSelection(start,std::min<int>(position,editor->text().size()-start));
            }
            return;
        }
        if (key=="eventprocprefix") {
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null event prefix");
            const auto prefix=text(value.asText());
            static const QRegularExpression identifier("^[\\p{L}_][\\p{L}\\p{N}_]{0,63}$");
            if (!identifier.match(prefix).hasMatch()) throw linaccess::VbaRuntimeError(5,"Invalid event procedure prefix");
            for (auto& definition:preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
                definition.eventProcPrefix=prefix; return;
            }
            throw linaccess::VbaRuntimeError(2465,"Control unavailable");
        }
        if ((qobject_cast<QComboBox*>(widget) || dynamic_cast<AccessListBox*>(widget)) && (key=="rowsource" || key=="rowsourcetype" || key=="boundcolumn" || key=="columncount" || key=="columnwidths" || key=="columnheads")) {
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null combo property");
            for (auto& definition : preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
                const auto previous=definition;
                if (key=="rowsource") definition.rowSource=text(value.asText());
                if (key=="rowsourcetype") definition.rowSourceType=text(value.asText());
                if (key=="columnwidths") definition.columnWidths=text(value.asText());
                if(key=="columnheads") {
                    if(preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(3251,"Runtime ColumnHeads requires native Form view");
                    const auto scalar=linaccess::scalarVbaValue(value);
                    if(scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null ColumnHeads");
                    definition.columnHeads=linaccess::coerceVbaModuleValue(scalar,"boolean").boolean;
                }
                if (key=="boundcolumn" || key=="columncount") {
                    const auto number=linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
                    if (number<(key=="columncount" ? 1 : 0) || number>255) throw linaccess::VbaRuntimeError(5,"Invalid combo column configuration");
                    if (key=="columncount") definition.columnCount=static_cast<int>(number); else definition.boundColumn=static_cast<int>(number);
                }
                try { preview.requeryControl(bytes(name_)); }
                catch (...) { definition=previous; throw; }
                return;
            }
        }
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
        if (key=="visible" || key=="enabled" || key=="locked" || key=="tabstop") {
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null control state");
            for (auto& definition:preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
                if (key=="tabstop" && qobject_cast<QLabel*>(widget)) throw linaccess::VbaRuntimeError(438,"Control has no tab properties");
                if (key=="locked" && definition.type!="TextBox" && !dynamic_cast<AccessListBox*>(widget)) throw linaccess::VbaRuntimeError(3251,"Locked adapter is not implemented for this control");
                if (key=="visible") definition.visible=scalar.truthy();
                if (key=="enabled") definition.enabled=scalar.truthy();
                if (key=="locked") definition.locked=scalar.truthy();
                if (key=="tabstop") definition.tabStop=scalar.truthy();
                preview.applyControlStates(); return;
            }
        }
        if (key=="inputmask") {
            if (!qobject_cast<QLineEdit*>(widget)) throw linaccess::VbaRuntimeError(3251,"InputMask adapter requires a TextBox");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null InputMask");
            const auto mask=text(scalar.asText());
            if (!mask.isEmpty() && mask.compare("Password",Qt::CaseInsensitive)!=0) throw linaccess::VbaRuntimeError(3251,"InputMask pattern unsupported");
            for (auto& definition:preview.layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) { definition.inputMask=mask; preview.applyControlStates(); return; }
        }
        if (key=="tabindex") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form || qobject_cast<QLabel*>(widget)) throw linaccess::VbaRuntimeError(438,"Control has no writable TabIndex");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null TabIndex");
            const auto position=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
            if (position<0) throw linaccess::VbaRuntimeError(5,"Invalid TabIndex");
            if (position>32767) throw linaccess::VbaRuntimeError(6,"TabIndex Integer overflow");
            int target=-1;
            for(int index=0;index<preview.layout_.controls.size();++index) if(preview.layout_.controls[index].name.compare(name_,Qt::CaseInsensitive)==0) target=index;
            if(target<0) throw linaccess::VbaRuntimeError(2465,"Control unavailable");
            const auto section=preview.layout_.controls[target].section; QVector<int> order;
            for(int index=0;index<preview.layout_.controls.size();++index) {
                const auto& definition=preview.layout_.controls[index]; if(definition.section!=section) continue;
                if(definition.optionGroupIndex>=0 || definition.type=="OptionGroup" || definition.type=="Subform" || definition.type=="CustomControl" || definition.type=="Tab")
                    throw linaccess::VbaRuntimeError(3251,"TabIndex adapter requires a flat section");
                auto* control=preview.controls_.value(definition.name.toCaseFolded()).data();
                if(control && !qobject_cast<QLabel*>(control) && control->focusPolicy()!=Qt::NoFocus) order.append(index);
            }
            if(position>=order.size()) throw linaccess::VbaRuntimeError(5,"TabIndex exceeds section controls");
            std::stable_sort(order.begin(),order.end(),[&](int a,int b) { return preview.layout_.controls[a].tabIndex<preview.layout_.controls[b].tabIndex; });
            const auto previous=order.indexOf(target); if(previous<0) throw linaccess::VbaRuntimeError(438,"Control has no tab position");
            order.removeAt(previous); order.insert(static_cast<int>(position),target);
            for(int index=0;index<order.size();++index) preview.layout_.controls[order[index]].tabIndex=index;
            preview.applyControlStates(); return;
        }
        throw linaccess::VbaRuntimeError(438,"Control property is unavailable or read-only");
    }
    void assignIndexedMember(std::string_view name,const std::vector<linaccess::VbaValue>& args,const linaccess::VbaValue& value) override {
        auto& preview=checked(); auto* widget=control(); const auto key=text(name).toCaseFolded();
        if(key=="itemdata" && (dynamic_cast<AccessListBox*>(widget) || qobject_cast<QComboBox*>(widget))) throw linaccess::VbaRuntimeError(383,"ItemData is read-only");
        auto* list=dynamic_cast<AccessListBox*>(widget);
        if(!list || key!="selected") { linaccess::VbaObject::assignIndexedMember(name,args,value); return; }
        if(preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(3251,"Selection assignment requires native Form view");
        if(args.size()!=1) throw linaccess::VbaRuntimeError(450,"Selected requires one row index");
        if(preview.validatingControl_.compare(name_,Qt::CaseInsensitive)==0) throw linaccess::VbaRuntimeError(2115,"Selection cannot change during BeforeUpdate");
        if(!widget->property("la_rows_loaded").toBool()) preview.requeryControl(bytes(name_));
        const auto row=listRowIndex(args[0],list->rowCount()); const auto scalar=linaccess::scalarVbaValue(value);
        if(scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null Selected setting");
        const auto selected=linaccess::coerceVbaModuleValue(scalar,"boolean").boolean;
        if(!list->multiSelectMode()) {
            if(selected) preview.setFieldValue(bytes(name_),linaccess::qtFieldValue(list->boundData(row)));
            else if(list->rowSelected(row)) preview.setFieldValue(bytes(name_),{});
        }
        list->setRowSelected(row,selected); widget->setProperty("la_null",list->multiSelectMode()!=0 || list->selectedRows().isEmpty());
        preview.pendingControls_.remove(name_); preview.editOriginals_.remove(name_); preview.rememberControlEdit(widget);
    }
private:
    QComboBox* listCombo() const {
        auto& preview=checked();
        for(const auto& definition:preview.layout_.controls)
            if(definition.name.compare(name_,Qt::CaseInsensitive)==0 && definition.type=="ComboBox")
                if(auto* combo=qobject_cast<QComboBox*>(control())) return combo;
        throw linaccess::VbaRuntimeError(438,"Control has no native ComboBox list adapter");
    }
    linaccess::AccessVisualControl& defaultDefinition() const {
        if (checked().layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"DefaultValue requires a form control");
        for(auto& definition:checked().layout_.controls) if(definition.name.compare(name_,Qt::CaseInsensitive)==0) {
            if (!AccessFormPreview::supportsControlDefault(definition)) throw linaccess::VbaRuntimeError(438,"Control has no native DefaultValue adapter");
            return definition;
        }
        throw linaccess::VbaRuntimeError(2465,"Control unavailable");
    }
    static bool fontProperty(const QString& key) {
        return key=="fontname" || key=="fontsize" || key=="fontweight" || key=="fontbold" || key=="fontitalic" || key=="fontunderline";
    }
    linaccess::AccessVisualControl& alignmentDefinition() const {
        for (auto& definition:checked().layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
            if (!AccessFormPreview::supportsTextAlignment(definition)) throw linaccess::VbaRuntimeError(438,"Control has no native text adapter");
            return definition;
        }
        throw linaccess::VbaRuntimeError(2465,"Control unavailable");
    }
    linaccess::AccessVisualControl& fontDefinition() const {
        for (auto& definition:checked().layout_.controls) if (definition.name.compare(name_,Qt::CaseInsensitive)==0) {
            if (!AccessFormPreview::supportsControlFont(definition)) throw linaccess::VbaRuntimeError(438,"Control has no native font adapter");
            return definition;
        }
        throw linaccess::VbaRuntimeError(2465,"Control unavailable");
    }
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
class PreviewRecordsetObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewRecordsetObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& arguments) override {
        auto& preview=checked(); const auto key=text(name).toCaseFolded();
        auto snapshot=preview.recordsetSnapshot();
        if (key=="clone" || key=="recordcount" || key=="bookmarkable") return snapshot->member(name,arguments);
        if (key=="bookmark") {
            if (!arguments.empty()) throw linaccess::VbaRuntimeError(450,"Unexpected bookmark arguments");
            return linaccess::daoBookmarkAt(snapshot,preview.currentRecordIndex_-1);
        }
        throw linaccess::VbaRuntimeError(3251,"Direct form recordset member unsupported; use Clone for searching");
    }
    void assign(std::string_view,const linaccess::VbaValue&) override {
        (void)checked(); throw linaccess::VbaRuntimeError(3251,"Direct form recordset mutation unsupported");
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
        if(key=="visible") {
            if(preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Visible requires a native form");
            return linaccess::VbaValue::logical(!preview.isHidden());
        }
        if (key=="allowdeletions" || key=="allowadditions" || key=="allowedits" || key=="selheight") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Record selection/deletion properties require a form");
            if (key=="allowedits") return linaccess::VbaValue::logical(preview.layout_.allowEdits);
            if (key=="allowadditions") return linaccess::VbaValue::logical(preview.layout_.allowAdditions);
            return key=="allowdeletions" ? linaccess::VbaValue::logical(preview.layout_.allowDeletions) : linaccess::VbaValue::numeric(preview.recordSelected_ ? 1 : 0);
        }
        if(key=="openargs") return preview.openArgs_ ? linaccess::VbaValue::string(bytes(*preview.openArgs_)) : linaccess::VbaValue{};
        if(key=="currentrecord") {
            if(preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"CurrentRecord requires a form");
            if(!preview.dataSourceAvailable_ || preview.layout_.recordSource.trimmed().isEmpty()) throw linaccess::VbaRuntimeError(3251,"CurrentRecord requires an available bound source");
            if(preview.currentRecordIndex_>0) return linaccess::VbaValue::longInteger(preview.currentRecordIndex_);
            if(!preview.directTable_ || !preview.layout_.allowAdditions || preview.dataTruncated_) throw linaccess::VbaRuntimeError(3251,"CurrentRecord has no supported current row");
            return linaccess::VbaValue::longInteger(static_cast<std::int32_t>(preview.recordRows_.size()+1));
        }
        if (key == "recordsource") return linaccess::VbaValue::string(bytes(preview.layout_.recordSource));
        if (key=="recordset") {
            (void)preview.recordsetSnapshot();
            return linaccess::VbaValue::reference(std::make_shared<PreviewRecordsetObject>(preview));
        }
        if (key=="bookmark") return linaccess::daoBookmarkAt(preview.recordsetSnapshot(),preview.currentRecordIndex_-1);
        if(key=="orderby" || key=="orderbyon") {
            if(preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"OrderBy requires a form");
            return key=="orderby" ? linaccess::VbaValue::string(bytes(preview.orderBy_)) : linaccess::VbaValue::logical(preview.orderByOn_);
        }
        if (key=="filter") return linaccess::VbaValue::string(bytes(preview.filter_));
        if (key=="filteron") return linaccess::VbaValue::logical(preview.filterOn_);
        if (key=="timerinterval") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Timer requires a form");
            return linaccess::VbaValue::numeric(preview.layout_.timerInterval);
        }
        if (key=="cycle") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Cycle requires a form");
            return linaccess::VbaValue::numeric(preview.layout_.cycle);
        }
        if (key=="keypreview") return linaccess::VbaValue::logical(preview.layout_.keyPreview);
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
        if (key=="recalc") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Recalc requires a form");
            preview.refreshCalculatedControls(); return {};
        }
        if (key=="undo") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Undo requires a form");
            if (preview.dirty_ || preview.insertionStarted_) preview.loadRecord(); return {};
        }
        if (key == "setfocus") { preview.activateWindow(); preview.setFocus(); return {}; }
        if (auto value=preview.getModuleMember(name,true)) return *value;
        return preview.getFieldValue(bytes(preview.controlMemberName(text(name))));
    }
    linaccess::VbaValue item(std::string_view name) override {
        return checked().getFieldValue(name);
    }
    linaccess::VbaValue objectItem(std::string_view name) override {
        return makePreviewControlsObject(checked())->item(name);
    }
    linaccess::VbaValue objectMember(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        const auto key=text(name).toCaseFolded();
        if (!args.empty() || key=="name" || key=="caption" || key=="visible" || key=="tag" || key=="allowadditions" || key=="allowdeletions" || key=="allowedits" || key=="selheight" || key=="recordsource" || key=="currentrecord" || key=="openargs" || key=="dirty" || key=="newrecord" || key=="controls" || key=="requery" || key=="recalc" || key=="undo" || key=="keypreview" || key=="timerinterval" || key=="cycle" || key=="setfocus" || key=="filter" || key=="filteron" || key=="orderby" || key=="orderbyon")
            return member(name,args);
        if (auto value=checked().getModuleMember(name,true)) return *value;
        return objectItem(bytes(checked().controlMemberName(text(name))));
    }
    void assignItem(std::string_view name,const linaccess::VbaValue& value) override {
        PreviewControlObject(checked(),text(name)).assign("Value",value);
    }
    void assign(std::string_view name,const linaccess::VbaValue& value) override {
        auto& preview = checked(); const auto key = text(name).toCaseFolded();
        if(key=="visible") {
            const auto scalar=linaccess::scalarVbaValue(value);
            if(scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null Visible");
            const auto visible=linaccess::coerceVbaModuleValue(scalar,"boolean").boolean;
            checked().setDocumentVisible(visible); return;
        }
        if (key=="selheight") throw linaccess::VbaRuntimeError(438,"Changing record selection height unsupported");
        if (key=="recordset") throw linaccess::VbaRuntimeError(3251,"Replacing the form recordset is unsupported");
        if (key=="bookmark") {
            const auto target=linaccess::daoBookmarkPosition(preview.recordsetSnapshot(),value)+1;
            preview.selectRecord(static_cast<int>(target));
            if (preview.currentRecordIndex_!=target) throw linaccess::VbaRuntimeError(2501,"Bookmark navigation cancelled");
            return;
        }
        if (key=="allowedits") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"AllowEdits requires a form");
            if (preview.recordCommandRunning_) throw linaccess::VbaRuntimeError(3251,"Changing edit policy during record command unsupported");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null AllowEdits");
            const auto enabled=linaccess::coerceVbaModuleValue(scalar,"boolean").boolean;
            preview.layout_.allowEdits=enabled; preview.applyControlStates(); return;
        }
        if (key=="allowadditions") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"AllowAdditions requires a form");
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null AllowAdditions");
            if (preview.recordCommandRunning_) throw linaccess::VbaRuntimeError(3251,"Changing addition policy during record command unsupported");
            if (!value.truthy() && preview.currentRecordIndex_==0 && preview.dirty_ && !preview.saveRecord()) throw linaccess::VbaRuntimeError(2501,"Addition policy save cancelled");
            preview.layout_.allowAdditions=value.truthy();
            if (auto* model=qobject_cast<QStandardItemModel*>(preview.recordSelector_->model());model && model->item(0)) model->item(0)->setEnabled(preview.layout_.allowAdditions);
            if (preview.currentRecordIndex_==0) preview.loadRecord();
            return;
        }
        if (key=="allowdeletions") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"AllowDeletions requires a form");
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null AllowDeletions");
            if (preview.recordCommandRunning_) throw linaccess::VbaRuntimeError(3251,"Changing deletion policy during record command unsupported");
            preview.layout_.allowDeletions=value.truthy(); return;
        }
        if (key=="cycle") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Cycle requires a form");
            const auto scalar=linaccess::scalarVbaValue(value);
            if (scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null Cycle");
            const auto mode=linaccess::sqlFieldValue(scalar,"INTEGER").toLongLong();
            if (mode<0 || mode>2) throw linaccess::VbaRuntimeError(5,"Invalid Cycle setting");
            if (mode==2 || !preview.flatTabScope()) throw linaccess::VbaRuntimeError(3251,"Cycle adapter requires a flat form and record mode");
            preview.layout_.cycle=static_cast<int>(mode); return;
        }
        if (key=="timerinterval") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"Timer requires a form");
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null TimerInterval");
            const auto interval=linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
            if (interval>std::numeric_limits<int>::max()) throw linaccess::VbaRuntimeError(6,"TimerInterval overflow");
            if (interval<0) throw linaccess::VbaRuntimeError(5,"Invalid TimerInterval");
            preview.setTimerInterval(static_cast<int>(interval)); return;
        }
        if (key=="keypreview") {
            if (preview.layout_.kind!=linaccess::AccessDocumentKind::Form) throw linaccess::VbaRuntimeError(438,"KeyPreview requires a form");
            if (value.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null KeyPreview");
            preview.layout_.keyPreview=value.truthy(); return;
        }
        if(key=="orderby" || key=="orderbyon") {
            const auto scalar=linaccess::scalarVbaValue(value);
            if(scalar.kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null sort property");
            const auto order=key=="orderby" ? text(scalar.asText()) : preview.orderBy_;
            const auto enabled=key=="orderbyon" ? linaccess::coerceVbaModuleValue(scalar,"boolean").boolean : preview.orderByOn_;
            preview.setOrderState(order,enabled); return;
        }
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
        if (key == "name" || key == "recordsource" || key == "newrecord" || key == "currentrecord" || key=="openargs")
            throw linaccess::VbaRuntimeError(438,"Form property is read-only");
        if (preview.setModuleMember(name,value,true)) return;
        PreviewControlObject(preview,preview.controlMemberName(text(name))).assign("Value",value);
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
            if(AccessFormPreview::databaseWindowActive()) throw linaccess::VbaRuntimeError(2450,"Active form unavailable in database window");
            auto* form=activePreviewDocument();
            if (!form) for (auto* widget=QApplication::activeWindow();widget;widget=widget->parentWidget())
                if (auto* candidate=dynamic_cast<AccessFormPreview*>(widget)) { form=candidate; break; }
            if (form && !form->closed_ && !form->openCancelled_ && form->db_.connectionName()==owner.db_.connectionName()
                && form->db_.databaseName()==owner.db_.databaseName() && form->layout_.kind==linaccess::AccessDocumentKind::Form)
                return linaccess::VbaValue::reference(makePreviewFormObject(*form));
            throw linaccess::VbaRuntimeError(2450,"Active form unavailable");
        }
        if (key=="activecontrol") {
            if(AccessFormPreview::databaseWindowActive()) throw linaccess::VbaRuntimeError(2474,"Active control unavailable in database window");
            QWidget* focus=QApplication::focusWidget();
            if (auto* document=activePreviewDocument();document && !document->closed_ && !document->openCancelled_
                && document->db_.connectionName()==owner.db_.connectionName() && document->db_.databaseName()==owner.db_.databaseName())
                focus=document->focusWidget();
            for (auto* focused=focus;focused;focused=focused->parentWidget())
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
    std::optional<linaccess::VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=text(name).toCaseFolded();
        using Signature=linaccess::VbaArgumentSignature;
        if(key=="openform") return Signature{{"FormName","View","FilterName","WhereCondition","DataMode","WindowMode","OpenArgs"},1};
        if(key=="openreport") return Signature{{"ReportName","View","FilterName","WhereCondition","WindowMode","OpenArgs"},1};
        if(key=="close") return Signature{{"ObjectType","ObjectName","Save"},0};
        if(key=="gotorecord") return Signature{{"ObjectType","ObjectName","Record","Offset"},0};
        if(key=="runsql") return Signature{{"SQLStatement","UseTransaction"},1};
        if(key=="runcommand") return Signature{{"Command"},1};
        if(key=="gotocontrol") return Signature{{"ControlName"},1};
        if(key=="requery") return Signature{{"ControlName"},0};
        if(key=="selectobject") return Signature{{"ObjectType","ObjectName","InDatabaseWindow"},1};
        if(key=="domenuitem") return Signature{{"MenuBar","MenuName","Command","Subcommand","Version"},3};
        return {};
    }
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview=checked(); const auto key=text(name).toCaseFolded();
        if (key=="domenuitem" || key=="runcommand" || key=="gotorecord" || key=="requery" || key=="maximize" || key=="runsql" || key=="selectobject")
            return dispatchPreviewDocumentCommand(preview,key.toStdString(),args);
        if (key=="gotocontrol") {
            if (args.size()!=1) throw linaccess::VbaRuntimeError(450,"GoToControl requires a control name");
            if(preview.hasActivated_ && AccessFormPreview::databaseWindowActive()) throw linaccess::VbaRuntimeError(2046,"GoToControl requires an active native document");
            if (args[0].kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null control name");
            auto requested=text(args[0].asText());
            if (requested.startsWith('[') && requested.endsWith(']')) { requested=requested.mid(1,requested.size()-2); requested.replace("]]","]"); }
            auto* active=activePreviewDocument();
            auto& target=preview.opened_ && active && !active->closed_ && !active->openCancelled_
                && active->db_.connectionName()==preview.db_.connectionName() && active->db_.databaseName()==preview.db_.databaseName() ? *active : preview;
            makePreviewFormObject(target)->member("Controls",{linaccess::VbaValue::string(bytes(requested))}).object->member("SetFocus",{});
            auto* widget=target.controls_.value(requested.toCaseFolded()).data();
            if (auto* editor=qobject_cast<QLineEdit*>(widget)) editor->selectAll();
            else if (auto* combo=qobject_cast<QComboBox*>(widget);combo && combo->isEditable()) combo->lineEdit()->selectAll();
            return {};
        }
        if (key=="quit") {
            if (!args.empty()) throw linaccess::VbaRuntimeError(3251,"Quit options unsupported");
            AccessFormPreview::quitApplication(); return {};
        }
        if (key=="openreport") {
            if (args.empty() || args.size()>6) throw linaccess::VbaRuntimeError(450,"OpenReport argument count invalid");
            std::vector<linaccess::VbaValue> values;
            for (const auto& arg:args) {
                auto value=linaccess::scalarVbaValue(arg);
                if (value.kind==linaccess::VbaValue::Kind::Null && values.size()!=5) throw linaccess::VbaRuntimeError(94,"Invalid Null OpenReport argument");
                values.push_back(std::move(value));
            }
            const auto number=[](const linaccess::VbaValue& value) {
                const auto integer=linaccess::sqlFieldValue(value,"INTEGER").toLongLong();
                if (integer<std::numeric_limits<int>::min() || integer>std::numeric_limits<int>::max()) throw linaccess::VbaRuntimeError(6,"OpenReport argument overflow");
                return static_cast<int>(integer);
            };
            const auto reportName=text(values[0].asText());
            if (reportName.isEmpty()) throw linaccess::VbaRuntimeError(2103,"Report name unavailable");
            const int view=values.size()>1 && values[1].kind!=linaccess::VbaValue::Kind::Empty ? number(values[1]) : 0;
            if (view!=2) throw linaccess::VbaRuntimeError(3251,"OpenReport supports native preview only");
            if (values.size()>2 && !values[2].asText().empty()) throw linaccess::VbaRuntimeError(3251,"Named report filters unsupported");
            const auto where=values.size()>3 ? text(values[3].asText()) : QString{};
            if (where.size()>32768) throw linaccess::VbaRuntimeError(5,"Report WhereCondition too long");
            if (values.size()>4 && values[4].kind!=linaccess::VbaValue::Kind::Empty && number(values[4])!=0) throw linaccess::VbaRuntimeError(3251,"Report window mode unsupported");
            std::optional<QString> openArgs;
            if(values.size()>5 && values[5].kind!=linaccess::VbaValue::Kind::Empty && values[5].kind!=linaccess::VbaValue::Kind::Null) {
                if(values[5].kind!=linaccess::VbaValue::Kind::Text) throw linaccess::VbaRuntimeError(3251,"Report OpenArgs currently requires text or Null");
                (void)linaccess::vbaStringCodeUnits(values[5].text); openArgs=text(values[5].text);
            }
            preview.openOtherDocument(linaccess::AccessDocumentKind::Report,reportName,where,std::move(openArgs)); return {};
        }
        if (key=="openform") {
            const auto request=linaccess::parseVbaOpenFormArguments(args);
            preview.openOtherForm(text(request.name),text(request.where),request.openArgs ? std::optional<QString>(text(*request.openArgs)) : std::nullopt); return {};
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
        const int type=args.empty() || args[0].kind==linaccess::VbaValue::Kind::Empty ? -1 : number(args[0]);
        const auto objectName=args.size()>1 && args[1].kind!=linaccess::VbaValue::Kind::Empty ? std::optional<QString>(text(args[1].asText())) : std::nullopt;
        const int save=args.size()>2 ? number(args[2]) : 0;
        preview.closeOtherDocument(type,objectName,save); return {};
    }
};
class PreviewApplicationObject final : public linaccess::VbaObject, private PreviewObjectBase {
public:
    explicit PreviewApplicationObject(AccessFormPreview& preview) : PreviewObjectBase(preview) {}
    linaccess::VbaValue member(std::string_view name,const std::vector<linaccess::VbaValue>& args) override {
        auto& preview = checked(); const auto key = text(name).toCaseFolded();
        if (key=="eval") {
            if (args.size()!=1) throw linaccess::VbaRuntimeError(450,"Eval requires one argument");
            if (args[0].kind==linaccess::VbaValue::Kind::Null) throw linaccess::VbaRuntimeError(94,"Invalid Null Eval argument");
            return preview.evalExpression(args[0].asText());
        }
        if (key=="dbengine" && args.empty()) return linaccess::VbaValue::reference(linaccess::makeDaoEngine(preview.db_));
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
