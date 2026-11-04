#pragma once
#include "vba_qt_value.hpp"
#include <QFileDialog>
#include <QFileInfo>
#include <QPointer>
#include <QScopeGuard>
#include <functional>

namespace linaccess::native_picker {
inline QString key(std::string_view value) { return QString::fromUtf8(value.data(),value.size()).toCaseFolded(); }
inline std::int32_t integer(const VbaValue& value) { return sqlFieldValue(scalarVbaValue(value),"INTEGER4").toInt(); }
inline QString string(const VbaValue& value) {
    const auto scalar=scalarVbaValue(value);
    if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null dialog property");
    const auto result=QString::fromStdString(scalar.asText());
    if(result.contains(QChar(0)) || result.contains('\n') || result.contains('\r')) throw VbaRuntimeError(5,"Invalid dialog property text");
    return result;
}
struct Filter { QString description,extensions; };
struct State {
    QPointer<QWidget> parent;
    std::function<void()> guard;
    int type{3},filterIndex{1}; bool multi{},showing{};
    QString title,initial,button;
    QVector<std::shared_ptr<Filter>> filters;
    QStringList selected;
};
inline int index(const VbaValue& value,qsizetype count) {
    const auto position=integer(value);
    if(position<1 || position>count) throw VbaRuntimeError(5,"Dialog collection index out of range");
    return position-1;
}
class FilterObject final : public VbaObject {
public:
    FilterObject(std::shared_ptr<State> state,std::shared_ptr<Filter> filter) : state_(std::move(state)),filter_(std::move(filter)) {}
    std::string_view className() const override { return "FileDialogFilter"; }
    bool isNativeType(std::string_view type) const override { return type=="office.filedialogfilter"; }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->guard(); if(!args.empty()) throw VbaRuntimeError(450,"Filter property requires no arguments");
        const auto property=key(name);
        if(property=="description") return VbaValue::string(filter_->description.toStdString());
        if(property=="extensions") return VbaValue::string(filter_->extensions.toStdString());
        throw VbaRuntimeError(438,"FileDialogFilter member unsupported");
    }
    void assign(std::string_view,const VbaValue&) override { state_->guard(); throw VbaRuntimeError(383,"Filter is read-only"); }
private:
    std::shared_ptr<State> state_; std::shared_ptr<Filter> filter_;
};
class Filters final : public VbaObject {
public:
    explicit Filters(std::shared_ptr<State> state) : state_(std::move(state)) {}
    std::string_view className() const override { return "FileDialogFilters"; }
    bool isNativeType(std::string_view type) const override { return type=="office.filedialogfilters"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(key(name)=="add") return VbaArgumentSignature{{"Description","Extensions","Position"},2};
        if(key(name)=="item") return VbaArgumentSignature{{"Index"},1};
        return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->guard(); const auto property=key(name);
        if(state_->showing && (property=="add" || property=="clear")) throw VbaRuntimeError(3251,"File dialog is active");
        if(property=="count" && args.empty()) return VbaValue::longInteger(static_cast<std::int32_t>(state_->filters.size()));
        if(property=="clear" && args.empty()) { state_->filters.clear(); state_->filterIndex=1; return {}; }
        if(property=="item" && args.size()==1)
            return VbaValue::reference(std::make_shared<FilterObject>(state_,state_->filters[index(args[0],state_->filters.size())]));
        if(property=="add") {
            if(args.size()<2 || args.size()>3) throw VbaRuntimeError(450,"Filters.Add requires two or three arguments");
            auto filter=std::make_shared<Filter>(Filter{string(args[0]),string(args[1])});
            if(filter->description.contains('(') || filter->description.contains(')')) throw VbaRuntimeError(3251,"Filter description delimiters unsupported");
            const auto patterns=filter->extensions.split(';',Qt::SkipEmptyParts);
            if(patterns.isEmpty()) throw VbaRuntimeError(5,"Filter extensions are empty");
            for(const auto& raw:patterns) {
                const auto pattern=raw.trimmed();
                if((pattern!="*" && !pattern.startsWith("*.")) || pattern.contains('/') || pattern.contains('\\')
                    || pattern.contains('(') || pattern.contains(')') || pattern.contains(' '))
                    throw VbaRuntimeError(3251,"Filter pattern unsupported");
            }
            int position=state_->filters.size();
            if(args.size()==3 && args[2].kind!=VbaValue::Kind::Missing) position=index(args[2],state_->filters.size()+1);
            state_->filters.insert(position,std::move(filter)); return {};
        }
        throw VbaRuntimeError(438,"FileDialogFilters member or arguments unsupported");
    }
    std::vector<VbaValue> enumerate() override {
        state_->guard();
        std::vector<VbaValue> result; result.reserve(state_->filters.size());
        for (const auto& filter:state_->filters)
            result.push_back(VbaValue::reference(std::make_shared<FilterObject>(state_,filter)));
        return result;
    }
    void assign(std::string_view,const VbaValue&) override { state_->guard(); throw VbaRuntimeError(383,"Filters properties are read-only"); }
private: std::shared_ptr<State> state_;
};
class SelectedItems final : public VbaObject {
public:
    explicit SelectedItems(std::shared_ptr<State> state) : state_(std::move(state)) {}
    std::string_view className() const override { return "FileDialogSelectedItems"; }
    bool isNativeType(std::string_view type) const override { return type=="office.filedialogselecteditems"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(key(name)=="item") return VbaArgumentSignature{{"Index"},1}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->guard(); const auto property=key(name);
        if(property=="count" && args.empty()) return VbaValue::longInteger(static_cast<std::int32_t>(state_->selected.size()));
        if(property=="item" && args.size()==1) return VbaValue::string(state_->selected[index(args[0],state_->selected.size())].toStdString());
        throw VbaRuntimeError(438,"FileDialogSelectedItems member or arguments unsupported");
    }
    std::vector<VbaValue> enumerate() override {
        state_->guard(); std::vector<VbaValue> result;
        for(const auto& path:state_->selected) result.push_back(VbaValue::string(path.toStdString())); return result;
    }
    void assign(std::string_view,const VbaValue&) override { state_->guard(); throw VbaRuntimeError(383,"SelectedItems is read-only"); }
private: std::shared_ptr<State> state_;
};
class Picker final : public VbaObject {
public:
    Picker(QWidget* parent,std::function<void()> guard) : state_(std::make_shared<State>()) { state_->parent=parent; state_->guard=std::move(guard); }
    std::string_view className() const override { return "FileDialog"; }
    bool isNativeType(std::string_view type) const override { return type=="office.filedialog"; }
    void setType(int type) {
        state_->guard();
        if(type!=3 && type!=4) throw VbaRuntimeError(3251,"Access FileDialog supports FilePicker and FolderPicker only");
        if(state_->showing) throw VbaRuntimeError(3251,"File dialog is already active");
        state_->type=type; if(type==4) state_->multi=false;
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->guard(); const auto property=key(name);
        if(property=="filters" || property=="selecteditems") {
            auto& weak=property=="filters" ? filters_ : selected_; auto object=weak.lock();
            if(!object) {
                object=property=="filters" ? std::static_pointer_cast<VbaObject>(std::make_shared<Filters>(state_))
                    : std::static_pointer_cast<VbaObject>(std::make_shared<SelectedItems>(state_)); weak=object;
            }
            return args.empty() ? VbaValue::reference(object) : object->member("Item",args);
        }
        if(!args.empty()) throw VbaRuntimeError(450,"FileDialog property requires no arguments");
        if(property=="dialogtype") return VbaValue::longInteger(state_->type);
        if(property=="allowmultiselect") return VbaValue::logical(state_->multi);
        if(property=="title") return VbaValue::string(state_->title.toStdString());
        if(property=="initialfilename") return VbaValue::string(state_->initial.toStdString());
        if(property=="buttonname") return VbaValue::string(state_->button.toStdString());
        if(property=="filterindex") return VbaValue::longInteger(state_->filterIndex);
        if(property!="show") throw VbaRuntimeError(438,"FileDialog member unsupported");
        if(state_->showing) throw VbaRuntimeError(3251,"File dialog is already active");
        state_->showing=true; const auto reset=qScopeGuard([&] { state_->showing=false; });
        QFileDialog dialog(state_->parent,state_->title); dialog.setObjectName("access-file-dialog");
        dialog.setAcceptMode(QFileDialog::AcceptOpen);
        dialog.setFileMode(state_->type==4 ? QFileDialog::Directory : state_->multi ? QFileDialog::ExistingFiles : QFileDialog::ExistingFile);
        if(state_->type==4) dialog.setOption(QFileDialog::ShowDirsOnly);
        if(!state_->initial.isEmpty()) {
            const QFileInfo initial(state_->initial);
            if(initial.isDir()) dialog.setDirectory(initial.absoluteFilePath());
            else { dialog.setDirectory(initial.absolutePath()); dialog.selectFile(initial.fileName()); }
        }
        if(!state_->button.isEmpty()) dialog.setLabelText(QFileDialog::Accept,state_->button);
        QStringList filters;
        for(const auto& filter:state_->filters) {
            QStringList patterns;
            for(const auto& raw:filter->extensions.split(';',Qt::SkipEmptyParts)) {
                auto pattern=raw.trimmed(); if(pattern=="*.*") pattern="*"; patterns.append(pattern);
            }
            filters.append(filter->description+" ("+patterns.join(' ')+')');
        }
        if(!filters.isEmpty()) {
            dialog.setNameFilters(filters);
            dialog.selectNameFilter(filters[qBound(0,state_->filterIndex-1,int(filters.size())-1)]);
        }
        state_->selected.clear(); const auto result=dialog.exec(); state_->guard();
        if(result!=QDialog::Accepted) return VbaValue::longInteger(0);
        state_->selected=dialog.selectedFiles();
        if(state_->selected.isEmpty()) throw VbaRuntimeError(3251,"Accepted dialog has no selected path");
        if(!filters.isEmpty()) state_->filterIndex=filters.indexOf(dialog.selectedNameFilter())+1;
        return VbaValue::longInteger(-1);
    }
    void assign(std::string_view name,const VbaValue& value) override {
        state_->guard(); if(state_->showing) throw VbaRuntimeError(3251,"File dialog is active");
        const auto property=key(name);
        if(property=="allowmultiselect") {
            const auto scalar=scalarVbaValue(value);
            if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null selection mode");
            const bool multi=scalar.truthy();
            if(multi && state_->type==4) throw VbaRuntimeError(3251,"Multiple folder selection unsupported");
            state_->multi=multi; return;
        }
        if(property=="title") { state_->title=string(value); return; }
        if(property=="initialfilename") { state_->initial=string(value); return; }
        if(property=="buttonname") { state_->button=string(value); return; }
        if(property=="filterindex") {
            const int number=integer(value);
            if(number<1 || number>qMax(1,int(state_->filters.size()))) throw VbaRuntimeError(5,"FilterIndex out of range");
            state_->filterIndex=number; return;
        }
        if(property=="dialogtype" || property=="filters" || property=="selecteditems") throw VbaRuntimeError(383,"FileDialog property is read-only");
        throw VbaRuntimeError(438,"FileDialog property unsupported");
    }
private:
    std::shared_ptr<State> state_; std::weak_ptr<VbaObject> filters_,selected_;
};
inline std::shared_ptr<VbaObject> create(QWidget* parent,int type,std::function<void()> guard) {
    if(type!=3 && type!=4) throw VbaRuntimeError(3251,"Access FileDialog supports FilePicker and FolderPicker only");
    auto object=parent->property("la_native_file_dialog").value<std::shared_ptr<VbaObject>>();
    if(!object) { object=std::make_shared<Picker>(parent,std::move(guard)); parent->setProperty("la_native_file_dialog",QVariant::fromValue(object)); }
    auto picker=std::dynamic_pointer_cast<Picker>(object);
    if(!picker) throw VbaRuntimeError(3251,"Native file dialog context invalid");
    picker->setType(type); return object;
}
}
