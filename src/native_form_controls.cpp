#include "native_form_controls.hpp"
#include "native_form_events.hpp"
#include <QDialog>
#include <QDialogButtonBox>
#include <QComboBox>
#include <QTableWidget>
#include <QHeaderView>
#include <QPushButton>
#include <QLabel>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QRegularExpression>
#include <cmath>
namespace linaccess {
namespace {
const QStringList types{"TextBox","Label","CommandButton","CheckBox","ComboBox","ListBox"};
const QStringList keys{"type","name","caption","control_source","left","top","width","height","row_source_type","row_source","onclick"};
QString quoted(const QString& value) {
    if(value.size()>65536 || !value.isValidUtf16() || value.contains(QChar(0)) || value.contains('\r') || value.contains('\n')) throw VbaRuntimeError(5,"Invalid native control property text");
    auto text=value;text.replace('"',"\"\"");return '"'+text+'"';
}
QString eventPrefix(QString name) {for(auto& c:name) if(!c.isLetterOrNumber() && c!='_') c='_';if(name.isEmpty() || name.front().isDigit()) name.prepend('_');return name.toCaseFolded();}
}
NativeFormControls nativeFormControls(const QJsonValue& definition,const QStringList& fields,const QStringList& schema,const QSet<QString>& procedures) {
    NativeFormControls result;
    if(fields.size()>255) throw VbaRuntimeError(5,"Native form accepts at most 255 bound fields");
    if(definition.isUndefined()) return result;
    if(!definition.isArray() || definition.toArray().size()>128) throw VbaRuntimeError(5,"Native form accepts at most 128 additional controls");
    QSet<QString> used,prefixes;
    for(const auto& name:fields) {used.insert(name.toCaseFolded());prefixes.insert(eventPrefix(name));}
    for(const auto* name:{"btnNeu","btnSpeichern","btnAktion"}) {used.insert(QString::fromLatin1(name).toCaseFolded());prefixes.insert(QString::fromLatin1(name).toCaseFolded());}
    int ordinal=fields.size();
    for(const auto& value:definition.toArray()) {
        if(!value.isObject()) throw VbaRuntimeError(5,"Native control requires an object");const auto object=value.toObject();
        for(auto entry=object.begin();entry!=object.end();++entry) if(!keys.contains(entry.key())) throw VbaRuntimeError(3251,"Unsupported native control property");
        const auto string=[&](const QString& key,const QString& fallback=QString{}) {
            const auto value=object.value(key);if(!value.isUndefined() && !value.isString()) throw VbaRuntimeError(13,"Native control property requires text");
            const auto text=value.isUndefined() ? fallback : value.toString();(void)quoted(text);return text;
        };
        const auto type=string("type"),name=string("name");
        static const QRegularExpression identifier("^[\\p{L}_][\\p{L}\\p{N}_]*$");
        if(!types.contains(type) || name.size()>64 || !identifier.match(name).hasMatch()) throw VbaRuntimeError(5,"Native control requires a supported type and identifier name");
        if(used.contains(name.toCaseFolded()) || prefixes.contains(eventPrefix(name))) throw VbaRuntimeError(5,"Duplicate native control name or event prefix");
        used.insert(name.toCaseFolded());prefixes.insert(eventPrefix(name));result.names.append(name);
        const auto source=string("control_source"),caption=string("caption"),rows=string("row_source"),rowType=string("row_source_type","Table/Query");
        const bool list=type=="ComboBox" || type=="ListBox";
        if((type=="Label" || type=="CommandButton") && !source.isEmpty()) throw VbaRuntimeError(5,"This control has no ControlSource");
        if((type=="TextBox" || list) && !caption.isEmpty()) throw VbaRuntimeError(5,"Use a separate Label for this control caption");
        if(!list && !rows.isEmpty()) throw VbaRuntimeError(5,"Only list controls have a RowSource");
        if(list && rowType!="Table/Query" && rowType!="Value List") throw VbaRuntimeError(3251,"Unsupported native row source type");
        if(!source.isEmpty() && !source.startsWith('=')) {
            auto field=source;if(field.startsWith('[') && field.endsWith(']')) field=field.mid(1,field.size()-2).replace("]]","]");
            int matches=0;for(const auto& column:schema) matches+=column.compare(field,Qt::CaseInsensitive)==0;
            if(matches!=1) throw VbaRuntimeError(matches ? 3251 : 3265,"ControlSource field unavailable or ambiguous");
        }
        const auto integer=[&](const QString& key,int fallback,int minimum) {
            const auto value=object.value(key);if(value.isUndefined()) return fallback;const auto number=value.toDouble(-1);
            if(!value.isDouble() || !std::isfinite(number) || number!=std::floor(number) || number<minimum || number>30000) throw VbaRuntimeError(5,"Invalid native control twip geometry");return static_cast<int>(number);
        };
        const int left=integer("left",240,0),top=integer("top",240+ordinal*420,0),width=integer("width",2200,1),height=integer("height",300,1);
        if(left+width>30000 || top+height>30000) throw VbaRuntimeError(5,"Native control extends beyond supported geometry");
        result.width=qMax(result.width,left+width+240);result.bottom=qMax(result.bottom,top+height+240);
        auto& lines=result.lines;lines<<"        Begin "+type<<"            Name = "+quoted(name)<<"            EventProcPrefix = "+quoted(name)
            <<"            Left = "+QString::number(left)<<"            Top = "+QString::number(top)<<"            Width = "+QString::number(width)<<"            Height = "+QString::number(height)<<"            TabIndex = "+QString::number(ordinal++);
        if(!caption.isEmpty()) lines<<"            Caption = "+quoted(caption);
        if(!source.isEmpty()) lines<<"            ControlSource = "+quoted(source);
        if(list) lines<<"            RowSourceType = "+quoted(rowType)<<"            RowSource = "+quoted(rows);
        const auto click=string("onclick");QMap<QString,QString> overrides;
        if(!click.isEmpty()) overrides=nativeEventOverrides(QJsonObject{{"OnClick",click}},nativeFieldMacroEvents());
        for(const auto* event:{"OnClick","OnChange","BeforeUpdate","AfterUpdate","OnDblClick","OnKeyDown","OnKeyPress","OnKeyUp","OnMouseDown","OnMouseUp"}) {
            const QString property=QString::fromLatin1(event);const auto suffix=property.startsWith("On") ? property.mid(2) : property;
            if(overrides.contains(property)) lines<<"            "+property+" = "+quoted(overrides.value(property));
            else if(procedures.contains((name+"_"+suffix).toCaseFolded())) lines<<"            "+property+" = \"[Event Procedure]\"";
        }
        lines<<"        End";
    }
    return result;
}
bool editNativeFormControls(QWidget* parent,QJsonArray& controls,const QStringList& fields,const QStringList& schema) {
    (void)nativeFormControls(controls,fields,schema,{});
    QDialog dialog(parent);dialog.setObjectName("native-form-controls-editor");dialog.setWindowTitle("Steuerelemente");dialog.resize(1020,540);
    auto* layout=new QVBoxLayout(&dialog);auto* help=new QLabel("Eigene Steuerelemente anlegen. Position und Größe in Twips. Beschriftungen für Label, Schaltfläche und Checkbox; andere Eingabefelder erhalten bei Bedarf ein separates Label. Leeres Klickereignis verwendet eine vorhandene VBA-Prozedur automatisch.",&dialog);help->setWordWrap(true);layout->addWidget(help);
    auto* grid=new QTableWidget(0,keys.size(),&dialog);grid->setObjectName("native-form-controls-grid");grid->setHorizontalHeaderLabels({"Typ","Name","Beschriftung","Steuerelementinhalt","Links","Oben","Breite","Höhe","Herkunftsart","Datensatzherkunft","Beim Klicken"});grid->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);layout->addWidget(grid);
    const auto append=[&](const QJsonObject& object) {
        const int row=grid->rowCount();grid->insertRow(row);
        for(int column=0;column<keys.size();++column) {
            if(column==0 || column==8) {auto* combo=new QComboBox(grid);combo->addItems(column==0 ? types : QStringList{"Table/Query","Value List"});combo->setCurrentText(object.value(keys[column]).toString(column==0 ? "TextBox" : "Table/Query"));grid->setCellWidget(row,column,combo);}
            else {const int defaults[]{240,240+(static_cast<int>(fields.size())+row)*420,2200,300};const auto value=object.contains(keys[column]) ? object.value(keys[column]) : column>=4 && column<=7 ? QJsonValue(defaults[column-4]) : QJsonValue(QString{});auto* item=new QTableWidgetItem(value.isDouble() ? QString::number(value.toInt()) : value.toString());grid->setItem(row,column,item);}
        }
    };
    for(const auto& value:controls) append(value.toObject());
    auto* buttons=new QHBoxLayout;auto* add=new QPushButton("Steuerelement hinzufügen",&dialog);add->setObjectName("native-form-controls-add");buttons->addWidget(add);auto* remove=new QPushButton("Ausgewähltes entfernen",&dialog);remove->setObjectName("native-form-controls-remove");buttons->addWidget(remove);layout->addLayout(buttons);
    QObject::connect(add,&QPushButton::clicked,&dialog,[&] {if(grid->rowCount()>=128) return;int ordinal=1;QString name;bool used;do {name="Control"+QString::number(ordinal++);used=fields.contains(name,Qt::CaseInsensitive);for(int row=0;row<grid->rowCount();++row) used|=grid->item(row,1)->text().compare(name,Qt::CaseInsensitive)==0;}while(used);append(QJsonObject{{"type","TextBox"},{"name",name},{"left",240},{"top",240+(fields.size()+grid->rowCount())*420},{"width",2200},{"height",300}});grid->setCurrentCell(grid->rowCount()-1,1);});
    QObject::connect(remove,&QPushButton::clicked,&dialog,[&] {if(grid->currentRow()>=0) grid->removeRow(grid->currentRow());});
    auto* status=new QLabel(&dialog);status->setObjectName("native-form-controls-status");status->setWordWrap(true);layout->addWidget(status);
    auto* done=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);layout->addWidget(done);QObject::connect(done,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    QObject::connect(done,&QDialogButtonBox::accepted,&dialog,[&] {
        try {QJsonArray next;
            for(int row=0;row<grid->rowCount();++row) {QJsonObject object;
                for(int column=0;column<keys.size();++column) {
                    if(column==0 || column==8) object.insert(keys[column],qobject_cast<QComboBox*>(grid->cellWidget(row,column))->currentText());
                    else if(column>=4 && column<=7) {bool valid=false;const auto number=grid->item(row,column)->text().toInt(&valid);if(!valid) throw VbaRuntimeError(5,"Position und Größe müssen ganze Twip-Werte sein.");object.insert(keys[column],number);}
                    else object.insert(keys[column],grid->item(row,column)->text());
                }
                next.append(object);
            }
            (void)nativeFormControls(next,fields,schema,{});controls=next;dialog.accept();
        } catch(const VbaRuntimeError& error) {status->setText("Nicht übernommen ("+QString::number(error.number)+"): "+QString::fromUtf8(error.what()));}
    });
    return dialog.exec()==QDialog::Accepted;
}
}
