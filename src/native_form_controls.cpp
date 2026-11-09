#include "native_form_controls.hpp"
#include "native_form_events.hpp"
#include "native_column_widths.hpp"
#include "native_control_default.hpp"
#include <QDialog>
#include <QDialogButtonBox>
#include <QComboBox>
#include <QCheckBox>
#include <QSpinBox>
#include <QGroupBox>
#include <QSignalBlocker>
#include <functional>
#include <memory>
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
const QStringList keys{"type","name","caption","control_source","left","top","width","height","row_source_type","row_source","onclick","column_count","bound_column","column_widths","default_value","format","input_mask"};
const QStringList stateKeys{"visible","enabled","locked","tab_stop","tab_index"};
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
        for(auto entry=object.begin();entry!=object.end();++entry) if(!keys.contains(entry.key()) && !stateKeys.contains(entry.key())) throw VbaRuntimeError(3251,"Unsupported native control property");
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
        const auto defaultValue=string("default_value"),format=string("format"),inputMask=string("input_mask");
        if(defaultValue.size()>255 || format.size()>1024) throw VbaRuntimeError(5,"Native default/format property exceeds its limit");
        if(!defaultValue.isEmpty()) {
            if(type=="Label" || type=="CommandButton" || source.startsWith('=')) throw VbaRuntimeError(5,"This control has no editable default value");
            (void)nativeControlDefaultLiteral(defaultValue);
        }
        if((!format.isEmpty() || !inputMask.isEmpty()) && type!="TextBox") throw VbaRuntimeError(5,"Format and InputMask require a TextBox");
        if(!inputMask.isEmpty() && inputMask.compare("Password",Qt::CaseInsensitive)!=0) throw VbaRuntimeError(3251,"Only native Password input masks are implemented");
        const auto integer=[&](const QString& key,int fallback,int minimum) {
            const auto value=object.value(key);if(value.isUndefined()) return fallback;const auto number=value.toDouble(-1);
            if(!value.isDouble() || !std::isfinite(number) || number!=std::floor(number) || number<minimum || number>(key=="tab_index" ? 32767 : 30000)) throw VbaRuntimeError(5,"Invalid native control integer property");return static_cast<int>(number);
        };
        const auto flag=[&](const QString& key,bool fallback) {const auto value=object.value(key);if(value.isUndefined()) return fallback;if(!value.isBool()) throw VbaRuntimeError(13,"Native control state requires a Boolean");return value.toBool();};
        const bool visible=flag("visible",true),enabled=flag("enabled",true),locked=flag("locked",false),tabStop=flag("tab_stop",type!="Label");
        const int tabIndex=integer("tab_index",type=="Label" ? 0 : ordinal,0);
        if(tabIndex>32767 || (type=="Label" && (tabStop || tabIndex!=0))) throw VbaRuntimeError(5,"Invalid native tab configuration");
        if(locked && type!="TextBox" && type!="ListBox") throw VbaRuntimeError(3251,"Locked authoring requires a TextBox or ListBox");
        const int columns=integer("column_count",1,1),bound=integer("bound_column",1,0);const auto columnWidths=string("column_widths");
        if(columns>255 || bound>columns) throw VbaRuntimeError(5,"Invalid native list column configuration");
        if(!list && (columns!=1 || bound!=1 || !columnWidths.isEmpty())) throw VbaRuntimeError(5,"Only list controls have column properties");
        const auto widths=nativeColumnWidths(columnWidths);
        if(widths.size()>columns) throw VbaRuntimeError(5,"Column widths exceed the configured column count");
        const int left=integer("left",240,0),top=integer("top",240+ordinal*420,0),width=integer("width",2200,1),height=integer("height",300,1);
        if(left+width>30000 || top+height>30000) throw VbaRuntimeError(5,"Native control extends beyond supported geometry");
        result.width=qMax(result.width,left+width+240);result.bottom=qMax(result.bottom,top+height+240);
        auto& lines=result.lines;lines<<"        Begin "+type<<"            Name = "+quoted(name)<<"            EventProcPrefix = "+quoted(name)
            <<"            Left = "+QString::number(left)<<"            Top = "+QString::number(top)<<"            Width = "+QString::number(width)<<"            Height = "+QString::number(height)<<"            Visible = "+QString::number(visible ? -1 : 0)<<"            Enabled = "+QString::number(enabled ? -1 : 0);
        if(type!="Label") lines<<"            TabIndex = "+QString::number(tabIndex)<<"            TabStop = "+QString::number(tabStop ? -1 : 0);
        if(type=="TextBox" || type=="ListBox") lines<<"            Locked = "+QString::number(locked ? -1 : 0);
        ++ordinal;
        if(!caption.isEmpty()) lines<<"            Caption = "+quoted(caption);
        if(!source.isEmpty()) lines<<"            ControlSource = "+quoted(source);
        if(!defaultValue.isEmpty()) lines<<"            DefaultValue = "+quoted(defaultValue);
        if(!format.isEmpty()) lines<<"            Format = "+quoted(format);
        if(!inputMask.isEmpty()) lines<<"            InputMask = "+quoted(inputMask);
        if(list) lines<<"            RowSourceType = "+quoted(rowType)<<"            RowSource = "+quoted(rows)
            <<"            ColumnCount = "+QString::number(columns)<<"            BoundColumn = "+QString::number(bound)<<"            ColumnWidths = "+quoted(columnWidths);
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
    auto* layout=new QVBoxLayout(&dialog);auto* help=new QLabel("Eigene Steuerelemente anlegen. Position und Größe in Twips. Beschriftungen für Label, Schaltfläche und Checkbox; andere Eingabefelder erhalten bei Bedarf ein separates Label. Spaltenbreiten: Twips oder cm/mm/in/pt, mit Semikolon getrennt; 0 blendet eine Spalte aus. Gebundene Spalte: 1-basiert, 0 für Zeilennummer. Standardwerte: Zahlen, Wahrheitswerte, Null oder Text in doppelten Anführungszeichen. Eingabemaske: leer oder Password. Leeres Klickereignis verwendet eine vorhandene VBA-Prozedur automatisch.",&dialog);help->setWordWrap(true);layout->addWidget(help);
    QObject stateConnections; // Disconnect before stack callbacks and child widgets are destroyed.
    auto refreshStates=std::make_shared<std::function<void()>>();
    auto* grid=new QTableWidget(0,keys.size(),&dialog);grid->setObjectName("native-form-controls-grid");grid->setHorizontalHeaderLabels({"Typ","Name","Beschriftung","Steuerelementinhalt","Links","Oben","Breite","Höhe","Herkunftsart","Datensatzherkunft","Beim Klicken","Spaltenanzahl","Gebundene Spalte","Spaltenbreiten","Standardwert","Format","Eingabemaske"});grid->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);layout->addWidget(grid);
    const auto append=[&](const QJsonObject& object) {
        const int row=grid->rowCount();grid->insertRow(row);
        for(int column=0;column<keys.size();++column) {
            if(column==0 || column==8) {auto* combo=new QComboBox(grid);combo->addItems(column==0 ? types : QStringList{"Table/Query","Value List"});combo->setCurrentText(object.value(keys[column]).toString(column==0 ? "TextBox" : "Table/Query"));grid->setCellWidget(row,column,combo);
                if(column==0) QObject::connect(combo,&QComboBox::currentTextChanged,&stateConnections,[grid,combo,refreshStates] {for(int index=0;index<grid->rowCount();++index) if(grid->cellWidget(index,0)==combo) {grid->setCurrentCell(index,0);break;}if(*refreshStates) (*refreshStates)();});
            }
            else {const int defaults[]{240,240+(static_cast<int>(fields.size())+row)*420,2200,300};const auto value=object.contains(keys[column]) ? object.value(keys[column]) : column>=4 && column<=7 ? QJsonValue(defaults[column-4]) : column==11 || column==12 ? QJsonValue(1) : QJsonValue(QString{});auto* item=new QTableWidgetItem(value.isDouble() ? QString::number(value.toInt()) : value.toString());grid->setItem(row,column,item);}
        }
        QJsonObject state;for(const auto& key:stateKeys) if(object.contains(key)) state.insert(key,object.value(key));grid->item(row,1)->setData(Qt::UserRole,state);
    };
    for(const auto& value:controls) append(value.toObject());
    auto* states=new QGroupBox("Ausgewähltes Steuerelement",&dialog);auto* stateLayout=new QHBoxLayout(states);layout->addWidget(states);
    auto* visible=new QCheckBox("Sichtbar",states);visible->setObjectName("native-form-control-visible");stateLayout->addWidget(visible);
    auto* enabled=new QCheckBox("Aktiviert",states);enabled->setObjectName("native-form-control-enabled");stateLayout->addWidget(enabled);
    auto* locked=new QCheckBox("Gesperrt",states);locked->setObjectName("native-form-control-locked");stateLayout->addWidget(locked);
    auto* tabStop=new QCheckBox("Tabstopp",states);tabStop->setObjectName("native-form-control-tab-stop");stateLayout->addWidget(tabStop);
    stateLayout->addWidget(new QLabel("Tabreihenfolge",states));auto* tabIndex=new QSpinBox(states);tabIndex->setObjectName("native-form-control-tab-index");tabIndex->setRange(0,32767);stateLayout->addWidget(tabIndex);
    const int fieldCount=fields.size();
    *refreshStates=[grid,states,visible,enabled,locked,tabStop,tabIndex,fieldCount] {
        auto* item=grid->item(grid->currentRow(),1);states->setEnabled(item!=nullptr);if(!item) return;
        const auto type=qobject_cast<QComboBox*>(grid->cellWidget(grid->currentRow(),0))->currentText();const auto data=item->data(Qt::UserRole).toJsonObject();
        const QSignalBlocker a(visible),b(enabled),c(locked),d(tabStop),e(tabIndex);
        visible->setChecked(data.value("visible").toBool(true));enabled->setChecked(data.value("enabled").toBool(true));
        const bool canLock=type=="TextBox" || type=="ListBox";locked->setEnabled(canLock);locked->setChecked(canLock && data.value("locked").toBool());
        tabStop->setEnabled(type!="Label");tabStop->setChecked(type!="Label" && data.value("tab_stop").toBool(true));
        tabIndex->setEnabled(type!="Label");tabIndex->setValue(type=="Label" ? 0 : data.value("tab_index").toInt(fieldCount+grid->currentRow()));
    };
    const auto storeStates=[grid,visible,enabled,locked,tabStop,tabIndex] {
        if(auto* item=grid->item(grid->currentRow(),1)) item->setData(Qt::UserRole,QJsonObject{{"visible",visible->isChecked()},{"enabled",enabled->isChecked()},{"locked",locked->isChecked()},{"tab_stop",tabStop->isChecked()},{"tab_index",tabIndex->value()}});
    };
    QObject::connect(grid,&QTableWidget::currentCellChanged,&stateConnections,[refreshStates] {if(*refreshStates) (*refreshStates)();});
    for(auto* box:{visible,enabled,locked,tabStop}) QObject::connect(box,&QCheckBox::toggled,&stateConnections,storeStates);
    QObject::connect(tabIndex,&QSpinBox::valueChanged,&stateConnections,storeStates);
    if(grid->rowCount()) grid->setCurrentCell(0,1);(*refreshStates)();
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
                    else if((column>=4 && column<=7) || column==11 || column==12) {bool valid=false;const auto number=grid->item(row,column)->text().toInt(&valid);if(!valid) throw VbaRuntimeError(5,"Position, Größe und Spaltenangaben müssen ganze Werte sein.");object.insert(keys[column],number);}
                    else object.insert(keys[column],grid->item(row,column)->text());
                }
                const auto state=grid->item(row,1)->data(Qt::UserRole).toJsonObject();const auto type=object.value("type").toString();
                object.insert("visible",state.value("visible").toBool(true));object.insert("enabled",state.value("enabled").toBool(true));
                object.insert("locked",(type=="TextBox" || type=="ListBox") && state.value("locked").toBool());
                object.insert("tab_stop",type!="Label" && state.value("tab_stop").toBool(true));
                object.insert("tab_index",type=="Label" ? 0 : state.value("tab_index").toInt(fieldCount+row));
                next.append(object);
            }
            (void)nativeFormControls(next,fields,schema,{});controls=next;dialog.accept();
        } catch(const VbaRuntimeError& error) {status->setText("Nicht übernommen ("+QString::number(error.number)+"): "+QString::fromUtf8(error.what()));}
    });
    return dialog.exec()==QDialog::Accepted;
}
}
