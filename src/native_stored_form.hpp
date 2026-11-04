#pragma once
#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include <QByteArray>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSet>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlRecord>
#include <array>
#include <cmath>

namespace linaccess {
struct NativeStoredForm { QString name; QByteArray layout,code; };

// Runtime adapter for owned _la_objects definitions, not an Access export or
// an alteration of imported SaveAsText/VBA bytes.
inline NativeStoredForm nativeStoredForm(QString name,const QJsonObject& definition,QSqlDatabase database) {
    const auto quoted=[](QString value) {
        if(value.contains('\n') || value.contains('\r') || value.contains(QChar::Null))
            throw VbaRuntimeError(3251,"Multiline native form metadata is not implemented");
        value.replace('"',"\"\""); return QStringLiteral("\"")+value+QStringLiteral("\"");
    };
    const auto table=definition.value("table").toString();
    const auto fields=definition.value("fields").toArray();
    if(!database.isOpen() || table.isEmpty() || fields.isEmpty() || fields.size()>255
        || !database.tables().contains(table,Qt::CaseInsensitive))
        throw VbaRuntimeError(3251,"Native stored form requires a table and 1-255 fields");
    const auto schema=database.record(table);
    QStringList names; QSet<QString> used;
    for(const auto& field:fields) {
        if(!field.isString() || field.toString().isEmpty() || field.toString().startsWith('='))
            throw VbaRuntimeError(3251,"Native stored form field binding is unsupported");
        const auto column=field.toString(); bool found=false;
        for(int index=0;index<schema.count();++index) if(schema.fieldName(index).compare(column,Qt::CaseInsensitive)==0) { found=true; break; }
        if(!found || used.contains(column.toCaseFolded())) throw VbaRuntimeError(3251,"Native stored form field is missing or duplicated");
        (void)quoted(column); names.push_back(column); used.insert(column.toCaseFolded());
    }
    const auto unique=[&used](QString base) {
        auto candidate=base; int suffix=1;
        while(used.contains(candidate.toCaseFolded())) candidate=base+QString::number(suffix++);
        used.insert(candidate.toCaseFolded()); return candidate;
    };
    auto code=definition.value("vba").toString(); QSet<QString> procedures;
    const auto utf8=code.toUtf8(); const auto analysis=analyzeVba(std::string_view(utf8.constData(),static_cast<std::size_t>(utf8.size())));
    for(const auto& procedure:analysis.procedures) procedures.insert(QString::fromStdString(procedure.name).toCaseFolded());
    const auto prefix=[](QString value) {
        for(auto& character:value) if(!character.isLetterOrNumber() && character!='_') character='_';
        if(value.isEmpty() || value.front().isDigit()) value.prepend('_');
        return value;
    };
    const auto positions=definition.value("field_layout").toObject();
    QList<std::array<int,4>> geometry;
    int formWidth=6000,buttonTop=240+names.size()*420;
    for(int index=0;index<names.size();++index) {
        const auto saved=positions.value(names[index]);
        if(!saved.isUndefined() && !saved.isObject()) throw VbaRuntimeError(3251,"Native field layout must be an object");
        const auto properties=saved.toObject();
        std::array<int,4> rectangle{2160,240+index*420,3300,300};
        const std::array<QString,4> keys{"left","top","width","height"};
        for(int part=0;part<4;++part) if(properties.contains(keys[part])) {
            const auto value=properties.value(keys[part]); const auto number=value.toDouble(-1);
            if(!value.isDouble() || !std::isfinite(number) || number!=std::floor(number) || number<(part<2 ? 0 : 1) || number>30000)
                throw VbaRuntimeError(5,"Native field geometry is outside supported twip bounds");
            rectangle[part]=static_cast<int>(number);
        }
        if(rectangle[0]+rectangle[2]>30000 || rectangle[1]+rectangle[3]>30000)
            throw VbaRuntimeError(5,"Native field extends outside supported layout bounds");
        formWidth=qMax(formWidth,rectangle[0]+rectangle[2]+240);
        buttonTop=qMax(buttonTop,rectangle[1]+rectangle[3]+240);
        geometry.push_back(rectangle);
    }
    QStringList lines{"Begin Form","    Caption = "+quoted(definition.value("caption").toString(name)),
        "    RecordSource = "+quoted(table),"    Width = "+QString::number(formWidth)};
    for(const auto* event:{"OnOpen","OnLoad","OnCurrent","OnActivate","OnDeactivate","OnUnload","OnClose","BeforeInsert","AfterInsert","BeforeUpdate","AfterUpdate","OnDirty","OnError"}) {
        const QString property=QString::fromLatin1(event);
        auto procedure=property.startsWith("On") ? property.mid(2) : property;
        if(procedures.contains(("Form_"+procedure).toCaseFolded())) lines<<"    "+property+" = \"[Event Procedure]\"";
    }
    lines<<"    Begin Section"<<"        Name = \"Detail\""<<"        Height = "+QString::number(buttonTop+700);
    const auto captions=definition.value("field_captions").toObject();
    for(int index=0;index<names.size();++index) {
        const auto& field=names[index]; const auto eventPrefix=prefix(field);
        const auto& rectangle=geometry[index]; const int top=rectangle[1];
        lines<<"        Begin Label"<<"            Name = "+quoted(unique("LinAccessLabel"))
             <<"            Caption = "+quoted(captions.value(field).toString(field))
             <<"            Left = "+QString::number(qMax(0,rectangle[0]-1920))<<"            Top = "+QString::number(top)<<"            Width = 1800"<<"            Height = 300"<<"        End";
        lines<<"        Begin TextBox"<<"            Name = "+quoted(field)<<"            ControlSource = "+quoted(field)
             <<"            EventProcPrefix = "+quoted(eventPrefix)<<"            Left = "+QString::number(rectangle[0])<<"            Top = "+QString::number(top)
             <<"            Width = "+QString::number(rectangle[2])<<"            Height = "+QString::number(rectangle[3])<<"            TabIndex = "+QString::number(index);
        for(const auto* event:{"OnClick","OnChange","BeforeUpdate","AfterUpdate","OnDblClick","OnKeyDown","OnKeyPress","OnKeyUp","OnMouseDown","OnMouseUp"}) {
            const QString property=QString::fromLatin1(event);
            const auto procedure=property.startsWith("On") ? property.mid(2) : property;
            if(procedures.contains((eventPrefix+"_"+procedure).toCaseFolded())) lines<<"            "+property+" = \"[Event Procedure]\"";
        }
        lines<<"        End";
    }
    const auto button=[&](const QString& eventPrefix,const QString& caption,int left,const QString& fallback) {
        const auto handler=eventPrefix+"_Click";
        if(!procedures.contains(handler.toCaseFolded())) {
            if(fallback.isEmpty()) return;
            code+="\nPrivate Sub "+handler+"()\n"+fallback+"\nEnd Sub\n";
        }
        lines<<"        Begin CommandButton"<<"            Name = "+quoted(unique(eventPrefix))<<"            EventProcPrefix = "+quoted(eventPrefix)
             <<"            Caption = "+quoted(caption)<<"            Left = "+QString::number(left)<<"            Top = "+QString::number(buttonTop)
             <<"            Width = 1500"<<"            Height = 420"<<"            TabIndex = "+QString::number(names.size()+left/1700)
             <<"            OnClick = \"[Event Procedure]\"";
        for(const auto* event:{"MouseDown","MouseUp"})
            if(procedures.contains((eventPrefix+"_"+event).toCaseFolded())) lines<<"            On"+QString::fromLatin1(event)+" = \"[Event Procedure]\"";
        lines<<"        End";
    };
    button("btnNeu","Neu",240,"DoCmd.GoToRecord , , acNewRec");
    button("btnSpeichern","Speichern",1940,"DoCmd.RunCommand acCmdSaveRecord");
    button("btnAktion","Aktion",3640,{});
    lines<<"    End"<<"End";
    return {std::move(name),lines.join('\n').toUtf8(),code.toUtf8()};
}

inline std::optional<NativeStoredForm> loadNativeStoredForm(QSqlDatabase database,const QString& name) {
    if(!database.isOpen() || !database.tables().contains("_la_objects",Qt::CaseInsensitive)) return {};
    QSqlQuery query(database); QString canonical; int rows=0;
    if(!query.exec("SELECT name FROM main._la_objects WHERE kind='form'")) throw VbaRuntimeError(3146,"Native form catalog query failed");
    while(query.next()) {
        if(++rows>1000000) throw VbaRuntimeError(3251,"Native form catalog limit exceeded");
        const auto candidate=query.value(0).toString();
        if(candidate.compare(name,Qt::CaseInsensitive)!=0) continue;
        if(!canonical.isEmpty()) throw VbaRuntimeError(5,"Native form name is ambiguous");
        canonical=candidate;
    }
    query.finish(); if(canonical.isEmpty()) return {};
    if(database.tables().contains("_la_access_artifacts",Qt::CaseInsensitive)) {
        if(!query.exec("SELECT name FROM main._la_access_artifacts WHERE kind='form'")) throw VbaRuntimeError(3146,"Imported form catalog query failed");
        while(query.next()) {
            if(++rows>1000000) throw VbaRuntimeError(3251,"Native form catalog limit exceeded");
            if(query.value(0).toString().compare(canonical,Qt::CaseInsensitive)==0) throw VbaRuntimeError(5,"Stored and imported form names conflict");
        }
        query.finish();
    }
    query.prepare("SELECT definition FROM main._la_objects WHERE kind='form' AND name=?"); query.addBindValue(canonical);
    if(!query.exec() || !query.next()) throw VbaRuntimeError(3146,"Native form definition query failed");
    const auto bytes=query.value(0).toByteArray(); query.finish();
    QJsonParseError error; const auto document=QJsonDocument::fromJson(bytes,&error);
    if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(3251,"Native form definition is invalid");
    return nativeStoredForm(canonical,document.object(),database);
}
}
