#pragma once

#include <QString>
#include <QStringList>
#include <QVector>
#include <QMap>

namespace linaccess {
// A deliberately partial, source-only decoder for Access SaveAsText form
// definitions. It does not reconstruct exact geometry, nested subforms,
// expression controls, combo row sources or any Access-native layout binary.
struct AccessFormDraft {
    QString recordSource;
    QStringList controlSources;
    QStringList eventNames;
};

AccessFormDraft parseAccessFormText(const QString& text);

// Access 2000/2003 SaveAsText uses TWIPs (1/1440 inch) for geometry.
// This preserves the layout and event metadata for controls we can draw; it
// does not claim to reproduce the proprietary Access rendering engine.
struct AccessVisualControl {
    QString type;
    QString name;
    QString eventProcPrefix;
    QString caption;
    QString tag;
    QString controlSource;
    QString defaultValue;
    int optionValue{1};
    int optionGroupIndex{-1};
    QString rowSource;
    QString rowSourceType;
    int boundColumn{1};
    int columnCount{1};
    QString columnWidths;
    QString sourceObject;
    QString linkMasterFields;
    QString linkChildFields;
    QString section;
    QString onClick;
    QString afterUpdate;
    QString beforeUpdate;
    QString onDblClick;
    QString onKeyDown;
    QString onKeyPress;
    QString onKeyUp;
    QString className;
    QString oleClass;
    int left{};
    int top{};
    int width{1701};
    int height{283};
    bool visible{true};
};

enum class AccessDocumentKind { Form, Report };
struct AccessFormLayout {
    AccessDocumentKind kind{AccessDocumentKind::Form};
    QString caption;
    QString tag;
    QString recordSource;
    int width{9000};
    QString filter;
    bool filterOnLoad{};
    bool modal{};
    bool keyPreview{};
    bool allowDeletions{true};
    int height{5000};
    QVector<AccessVisualControl> controls;
    QMap<QString,QString> events;
};

AccessFormLayout parseAccessFormLayout(const QString& text);
}
