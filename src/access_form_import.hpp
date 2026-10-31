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
    QString caption;
    QString controlSource;
    QString rowSource;
    QString sourceObject;
    QString section;
    QString onClick;
    QString afterUpdate;
    QString beforeUpdate;
    int left{};
    int top{};
    int width{1701};
    int height{283};
    bool visible{true};
};

struct AccessFormLayout {
    QString caption;
    QString recordSource;
    int width{9000};
    int height{5000};
    QVector<AccessVisualControl> controls;
    QMap<QString,QString> events;
};

AccessFormLayout parseAccessFormLayout(const QString& text);
}
