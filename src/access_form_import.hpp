#pragma once

#include <QString>
#include <QStringList>
#include <QVector>
#include <QMap>
#include <optional>

namespace linaccess {
// A deliberately partial, source-only decoder for Access SaveAsText form
// definitions. Opaque Access layout binaries and complete rendering semantics
// remain outside this decoder; native runtime adapters implement bounded features.
struct AccessFormDraft {
    QString recordSource;
    QStringList controlSources;
    QStringList eventNames;
};

AccessFormDraft parseAccessFormText(const QString& text);

// Access 2000/2003 SaveAsText uses TWIPs (1/1440 inch) for geometry.
// This preserves the layout and event metadata for controls we can draw; it
// does not claim to reproduce the proprietary Access rendering engine.
struct AccessControlFont {
    std::optional<QString> name;
    std::optional<int> size;
    std::optional<int> weight;
    std::optional<bool> italic;
    std::optional<bool> underline;
};

struct AccessVisualControl {
    QString type;
    QString name;
    QString eventProcPrefix;
    QString caption;
    QString tag;
    QString controlSource;
    QString defaultValue;
    QString inputMask;
    AccessControlFont font;
    int textAlign{};
    std::optional<qint32> foreColor;
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
    QString onChange;
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
    bool enabled{true};
    bool locked{};
    bool tabStop{true};
    int tabIndex{};
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
    int timerInterval{};
    int cycle{};
    bool allowEdits{true};
    bool allowDeletions{true};
    bool allowAdditions{true};
    int height{5000};
    QVector<AccessVisualControl> controls;
    QMap<QString,QString> events;
};

AccessFormLayout parseAccessFormLayout(const QString& text);
}
