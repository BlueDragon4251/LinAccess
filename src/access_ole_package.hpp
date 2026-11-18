#pragma once

#include <QByteArray>
#include <QByteArrayView>
#include <QString>
#include <optional>

namespace linaccess {

struct AccessOlePackage {
    QString displayName; // Untrusted metadata. The caller chooses any destination path.
    QByteArray payload;
};

// Reads a bounded, embedded Access Package with the simple OLE1 wrapper.
// The input BLOB is read-only. Malformed and unsupported values return nullopt.
std::optional<AccessOlePackage> extractAccessOlePackage(QByteArrayView blob);

}
