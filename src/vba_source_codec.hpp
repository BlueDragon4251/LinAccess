#pragma once

#include <QByteArray>
#include <QString>

namespace linaccess {
// VBIDE exports from older Microsoft Access versions are frequently in a
// Windows ANSI codepage. Preserve the raw byte stream alongside this view.
// CP1252 is a fallback, not a guess about all possible system codepages.
QString decodeVbaSource(const QByteArray& raw);
}
