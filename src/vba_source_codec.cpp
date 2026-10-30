#include "vba_source_codec.hpp"

#include <QChar>

namespace linaccess {
QString decodeVbaSource(const QByteArray& raw) {
    if (raw.startsWith("\xEF\xBB\xBF")) return QString::fromUtf8(raw.mid(3));
    if (raw.startsWith("\xFF\xFE")) {
        QString text;
        if (raw.size() % 2 != 0) return QStringLiteral("[Beschädigte UTF-16-Quelle]");
        text.reserve((raw.size() - 2) / 2);
        for (qsizetype i = 2; i < raw.size(); i += 2) {
            const auto value = static_cast<ushort>(static_cast<unsigned char>(raw[i]))
                | (static_cast<ushort>(static_cast<unsigned char>(raw[i + 1])) << 8);
            text.append(QChar(value));
        }
        return text;
    }
    const QString utf8 = QString::fromUtf8(raw);
    if (utf8.toUtf8() == raw) return utf8;

    // Windows-1252, including typographical characters in the 0x80-0x9F band.
    // Undefined bytes map to U+FFFD (a visible marker, not silent data loss).
    static constexpr char16_t cp1252Special[] = {
        0x20AC, 0xFFFD, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
        0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0xFFFD, 0x017D, 0xFFFD,
        0xFFFD, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
        0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0xFFFD, 0x017E, 0x0178,
    };
    QString result;
    result.reserve(raw.size());
    for (const unsigned char ch : raw) {
        if (ch >= 0x80 && ch <= 0x9f) result.append(QChar(cp1252Special[ch - 0x80]));
        else result.append(QChar(ch));
    }
    return result;
}
}
