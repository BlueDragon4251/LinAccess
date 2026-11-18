#include "access_ole_package.hpp"

#include <QtEndian>
#include <algorithm>

namespace linaccess {
namespace {

constexpr qsizetype maxBlobBytes = 16 * 1024 * 1024;
constexpr qsizetype maxPathUnits = 32767;
constexpr QByteArrayView packageClass("Package\0", 8);
constexpr QByteArrayView emptyPresentation("\x01\x05\0\0\0\0\0\0\x01\xad\x05\xfe", 12);

class PackageReader {
public:
    explicit PackageReader(QByteArrayView bytes) : remaining_(bytes) {}

    qsizetype remaining() const { return remaining_.size(); }

    QByteArrayView read(quint32 size) {
        if (size > static_cast<quint32>(remaining())) {
            return {};
        }
        const auto bytes = remaining_.first(size);
        remaining_ = remaining_.sliced(size);
        return bytes;
    }

    template<class T>
    std::optional<T> number() {
        const auto field = read(sizeof(T));
        if (field.isNull()) {
            return std::nullopt;
        }
        return qFromLittleEndian<T>(field.data());
    }

    QByteArrayView readAscii() {
        const auto end = remaining_.first(std::min(remaining(), maxPathUnits + 1)).indexOf('\0');
        return end < 0 ? QByteArrayView{} : read(end + 1);
    }

private:
    QByteArrayView remaining_;
};

bool isTerminatedAscii(QByteArrayView text) {
    return !text.isEmpty() && text.size() <= maxPathUnits + 1 && text.back() == '\0'
        && std::all_of(text.begin(), text.end() - 1, [](char value) {
            return value >= ' ' && value <= '~';
        });
}

std::optional<QString> readUnicodeText(PackageReader& reader) {
    const auto units = reader.number<quint32>();
    if (!units || *units > maxPathUnits || *units > static_cast<quint32>(reader.remaining() / 2)) {
        return std::nullopt;
    }
    const auto bytes = reader.read(*units * 2);
    QString text(*units, Qt::Uninitialized);
    for (qsizetype i = 0; i < text.size(); ++i) {
        text[i] = QChar(qFromLittleEndian<quint16>(bytes.data() + i * 2));
    }
    if (!text.isValidUtf16() || text.contains(QChar(0))) {
        return std::nullopt;
    }
    return text;
}

bool isDisplayName(const QString& name) {
    return !name.isEmpty() && name != u"." && name != u".."
        && !name.contains(u'/') && !name.contains(u'\\')
        && std::none_of(name.begin(), name.end(), [](QChar character) {
            return character.category() == QChar::Other_Control;
        });
}

std::optional<AccessOlePackage> extractPackageStream(QByteArrayView data) {
    PackageReader stream(data);
    if (stream.number<quint16>() != 0x0002) {
        return std::nullopt;
    }
    const auto name = stream.readAscii();
    const auto sourcePath = stream.readAscii();
    // Linked files carry a different stream type. Paths are only parsed for structure.
    if (!isTerminatedAscii(name) || name.size() <= 1 || !isTerminatedAscii(sourcePath)
        || stream.number<quint32>() != 0x00030000) {
        return std::nullopt;
    }
    const auto pathSize = stream.number<quint32>();
    if (!pathSize || !isTerminatedAscii(stream.read(*pathSize))) {
        return std::nullopt;
    }
    const auto payloadSize = stream.number<quint32>();
    if (!payloadSize) {
        return std::nullopt;
    }
    const auto payload = stream.read(*payloadSize);
    if (payload.isNull()) {
        return std::nullopt;
    }

    // The supported variant includes all three UTF-16LE fields after the payload.
    const auto localPath = readUnicodeText(stream);
    const auto displayName = readUnicodeText(stream);
    const auto originalPath = readUnicodeText(stream);
    if (!localPath || !displayName || !originalPath || stream.remaining() != 0
        || !isDisplayName(*displayName)) {
        return std::nullopt;
    }
    return AccessOlePackage{*displayName, QByteArray(payload.data(), payload.size())};
}

}

std::optional<AccessOlePackage> extractAccessOlePackage(QByteArrayView blob) {
    if (blob.size() > maxBlobBytes) {
        return std::nullopt;
    }
    PackageReader reader(blob);
    if (reader.number<quint16>() != 0x1c15) {
        return std::nullopt;
    }
    const auto headerSize = reader.number<quint16>();
    const auto objectType = reader.number<quint32>();
    const auto prettySize = reader.number<quint16>();
    const auto classSize = reader.number<quint16>();
    const auto prettyOffset = reader.number<quint16>();
    const auto classOffset = reader.number<quint16>();
    constexpr quint16 fixedHeaderSize = 20;

    if (!headerSize || objectType != 2 || !prettySize || classSize != packageClass.size()
        || prettyOffset != fixedHeaderSize || classOffset != fixedHeaderSize + *prettySize
        || *headerSize != fixedHeaderSize + *prettySize + *classSize
        || reader.number<quint32>() != 0xffffffffu
        || !isTerminatedAscii(reader.read(*prettySize)) || reader.read(*classSize) != packageClass) {
        return std::nullopt;
    }

    // OLE1 embedded Package header with empty topic and item names.
    if (reader.number<quint32>() != 0x00000501 || reader.number<quint32>() != 2
        || reader.number<quint32>() != packageClass.size() || reader.read(packageClass.size()) != packageClass
        || reader.number<quint32>() != 0 || reader.number<quint32>() != 0) {
        return std::nullopt;
    }
    const auto dataSize = reader.number<quint32>();
    if (!dataSize) {
        return std::nullopt;
    }
    const auto data = reader.read(*dataSize);
    if (data.isNull() || reader.read(emptyPresentation.size()) != emptyPresentation || reader.remaining() != 0) {
        return std::nullopt;
    }
    return extractPackageStream(data);
}

}
