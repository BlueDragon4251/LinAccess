#include "access_ole_package.hpp"
#include "access_table_import.hpp"

#include <QCoreApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QSqlQuery>
#include <QtEndian>
#include <array>
#include <iostream>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

namespace {

using linaccess::extractAccessOlePackage;

// Captured from JetOlePackageReference with Jackcess 5.0.4. These small technical
// vectors exercise the parser in ordinary builds without Java or an MDB fixture.
const auto embedded = QByteArray::fromHex(
    "151c3200020000001600080014002a00ffffffff5061636b6167657220536865"
    "6c6c204f626a656374005061636b616765000105000002000000080000005061"
    "636b616765000000000000000000c20000000200726573756d652d3f2e62696e"
    "00433a5c6f6c655c726573756d652d3f2e62696e000000030014000000433a5c"
    "6f6c655c726573756d652d3f2e62696e000900000000ff2c220a0d0105001400"
    "000043003a005c006f006c0065005c007200e900730075006d00e9002d003dd8"
    "c4dc2e00620069006e000d0000007200e900730075006d00e9002d003dd8c4dc"
    "2e00620069006e001400000043003a005c006f006c0065005c007200e9007300"
    "75006d00e9002d003dd8c4dc2e00620069006e00010500000000000001ad05fe");
const auto empty = QByteArray::fromHex(
    "151c3200020000001600080014002a00ffffffff5061636b6167657220536865"
    "6c6c204f626a656374005061636b616765000105000002000000080000005061"
    "636b616765000000000000000000980000000200656d7074792e62696e00433a"
    "5c6f6c655c656d7074792e62696e000000030011000000433a5c6f6c655c656d"
    "7074792e62696e00000000001000000043003a005c006f006c0065005c006500"
    "6d007000740079002e00620069006e000900000065006d007000740079002e00"
    "620069006e001000000043003a005c006f006c0065005c0065006d0070007400"
    "79002e00620069006e00010500000000000001ad05fe");
const auto linked = QByteArray::fromHex(
    "151c3200020000001600080014002a00ffffffff5061636b6167657220536865"
    "6c6c204f626a656374005061636b616765000105000002000000080000005061"
    "636b6167650000000000000000003d000000020065787465726e616c2e62696e"
    "00433a5c6f6c655c65787465726e616c2e62696e00000001000100433a5c6f6c"
    "655c65787465726e616c2e62696e00010500000000000001ad05fe");
const auto embeddedName = QStringLiteral("r\u00e9sum\u00e9-\U0001f4c4.bin");
const auto embeddedBytes = QByteArray::fromHex("00ff2c220a0d010500");

void verifyParser() {
    auto original = embedded;
    const auto result = extractAccessOlePackage(original);
    assert(result && result->displayName == embeddedName && result->payload == embeddedBytes);
    assert(original == embedded);
    original.fill(0);
    assert(result->payload == embeddedBytes);

    const auto emptyResult = extractAccessOlePackage(empty);
    assert(emptyResult && emptyResult->displayName == "empty.bin" && emptyResult->payload.isEmpty());
    assert(!emptyResult->payload.isNull());
    assert(!extractAccessOlePackage(linked));
    assert(!extractAccessOlePackage(embeddedBytes));
    assert(!extractAccessOlePackage(QByteArray::fromHex("d0cf11e0a1b11ae1")));

    for (qsizetype size = 0; size < embedded.size(); ++size) {
        assert(!extractAccessOlePackage(QByteArrayView(embedded).first(size)));
    }
    assert(!extractAccessOlePackage(embedded + '\0'));
    assert(!extractAccessOlePackage(QByteArray(16 * 1024 * 1024 + 1, '\0')));

    // Oversized lengths for native data, local path, file data and Unicode filename.
    for (const auto offset : {78, 121, 145, 202}) {
        auto changed = embedded;
        qToLittleEndian<quint32>(0xffffffffu, changed.data() + offset);
        assert(!extractAccessOlePackage(changed));
    }
    // Header length, format class, sized path terminator, filename slash, invalid surrogate and footer.
    for (const auto offset : {2, 62, 144, 206, 223, 287}) {
        auto changed = embedded;
        changed[offset] = offset == 223 ? static_cast<char>(0xd8) : '/';
        assert(!extractAccessOlePackage(changed));
    }
}

QByteArray readFile(const QString& path) {
    QFile file(path);
    assert(file.open(QIODevice::ReadOnly));
    return file.readAll();
}

void verifyImportedFixture(const QString& path) {
    const auto sourceHash = QCryptographicHash::hash(readFile(path), QCryptographicHash::Sha256);
    const auto directory = QFileInfo(path).dir();
    auto db = QSqlDatabase::addDatabase("QSQLITE", "jet-ole-package");
    db.setDatabaseName(":memory:");
    assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    query.finish();
    assert(linaccess::importAccessDatabase(db, path) == 1);
    assert(query.exec("SELECT id,payload,typeof(payload) FROM ole_package_probe ORDER BY id"));
    const std::array expectedWrappers{embedded, empty, linked};
    for (int id = 1; id <= 3; ++id) {
        assert(query.next() && query.value(0).toInt() == id && query.value(2).toString() == "blob");
        const auto wrapper = query.value(1).toByteArray();
        const auto expected = readFile(directory.filePath(QString("ole-wrapper-%1.bin").arg(id)));
        assert(wrapper == expected && wrapper == expectedWrappers[id - 1]);
        const auto result = extractAccessOlePackage(wrapper);
        if (id == 3) {
            assert(!result);
        } else {
            const auto payload = readFile(directory.filePath(QString("ole-content-%1.bin").arg(id)));
            assert(result && result->displayName == (id == 1 ? embeddedName : QStringLiteral("empty.bin")));
            assert(result->payload == payload);
        }
    }
    assert(!query.next());
    query.finish();
    assert(query.exec("SELECT payload FROM ole_package_probe WHERE id=1"));
    assert(query.next() && query.value(0).toByteArray() == embedded);
    assert(QCryptographicHash::hash(readFile(path), QCryptographicHash::Sha256) == sourceHash);
    std::cout << "Jet4 OLE Package: imported BLOBs, embedded contents and unchanged source independently verified\n";
}

}

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    verifyParser();
    const auto fixture = qEnvironmentVariable("LINACCESS_JET_OLE_PACKAGE_REFERENCE");
    if (!fixture.isEmpty()) {
        verifyImportedFixture(fixture);
    } else {
        std::cout << "OLE Package parser vectors passed. Optional Jet4 fixture path was unset.\n";
    }
}
