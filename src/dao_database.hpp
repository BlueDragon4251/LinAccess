#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
#include <QMap>
#include <QStringList>
#include <QVariant>
#include <QVector>

namespace linaccess {
// SQLite-backed subset of DAO. Unsupported members throw; original MDBs are
// never opened for writing. Database must refer to the LinAccess project copy.
std::shared_ptr<VbaObject> makeDaoDatabase(QSqlDatabase database);
std::shared_ptr<VbaObject> makeDaoSnapshot(QSqlDatabase database,const QStringList& columns,
                                        const QMap<QString,QString>& types,const QVector<QVariantList>& rows,const QVector<QVariant>& ids={});
void refreshDaoSnapshot(const std::shared_ptr<VbaObject>& recordset,const QStringList& columns,
                        const QMap<QString,QString>& types,const QVector<QVariantList>& rows,const QVector<QVariant>& ids);
VbaValue daoBookmarkAt(const std::shared_ptr<VbaObject>& recordset,qsizetype position);
qsizetype daoBookmarkPosition(const std::shared_ptr<VbaObject>& recordset,const VbaValue& bookmark);
}
