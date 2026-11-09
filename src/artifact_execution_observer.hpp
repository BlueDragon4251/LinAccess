#pragma once
#include <QSqlDatabase>
#include <QString>
#include <QByteArray>
namespace linaccess {
void recordArtifactExecution(QSqlDatabase,const QString& kind,const QString& name,const QByteArray& source,const QString& status,int error=0);
}
