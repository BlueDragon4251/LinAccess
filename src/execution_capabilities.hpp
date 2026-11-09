#pragma once
#include "vba_runtime.hpp"
#include <QJsonArray>
#include <QJsonObject>
#include <QSqlDatabase>
namespace linaccess {
// Session observations contain source hashes and status metadata, never source/values.
void recordVbaObservation(QSqlDatabase,std::string_view,std::string_view,const VbaExecutionResult&,bool,const QString& context);
void clearExecutionObservations(QSqlDatabase);
QJsonArray executionObservations(QSqlDatabase);
QJsonObject projectCapabilityReport(QSqlDatabase);
}
