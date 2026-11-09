#pragma once
class MainWindow;
class QSqlDatabase;
void testNativeRepaintObject(MainWindow&,QSqlDatabase);

void testNativeSetValueDoubleEvaluation(MainWindow& window,QSqlDatabase database);

void testNativeFormRecordFields(MainWindow&,QSqlDatabase);

void testNativeProjectFilters(MainWindow&,QSqlDatabase);
