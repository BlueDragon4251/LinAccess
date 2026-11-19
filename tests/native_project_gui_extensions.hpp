#pragma once
#include <functional>

class MainWindow;
class QSqlDatabase;
void testNativeProjectGuiExtensions(MainWindow&,QSqlDatabase,const std::function<void(const char*)>& phase);
