#include "mainwindow.hpp"
#include "native_module_debugger.hpp"
#include "native_standard_module_host.hpp"
#include "vba_debug_output.hpp"
#include <QAction>
#include <QApplication>
#include <QDialog>
#include <QDockWidget>
#include <QLabel>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
void exerciseNativeDebugOutput(MainWindow& window,QSqlDatabase database){using namespace linaccess;clearVbaDebugOutput();auto* dock=window.findChild<QDockWidget*>("native-debug-output-dock");auto* toggle=window.findChild<QAction*>("native-debug-output-toggle");assert(dock && toggle && !dock->isVisible());
    NativeStandardModuleHost host(&window,database,[]{});const auto result=executeVbaProcedure("Sub Output()\nDebug.Print \"real\";Spc(2);\"output\"\nEnd Sub","Output",host);assert(result.status==VbaExecutionStatus::Completed && vbaDebugOutputSnapshot().text=="real  output\n");toggle->trigger();QApplication::processEvents();auto* text=dock->findChild<QPlainTextEdit*>("native-debug-output-text");auto* clear=dock->findChild<QPushButton*>("native-debug-output-clear");assert(dock->isVisible() && text && text->isReadOnly() && text->toPlainText()=="real  output\n" && clear);clear->click();assert(text->toPlainText().isEmpty() && vbaDebugOutputSnapshot().text.empty());
    QDialog editor(&window);NativeModuleBreakpoints breakpoints;bool observed=false;QTimer resume;resume.setInterval(1);QObject::connect(&resume,&QTimer::timeout,&editor,[&]{auto* debugger=editor.findChild<QDialog*>("native-module-debugger");if(!debugger || !debugger->isVisible())return;auto* status=debugger->findChild<QLabel*>("native-debugger-status");if(status && status->text().startsWith("Stop-Anweisung")){auto* output=debugger->findChild<QPlainTextEdit*>("native-debugger-output-text");assert(output && output->isReadOnly() && output->toPlainText()=="before\n");observed=true;debugger->findChild<QPushButton*>("native-debugger-continue")->click();}});resume.start();
    const auto execution=runNativeModuleDebugger(editor,database,"OutputFixture","Public Function Output() As Long\nDebug.Print \"before\"\nStop\nDebug.Print \"after\"\nOutput=2\nEnd Function","Output",breakpoints,false,[]{});resume.stop();assert(observed && execution.execution.status==VbaExecutionStatus::Completed && execution.execution.typedReturnValue && execution.execution.typedReturnValue->number==2 && vbaDebugOutputSnapshot().text=="before\nafter\n");dock->hide();dock->show();QApplication::processEvents();assert(text->toPlainText()=="before\nafter\n");clear->click();assert(text->toPlainText().isEmpty());toggle->trigger();QApplication::processEvents();assert(!dock->isVisible());window.focusDatabaseWindow();
}
