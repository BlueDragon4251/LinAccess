#include "mainwindow.hpp"

#include <QApplication>

int main(int argc, char* argv[]) {
    QApplication application(argc, argv);
    application.setApplicationName("LinAccess");
    application.setOrganizationName("LinAccess Community");
    MainWindow window;
    window.show();
    return application.exec();
}
