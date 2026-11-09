#pragma once
#include "vba_expression.hpp"
#include <QLayout>
#include <QMainWindow>
#include <QProgressBar>
#include <QStatusBar>

namespace linaccess {
inline QStatusBar* nativeStatusSurface(QWidget& document) {
    if(auto* window=qobject_cast<QMainWindow*>(&document)) return window->statusBar();
    QStatusBar* status=nullptr;
    for(auto* ancestor=document.parentWidget();ancestor;ancestor=ancestor->parentWidget())
        if(auto* window=qobject_cast<QMainWindow*>(ancestor)) { status=window->statusBar(); break; }
    if(!status) {
        status=document.findChild<QStatusBar*>("access-status-bar",Qt::FindDirectChildrenOnly);
        if(!status) {
            if(!document.layout()) throw VbaRuntimeError(3251,"Native status surface unavailable");
            status=new QStatusBar(&document); status->setObjectName("access-status-bar");
            status->setSizeGripEnabled(false); document.layout()->addWidget(status);
        }
    }
    return status;
}
inline VbaValue nativeAccessStatus(QWidget& document,const std::vector<VbaValue>& arguments) {
    auto count=arguments.size();
    while(count && arguments[count-1].kind==VbaValue::Kind::Missing) --count;
    if(!count || count>3) throw VbaRuntimeError(450,"SysCmd requires one to three arguments");
    const auto number=[](const VbaValue& value) {
        return static_cast<int>(coerceVbaModuleValue(scalarVbaValue(value),"long").number);
    };
    const auto action=number(arguments[0]);
    if(action<1 || action>5) throw VbaRuntimeError(3251,"SysCmd action is not implemented");
    const auto expected=action==1 ? 3U : action==2 || action==4 ? 2U : 1U;
    if(count!=expected) throw VbaRuntimeError(450,"Unexpected SysCmd action arguments");
    for(std::size_t i=0;i<count;++i)
        if(arguments[i].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"Required SysCmd argument omitted");
    QString text; int value=0;
    if(action==1 || action==4) {
        text=QString::fromStdString(coerceVbaModuleValue(scalarVbaValue(arguments[1]),"string").asText());
        if(text.isEmpty() || text.contains(QChar(0)) || text.contains('\n') || text.contains('\r'))
            throw VbaRuntimeError(5,"Invalid SysCmd status text");
    }
    if(action==1 || action==2) {
        value=number(arguments[action==1 ? 2 : 1]);
        if(value<(action==1 ? 1 : 0)) throw VbaRuntimeError(5,"Invalid SysCmd meter value");
    }
    auto* status=nativeStatusSurface(document);
    auto* meter=status->findChild<QProgressBar*>("access-syscmd-meter",Qt::FindDirectChildrenOnly);
    if(action==2) {
        if(!meter) throw VbaRuntimeError(3251,"SysCmd meter is not initialized");
        if(value>meter->maximum()) throw VbaRuntimeError(5,"SysCmd meter exceeds maximum");
        meter->setValue(value); return {};
    }
    if(action==1) {
        if(!meter) {
            meter=new QProgressBar(status); meter->setObjectName("access-syscmd-meter");
            meter->setMinimumWidth(160); status->addPermanentWidget(meter);
        }
        meter->setRange(0,value); meter->setValue(0); meter->show(); status->showMessage(text);
    } else if(action==3 || action==4) {
        if(meter) { status->removeWidget(meter); delete meter; }
        if(action==4) status->showMessage(text); else status->clearMessage();
    } else status->clearMessage();
    return {};
}
} // namespace linaccess
