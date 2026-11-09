#pragma once
#include "native_macro_set_value_contract.hpp"
#include <functional>
namespace linaccess {
inline void nativeMacroSetValue(VbaHost& host,const std::vector<VbaValue>& args,const std::function<void()>& check) {
    check();if(args.size()!=2) throw VbaRuntimeError(450,"SetValue requires Item and Expression");const auto target=nativeMacroSetValueTarget(QString::fromStdString(args[0].asText()));const auto expression=nativeMacroSetValueExpression(QString::fromStdString(args[1].asText()));
    const auto reference=evaluateVbaObjectExpression(target.object.toUtf8().toStdString(),host,{}, {},nullptr,VbaCompareMode::Database);check();
    if(reference.kind!=VbaValue::Kind::Object || !reference.object) throw VbaRuntimeError(3251,"SetValue requires an actual native control or property target");
    const bool control=(reference.object->isNativeType("access.control") || reference.object->isNativeType("access.formfield"));if(!control && (!target.property || (!reference.object->isNativeType("access.form") && !reference.object->isNativeType("access.report")))) throw VbaRuntimeError(3251,"SetValue target type is not implemented");
    (void)reference.object->member("Name",{});check();
    auto evaluatedExpression=expression;
    if(expression.startsWith('=')) {
        const auto first=scalarVbaValue(host.eval(expression.mid(1).toUtf8().toStdString()));check();
        if(first.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Null SetValue expression text");
        if(first.kind==VbaValue::Kind::Error) throw VbaRuntimeError(first.errorNumber,"SetValue expression text is an Error value");
        evaluatedExpression=nativeMacroSetValueExpression(QString::fromStdString(first.asText()));
        if(evaluatedExpression.startsWith('=')) throw VbaRuntimeError(3251,"Repeated SetValue expression prefix is unsupported");
    }
    const auto value=host.eval(evaluatedExpression.toUtf8().toStdString());check();reference.object->assign(target.property ? target.property->toUtf8().toStdString() : "Value",value);check();
}
}
