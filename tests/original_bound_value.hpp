#pragma once
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_format.hpp"
#include "vba_qt_value.hpp"
#include <QLineEdit>
#include <cmath>
#include <limits>
#include <iostream>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
inline void verifyOriginalBoundText(AccessFormPreview& form,const linaccess::AccessVisualControl& control,QLineEdit* widget,const QVariant& raw,const QString& sqlType = {}) {
    assert(widget->property("la_null").toBool()==raw.isNull());
    auto expected=linaccess::qtFieldValue(raw,sqlType);
    const auto actual=makePreviewFormObject(form)->item(control.name.toStdString());
    // This corpus checks record/value binding. SQLite expression metadata does
    // not establish complete Access Variant subtype identity.
    const auto numeric=[](const auto& value) {return value.kind==linaccess::VbaValue::Kind::Number || value.kind==linaccess::VbaValue::Kind::Currency || value.kind==linaccess::VbaValue::Kind::Decimal || value.kind==linaccess::VbaValue::Kind::Boolean || value.kind==linaccess::VbaValue::Kind::Date;};
    if(actual.kind==linaccess::VbaValue::Kind::Date && expected.kind==linaccess::VbaValue::Kind::Text) expected=linaccess::qtFieldValue(raw,"DATETIME");
    if(numeric(actual) && numeric(expected)) {const auto value=actual.asNumber(),reference=expected.asNumber();assert(std::isfinite(value) && std::isfinite(reference) && std::abs(value-reference)<=qMax(1e-9,std::abs(reference)*1e-12));}
    else {if(actual.asText()!=expected.asText()) std::cerr<<"Original bound-value mismatch: actual-kind="<<static_cast<int>(actual.kind)<<", expected-kind="<<static_cast<int>(expected.kind)<<", format-units="<<control.format.size()<<"\n";assert(actual.asText()==expected.asText() && (actual.kind==linaccess::VbaValue::Kind::Null)==raw.isNull());}
    if(control.format.isEmpty() || raw.isNull()) {assert(widget->text()==linaccess::qtFieldText(raw,sqlType));return;}
    if(widget->hasFocus() && !widget->isReadOnly()) {
        const auto editing=actual.kind==linaccess::VbaValue::Kind::Number ? QString::number(actual.number,'g',std::numeric_limits<double>::max_digits10) : linaccess::qtFieldText(raw,sqlType);
        assert(widget->text()==editing);return;
    }
    try {const auto display=linaccess::formatAccessValue(actual,control.format.toUtf8().toStdString(),true);assert(widget->text()==QString::fromStdString(display.asText()) && !widget->property("la_format_error").toInt());}
    catch(const linaccess::VbaRuntimeError& error) {assert(widget->text()==QString("#Fehler (%1)").arg(error.number) && widget->property("la_format_error").toInt()==error.number);}
}
