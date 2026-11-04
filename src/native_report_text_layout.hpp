#pragma once
#include <QRectF>
#include <QVector>
#include <optional>

// Line rectangles use widget coordinates and the same layout as native painting.
// A renderer may split only outside their vertical interiors.
class NativeReportTextLayout {
public:
    virtual ~NativeReportTextLayout() = default;
    virtual QVector<QRectF> reportTextLines() const = 0;
    virtual void setReportTextClip(std::optional<QRect> clip) = 0;
};
