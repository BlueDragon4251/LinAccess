#pragma once
#include <QObject>
#include <QPointer>
#include <QJsonObject>
#include <QByteArray>
#include <QVector>
#include <QTimer>
#include <functional>
class QDialog;
class QBoxLayout;
class QPushButton;
class QLabel;
namespace linaccess {
// In-memory authoring drafts only. Undo/redo never calls the database save path.
class NativeDesignerHistory final : public QObject {
public:
    using Capture=std::function<QJsonObject()>;
    using Restore=std::function<void(const QJsonObject&)>;
    NativeDesignerHistory(QDialog& dialog,QBoxLayout& buttons,Capture capture,Restore restore);
    void record();
    void markSaved();
private:
    void flush();
    void move(int direction);
    void update();
    void disable();
    QPointer<QDialog> dialog_;
    QPushButton *undo_{},*redo_{};
    QLabel* status_{};
    Capture capture_;
    Restore restore_;
    QTimer timer_;
    QVector<QByteArray> states_;
    QByteArray saved_;
    qsizetype cursor_{},bytes_{};
    bool fresh_{},restoring_{},available_{true};
};
}
