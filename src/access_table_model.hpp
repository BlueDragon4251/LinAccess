#pragma once
#include <QSqlTableModel>
#include <QMap>
#include <QSet>
namespace linaccess {
class AccessTableModel final : public QSqlTableModel {
    Q_OBJECT
public:
    explicit AccessTableModel(QObject* parent,QSqlDatabase database);
    void setTable(const QString& table) override;
    bool select() override;
    void sort(int column,Qt::SortOrder order=Qt::AscendingOrder) override;
    void protectNextSort() { protectSort_=true; }
    QVariant data(const QModelIndex& index,int role=Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;
    bool setData(const QModelIndex& index,const QVariant& value,int role=Qt::EditRole) override;
signals:
    void editRejected(int number,const QString& message);
    void sortRestored(int column,Qt::SortOrder order);
protected:
    bool insertRowIntoTable(const QSqlRecord& record) override;
private:
    bool refreshTypes();
    bool reject(int number);
    QString columnType(int column) const;
    QMap<QString,QString> types_;
    QSet<QString> generated_;
    int metadataError_{};
    int sortColumn_{-1};
    Qt::SortOrder sortOrder_{Qt::AscendingOrder};
    bool protectSort_{};
};
}
