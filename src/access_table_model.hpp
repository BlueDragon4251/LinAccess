#pragma once
#include <QSqlTableModel>
#include <QMap>
#include <QSet>
#include <QVector>
#include <QPersistentModelIndex>
namespace linaccess {
class AccessTableModel final : public QSqlTableModel {
    Q_OBJECT
public:
    explicit AccessTableModel(QObject* parent,QSqlDatabase database);
    void setTable(const QString& table) override;
    bool select() override;
    void clear() override;
    void revertRow(int row) override;
    bool insertRows(int row,int count,const QModelIndex& parent={}) override;
    bool removeRows(int row,int count,const QModelIndex& parent={}) override;
    bool submitTransactionally();
    bool hasPendingRecord(int row) const;
    bool undoPendingRecord(int row);
    QMap<QString,QVariant> submittedPrimaryKey(int originalRow) const { return submittedKeys_.value(originalRow); }
    bool canSubmitCurrentRecordAlone(int row) const;
    bool rejectEdit(int number,const QString& message);
    bool isTextColumn(int column) const;
    bool isSearchableColumn(int column) const;
    QString fieldSqlType(int column) const;
    bool setReadOnly(bool value);
    bool isReadOnly() const { return readOnly_; }
    void sort(int column,Qt::SortOrder order=Qt::AscendingOrder) override;
    void protectNextSort() { protectSort_=true; }
    QVariant data(const QModelIndex& index,int role=Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;
    bool setData(const QModelIndex& index,const QVariant& value,int role=Qt::EditRole) override;
    QString clipboardText(const QModelIndex& index) const;
    bool repeatPreviousValue(const QModelIndex& target);
    bool setDataBlock(const QModelIndex& first,const QVector<QVector<QString>>& values);
signals:
    void editRejected(int number,const QString& message);
    void sortRestored(int column,Qt::SortOrder order);
protected:
    bool insertRowIntoTable(const QSqlRecord& record) override;
private:
    bool refreshTypes();
    bool reject(int number);
    QString columnType(int column) const;
    QVector<QPersistentModelIndex> rowAnchors(const QSet<int>& rows) const;
    QSet<int> anchoredRows(const QVector<QPersistentModelIndex>& anchors) const;
    QMap<QString,QString> types_;
    QSet<QString> generated_;
    int metadataError_{};
    int sortColumn_{-1};
    Qt::SortOrder sortOrder_{Qt::AscendingOrder};
    bool protectSort_{};
    QSet<int> dirtyRows_,deletedRows_;
    QMap<QString,QVariant> insertedKey_;
    QMap<int,QMap<QString,QVariant>> submittedKeys_;
    bool changing_{},saving_{},readOnly_{};
};
}
