#pragma once
#include <QSqlTableModel>
#include <QMap>
#include <QSet>
#include <QVector>
#include <QPersistentModelIndex>
#include <QStringList>
#include <optional>
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
    bool replaceAllWholeText(int column,const QString& search,const QString& replacement,Qt::CaseSensitivity sensitivity,int* count);
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
    QVariant headerData(int section,Qt::Orientation orientation,int role=Qt::DisplayRole) const override;
    QVariant data(const QModelIndex& index,int role=Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;
    bool setData(const QModelIndex& index,const QVariant& value,int role=Qt::EditRole) override;
    QString clipboardText(const QModelIndex& index) const;
    QString formattedText(const QModelIndex& index) const;
    bool repeatPreviousValue(const QModelIndex& target);
    bool setDataBlock(const QModelIndex& first,const QVector<QVector<QString>>& values);
    bool setDataBlock(int firstRow,int firstColumn,const QVector<QVector<QString>>& values);
    bool isPasteColumn(int column,bool binary=false) const;
    bool setTypedDataBlock(int firstRow,int firstColumn,const QVector<QVector<QVariant>>& values);
    bool setTypedRecordBlock(int firstRow,const QStringList& fields,const QVector<QVector<QVariant>>& values);
    bool canPasteOverUpdates() const;
    bool canUndoBlock() const;
    bool canRedoBlock() const;
    bool undoBlock();
    bool redoBlock();
signals:
    void editRejected(int number,const QString& message);
    void sortRestored(int column,Qt::SortOrder order);
protected:
    bool insertRowIntoTable(const QSqlRecord& record) override;
private:
    bool applyDataBlock(int firstRow,int firstColumn,const QVector<QVector<QVariant>>& values,bool typed,const QMap<QString,QVariant>* restoredDefaults=nullptr,bool skipGenerated=false);
    bool refreshTypes();
    bool reject(int number);
    QString columnType(int column) const;
    QVector<QPersistentModelIndex> rowAnchors(const QSet<int>& rows) const;
    QSet<int> anchoredRows(const QVector<QPersistentModelIndex>& anchors) const;
    QMap<QString,QString> types_,captions_,descriptions_,formats_;
    QSet<QString> generated_;
    int metadataError_{};
    int sortColumn_{-1};
    Qt::SortOrder sortOrder_{Qt::AscendingOrder};
    bool protectSort_{};
    QSet<int> dirtyRows_,deletedRows_;
    QMap<QString,QVariant> insertedKey_;
    QMap<int,QMap<QString,QVariant>> submittedKeys_;
    struct CachedUpdate {enum Kind {Update,Insert,Delete};int row{};Kind kind{Update};QVector<int> columns;QVector<QVariant> values;};
    QVector<CachedUpdate> captureUpdates() const;
    bool restoreUpdates(const QVector<CachedUpdate>& updates,int originalRows);
    struct BlockHistory {
        int firstRow{},firstColumn{},originalRows{};bool applied{},skipGenerated{};
        QVector<QVector<QVariant>> values;QMap<QString,QVariant> dateDefaults;QVector<CachedUpdate> before;
    };
    std::optional<BlockHistory> blockHistory_;
    bool blockMutation_{};
    bool changing_{},saving_{},readOnly_{},binaryWriting_{};
};
}
