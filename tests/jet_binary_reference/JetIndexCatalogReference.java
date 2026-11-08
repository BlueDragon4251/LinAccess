// Optional zero-row Jet4 reference for genuine primary names and directions.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;

class JetIndexCatalogReference {
    public static void main(String[] args) throws Exception {
        Path path=Path.of(args[0]).resolve("index-catalog-probe.mdb");
        if(Files.exists(path)) throw new IllegalStateException("Refusing to replace fixture");
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
            new TableBuilder("catalog_probe")
                .addColumn(new ColumnBuilder("primary_a",DataType.LONG))
                .addColumn(new ColumnBuilder("primary_b",DataType.LONG))
                .addColumn(new ColumnBuilder("other_key",DataType.LONG))
                .addIndex(new IndexBuilder("DeliberatelyNamedPrimary").addColumns(false,"primary_a").addColumns(true,"primary_b").setPrimaryKey())
                .addIndex(new IndexBuilder("ActualSecondary").addColumns(true,"other_key").addColumns(false,"primary_b"))
                .toTable(db);
            new TableBuilder("single_primary_probe")
                .addColumn(new ColumnBuilder("primary_key",DataType.LONG))
                .addIndex(new IndexBuilder("NamedDescendingSingle").addColumns(false,"primary_key").setPrimaryKey())
                .toTable(db);
        }
        try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
            Table table=db.getTable("catalog_probe");
            Index primary=table.getIndex("DeliberatelyNamedPrimary"),secondary=table.getIndex("ActualSecondary");
            if(table.getRowCount()!=0 || table.getIndexes().size()!=2 || !primary.isPrimaryKey() || !primary.isUnique()
                || primary.getColumns().get(0).isAscending() || !primary.getColumns().get(1).isAscending()
                || !secondary.getColumns().get(0).isAscending() || secondary.getColumns().get(1).isAscending())
                throw new IllegalStateException("Independent primary index readback mismatch");
            Index single=db.getTable("single_primary_probe").getIndex("NamedDescendingSingle");
            if(!single.isPrimaryKey() || single.getColumns().size()!=1 || single.getColumns().get(0).isAscending())
                throw new IllegalStateException("Single descending primary mismatch");
        }
        System.out.println("Technical Jet4 catalog reference: original primary name and composite directions independently verified");
    }
}
