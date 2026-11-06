// Optional real technical Jet4 schema; no business rows or app Java dependency.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;

class JetIndexReference {
    public static void main(String[] args) throws Exception {
        Path path=Path.of(args[0]).resolve("index-order-probe.mdb");
        if(Files.exists(path)) throw new IllegalStateException("Refusing to replace fixture");
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
            new TableBuilder("index_order_probe")
                .addColumn(new ColumnBuilder("first_key",DataType.LONG))
                .addColumn(new ColumnBuilder("second_key",DataType.LONG))
                .addIndex(new IndexBuilder("mixed_order").addColumns(false,"first_key").addColumns(true,"second_key").setUnique())
                .addIndex(new IndexBuilder("ascending_order").addColumns(true,"first_key"))
                .toTable(db);
        }
        try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
            Table table=db.getTable("index_order_probe");
            Index mixed=table.getIndex("mixed_order"),ascending=table.getIndex("ascending_order");
            if(table.getRowCount()!=0 || table.getIndexes().size()!=2 || !mixed.isUnique()
                || mixed.getColumns().size()!=2 || mixed.getColumns().get(0).isAscending()
                || !mixed.getColumns().get(1).isAscending() || !ascending.getColumns().get(0).isAscending())
                throw new IllegalStateException("Independent index readback mismatch");
        }
        System.out.println("Technical Jet4 index reference: zero rows, two named indexes, mixed directions independently verified");
    }
}
