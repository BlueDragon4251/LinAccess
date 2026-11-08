// Optional actual Jet4 Double precision/reference values; no business data.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
import java.util.*;
class JetDoublePrecisionReference {
    static void write(Path path,Double[] values) throws Exception {
        if(Files.exists(path)) throw new IllegalStateException("Refusing to replace fixture");
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
            Table table=new TableBuilder("double_precision_probe").addColumn(new ColumnBuilder("id",DataType.LONG))
                .addColumn(new ColumnBuilder("sample",DataType.DOUBLE)).toTable(db);
            for(int i=0;i<values.length;++i) table.addRow(i+1,values[i]);
        }
        int count=0;
        try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
            for(Row row:db.getTable("double_precision_probe")) {
                int id=((Number)row.get("id")).intValue();
                if(id<1 || id>values.length || !Objects.equals((Double)row.get("sample"),values[id-1])) throw new IllegalStateException("Double bits mismatch");
                ++count;
            }
        }
        if(count!=values.length) throw new IllegalStateException("Wrong row count");
    }
    public static void main(String[] args) throws Exception {
        Path root=Path.of(args[0]);
        write(root.resolve("double-precision-probe.mdb"),new Double[]{Double.longBitsToDouble(0x3ff0000000000001L),0.1d,Double.MIN_VALUE,Double.MAX_VALUE,-123.456d,0.0d,9007199254740992d,null});
        write(root.resolve("double-negative-zero.mdb"),new Double[]{-0.0d});
        System.out.println("Technical Jet4 Double reference: precision counterexample, finite extremes, NULL and negative-zero bits independently verified");
    }
}
