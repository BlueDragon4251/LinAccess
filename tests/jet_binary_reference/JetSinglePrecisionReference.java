// Optional actual Jet4 Float precision and refusal references; no business rows.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
import java.util.*;

class JetSinglePrecisionReference {
    static void write(Path path,Float[] values) throws Exception {
        if(Files.exists(path)) throw new IllegalStateException("Refusing to replace fixture");
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
            Table table=new TableBuilder("single_precision_probe")
                .addColumn(new ColumnBuilder("id",DataType.LONG))
                .addColumn(new ColumnBuilder("sample",DataType.FLOAT)).toTable(db);
            for(int i=0;i<values.length;++i) table.addRow(i+1,values[i]);
        }
        int count=0;
        try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
            for(Row row:db.getTable("single_precision_probe")) {
                int id=((Number)row.get("id")).intValue(); Float actual=(Float)row.get("sample");
                if(id<1 || id>values.length || !Objects.equals(actual,values[id-1])) throw new IllegalStateException("Independent precision readback mismatch");
                ++count;
            }
        }
        if(count!=values.length) throw new IllegalStateException("Wrong row count");
    }
    public static void main(String[] args) throws Exception {
        Path root=Path.of(args[0]);
        write(root.resolve("single-precision-probe.mdb"),new Float[]{Float.intBitsToFloat(509022793),0.1f,Float.MIN_VALUE,Float.MAX_VALUE,-123.456f,0.0f,16777216f,null});
        write(root.resolve("single-negative-zero.mdb"),new Float[]{-0.0f});
        System.out.println("Technical Jet4 precision references: decimal-loss counterexample, finite extremes, NULL and negative-zero bits independently verified");
    }
}
