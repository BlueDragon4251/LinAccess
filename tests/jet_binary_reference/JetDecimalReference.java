// Independent actual Jet4 Numeric reference; no business records or VBA.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
import java.math.BigDecimal;
import java.util.*;
class JetDecimalReference {
 public static void main(String[] args) throws Exception {
  Path root=Path.of(args[0]);Files.createDirectories(root);Path file=root.resolve("decimal-probe.mdb");
  if(Files.exists(file)) throw new IllegalStateException("Refusing to replace fixture");
  String[] large={"999999999999999999999999.9999","-999999999999999999999999.9999","9007199254740993.0001","0",null};
  String[] small={"0.0000000000000000000000000001","-0.0000000000000000000000000001","0.1234567890123456789012345678","0",null};
  try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,file.toFile())) {
   Table table=new TableBuilder("decimal_probe").addColumn(new ColumnBuilder("id",DataType.LONG))
    .addColumn(new ColumnBuilder("large",DataType.NUMERIC).setPrecision(28).setScale(4))
    .addColumn(new ColumnBuilder("small",DataType.NUMERIC).setPrecision(28).setScale(28)).toTable(db);
   for(int i=0;i<large.length;++i) table.addRow(i+1,large[i]==null?null:new BigDecimal(large[i]),small[i]==null?null:new BigDecimal(small[i]));
  }
  int count=0;
  try(Database db=new DatabaseBuilder(file.toFile()).setReadOnly(true).open()) {
   Table table=db.getTable("decimal_probe");if(table.getColumn("large").getPrecision()!=28 || table.getColumn("large").getScale()!=4 || table.getColumn("small").getScale()!=28) throw new IllegalStateException("Precision/scale mismatch");
   for(Row row:table) {int i=((Number)row.get("id")).intValue()-1;
    for(String name:new String[]{"large","small"}) {String expected=(name.equals("large")?large:small)[i];BigDecimal actual=(BigDecimal)row.get(name);if(expected==null ? actual!=null : actual==null || actual.compareTo(new BigDecimal(expected))!=0) throw new IllegalStateException("Exact Numeric mismatch");}
    ++count;
   }
  }
  if(count!=large.length) throw new IllegalStateException("Row count mismatch");
  System.out.println("Independent Jet4 Numeric values, precision/scale and NULL verified");
 }
}
