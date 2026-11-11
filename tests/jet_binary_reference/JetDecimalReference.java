// Independent actual Jet4/ACE2010 Numeric reference; no business records or VBA.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
import java.math.BigDecimal;
import java.util.*;
class JetDecimalReference {
 public static void main(String[] args) throws Exception {
  if(args.length<1 || args.length>2) throw new IllegalArgumentException("Output directory and optional V2000/V2010 format required");
  Database.FileFormat format=args.length==2?Database.FileFormat.valueOf(args[1]):Database.FileFormat.V2000;
  if(format!=Database.FileFormat.V2000 && format!=Database.FileFormat.V2010) throw new IllegalArgumentException("Only independently checked V2000 and V2010 formats are supported");
  Path root=Path.of(args[0]);Files.createDirectories(root);Path file=root.resolve(format==Database.FileFormat.V2000?"decimal-probe.mdb":"decimal-probe.accdb");
  if(Files.exists(file)) throw new IllegalStateException("Refusing to replace fixture");
  String[] large={"999999999999999999999999.9999","-999999999999999999999999.9999","9007199254740993.0001","0",null};
  String[] small={"0.0000000000000000000000000001","-0.0000000000000000000000000001","0.1234567890123456789012345678","0",null};
  try(Database db=DatabaseBuilder.create(format,file.toFile())) {
   Table table=new TableBuilder("decimal_probe").addColumn(new ColumnBuilder("id",DataType.LONG))
    .addColumn(new ColumnBuilder("large",DataType.NUMERIC).setPrecision(28).setScale(4))
    .addColumn(new ColumnBuilder("small",DataType.NUMERIC).setPrecision(28).setScale(28)).toTable(db);
   for(int i=0;i<large.length;++i) table.addRow(i+1,large[i]==null?null:new BigDecimal(large[i]),small[i]==null?null:new BigDecimal(small[i]));
   // Set defaults after explicit NULL reference rows: Jackcess applies defaults
   // to null insert inputs when the property already exists.
   table.getColumn("large").getProperties().put("DefaultValue",DataType.TEXT,"9007199254740993.0001");table.getColumn("large").getProperties().save();
   table.getColumn("small").getProperties().put("DefaultValue",DataType.TEXT,"0.0000000000000000000000000001");table.getColumn("small").getProperties().save();
  }
  int count=0;
  try(Database db=new DatabaseBuilder(file.toFile()).setReadOnly(true).open()) {
   Table table=db.getTable("decimal_probe");if(table.getColumn("large").getPrecision()!=28 || table.getColumn("large").getScale()!=4 || table.getColumn("small").getScale()!=28) throw new IllegalStateException("Precision/scale mismatch");
   if(!"9007199254740993.0001".equals(table.getColumn("large").getProperties().getValue("DefaultValue")) || !"0.0000000000000000000000000001".equals(table.getColumn("small").getProperties().getValue("DefaultValue"))) throw new IllegalStateException("Exact default property mismatch");
   for(Row row:table) {int i=((Number)row.get("id")).intValue()-1;
    for(String name:new String[]{"large","small"}) {String expected=(name.equals("large")?large:small)[i];BigDecimal actual=(BigDecimal)row.get(name);if(expected==null ? actual!=null : actual==null || actual.compareTo(new BigDecimal(expected))!=0) throw new IllegalStateException("Exact Numeric mismatch");}
    ++count;
   }
  }
  if(count!=large.length) throw new IllegalStateException("Row count mismatch");
  System.out.println("Independent Numeric values, precision/scale, literal defaults and NULL verified: "+format);
 }
}
