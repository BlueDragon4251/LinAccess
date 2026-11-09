// Development-only Jet4 Format property fixture; zero rows, no business data/VBA.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
class JetFieldFormatReference {
 public static void main(String[] args) throws Exception {
  Path path=Path.of(args[0]).resolve("field-format-probe.mdb");if(Files.exists(path)) throw new IllegalStateException("Refusing replacement");
  try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
   new TableBuilder("format_probe")
    .addColumn(new ColumnBuilder("amount",DataType.DOUBLE).putProperty("Format","0.00"))
    .addColumn(new ColumnBuilder("moment",DataType.SHORT_DATE_TIME).putProperty("Format","yyyy-mm-dd"))
    .addColumn(new ColumnBuilder("plain",DataType.TEXT).setLengthInUnits(20))
    .toTable(db);
  }
  try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
   Table table=db.getTable("format_probe");
   if(table.getRowCount()!=0 || !"0.00".equals(table.getColumn("amount").getProperties().getValue("Format")) || !"yyyy-mm-dd".equals(table.getColumn("moment").getProperties().getValue("Format")) || table.getColumn("plain").getProperties().getValue("Format")!=null) throw new IllegalStateException("Independent format property readback failed");
  }
  System.out.println("Technical Jet4 Format properties independently verified; zero records");
 }
}
