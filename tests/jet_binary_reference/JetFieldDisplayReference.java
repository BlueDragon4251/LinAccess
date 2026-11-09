// Optional zero-row Jet4 field-property reference; development tooling only.
// No business records, VBA or generated MDB is committed or shipped.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
class JetFieldDisplayReference {
 public static void main(String[] args) throws Exception {
  Path path=Path.of(args[0]).resolve("field-display-probe.mdb");if(Files.exists(path)) throw new IllegalStateException("Refusing replacement");
  try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
   new TableBuilder("property_probe")
    .addColumn(new ColumnBuilder("named",DataType.TEXT).setLengthInUnits(20).putProperty("Caption","Spaltenkopf \"Ä\"").putProperty("Description","Technischer Hinweis\nUnicode Ä"))
    .addColumn(new ColumnBuilder("uncaptioned",DataType.LONG))
    .addColumn(new ColumnBuilder("empty_caption",DataType.TEXT).setLengthInUnits(20).putProperty("Caption",""))
    .toTable(db);
  }
  try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
   Table table=db.getTable("property_probe");if(table.getRowCount()!=0 || !"Spaltenkopf \"Ä\"".equals(table.getColumn("named").getProperties().getValue("Caption")) || !"Technischer Hinweis\nUnicode Ä".equals(table.getColumn("named").getProperties().getValue("Description")) || table.getColumn("uncaptioned").getProperties().getValue("Caption")!=null || !"".equals(table.getColumn("empty_caption").getProperties().getValue("Caption"))) throw new IllegalStateException("Independent readback failed");
  }
  System.out.println("Technical Jet4 field display properties independently verified; zero records");
 }
}
