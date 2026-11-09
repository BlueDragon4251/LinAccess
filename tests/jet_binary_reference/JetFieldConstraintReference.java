// Development-only zero-row Jet4 property reference; no business records or VBA.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
class JetFieldConstraintReference {
 public static void main(String[] args) throws Exception {
  Path path=Path.of(args[0]).resolve("field-constraints-probe.mdb");if(Files.exists(path)) throw new IllegalStateException("Refusing replacement");
  try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
   new TableBuilder("constraint_probe")
    .addColumn(new ColumnBuilder("text_required",DataType.TEXT).setLengthInUnits(20).putProperty("Required",DataType.BOOLEAN,true).putProperty("AllowZeroLength",DataType.BOOLEAN,false))
    .addColumn(new ColumnBuilder("text_optional",DataType.TEXT).setLengthInUnits(20).putProperty("Required",DataType.BOOLEAN,false).putProperty("AllowZeroLength",DataType.BOOLEAN,true))
    .addColumn(new ColumnBuilder("memo",DataType.MEMO).putProperty("Required",DataType.BOOLEAN,false).putProperty("AllowZeroLength",DataType.BOOLEAN,false))
    .toTable(db);
  }
  try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
   Table table=db.getTable("constraint_probe");if(table.getRowCount()!=0) throw new IllegalStateException("Expected zero rows");
   for(String name:new String[]{"text_required","text_optional","memo"}) {
    PropertyMap props=table.getColumn(name).getProperties();
    if(!Boolean.valueOf(name.equals("text_required")).equals(props.getValue("Required")) || !Boolean.valueOf(name.equals("text_optional")).equals(props.getValue("AllowZeroLength"))) throw new IllegalStateException("Independent property readback failed");
   }
  }
  System.out.println("Technical Jet4 Required/AllowZeroLength independently verified; zero records");
 }
}
