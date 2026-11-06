// Optional real Jet4 type reference with technical scalar values only.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;

class JetSingleReference {
    public static void main(String[] args) throws Exception {
        Path path=Path.of(args[0]).resolve("single-type-probe.mdb");
        if(Files.exists(path)) throw new IllegalStateException("Refusing to replace fixture");
        Float[] expected={0.1f,16777216f,1.5f,null};
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,path.toFile())) {
            Table table=new TableBuilder("single_type_probe")
                .addColumn(new ColumnBuilder("id",DataType.LONG))
                .addColumn(new ColumnBuilder("sample",DataType.FLOAT)).toTable(db);
            for(int i=0;i<expected.length;++i) table.addRow(i+1,expected[i]);
        }
        int count=0;
        try(Database db=new DatabaseBuilder(path.toFile()).setReadOnly(true).open()) {
            Table table=db.getTable("single_type_probe");
            if(table.getColumn("sample").getType()!=DataType.FLOAT) throw new IllegalStateException("Wrong Jet type");
            for(Row row:table) {
                int id=((Number)row.get("id")).intValue(); Float actual=(Float)row.get("sample");
                if(id<1 || id>expected.length || !java.util.Objects.equals(actual,expected[id-1]))
                    throw new IllegalStateException("Independent Single readback mismatch");
                ++count;
            }
        }
        if(count!=4) throw new IllegalStateException("Wrong row count");
        System.out.println("Technical Jet4 Single reference: three Float values and NULL independently verified");
    }
}
