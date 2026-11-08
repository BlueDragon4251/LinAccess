// Optional actual Jet4 stored-counter-state reference; no business records.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
import java.security.*;
import java.io.*;
import java.util.*;
class JetCounterStateReference {
    static byte[] digest(Path path) throws Exception {
        MessageDigest hash=MessageDigest.getInstance("SHA-256");
        try(InputStream in=Files.newInputStream(path)) { byte[] b=new byte[65536]; int n; while((n=in.read(b))!=-1) hash.update(b,0,n); }
        return hash.digest();
    }
    public static void main(String[] args) throws Exception {
        Path root=Path.of(args[0]),source=root.resolve("counter-state-probe.mdb"),copy=root.resolve("independent-next-copy.mdb");
        if(Files.exists(source)||Files.exists(copy)) throw new IllegalStateException("Refusing to replace fixture");
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,source.toFile())) {
            Table nativeTable=new TableBuilder("native_counter").addColumn(new ColumnBuilder("id",DataType.LONG).setAutoNumber(true))
                .addIndex(new IndexBuilder("CounterPrimary").addColumns("id").setPrimaryKey()).toTable(db);
            Table managed=new TableBuilder("managed_counter").addColumn(new ColumnBuilder("id",DataType.LONG).setAutoNumber(true)).toTable(db);
            for(Table table:new Table[]{nativeTable,managed}) {
                for(int i=0;i<3;++i) table.addRow(Column.AUTO_NUMBER);
                List<Row> remove=new ArrayList<>(); for(Row row:table) if(table==managed || ((Number)row.get("id")).intValue()==3) remove.add(row);
                for(Row row:remove) table.deleteRow(row);
            }
        }
        byte[] before=digest(source); Files.copy(source,copy);
        try(Database db=new DatabaseBuilder(copy.toFile()).open()) {
            for(String name:new String[]{"native_counter","managed_counter"}) {
                Object[] added=db.getTable(name).addRow(Column.AUTO_NUMBER);
                if(((Number)added[0]).intValue()!=4) throw new IllegalStateException("Independent next counter mismatch");
            }
        }
        try(Database db=new DatabaseBuilder(source.toFile()).setReadOnly(true).open()) {
            if(db.getTable("native_counter").getRowCount()!=2 || db.getTable("managed_counter").getRowCount()!=0) throw new IllegalStateException("Source rows changed");
        }
        if(!Arrays.equals(before,digest(source))) throw new IllegalStateException("Source hash changed");
        System.out.println("Technical Jet4 counters: deleted highest/all rows, independent next values4 and unchanged source verified");
    }
}
