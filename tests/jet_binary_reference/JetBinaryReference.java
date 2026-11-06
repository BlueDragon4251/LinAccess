// Optional technical Jet fixtures only: no business records or runtime Java dependency.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
import java.util.*;

class JetBinaryReference {
    public static void main(String[] args) throws Exception {
        Path root=Path.of(args[0]); Path source=root.resolve("binary-probe.mdb");
        if(Files.exists(source)) throw new IllegalStateException("Refusing to replace fixture");
        Map<Integer,byte[]> values=new LinkedHashMap<>();
        values.put(1,new byte[]{0,(byte)255,66,34,44,10,13}); values.put(2,null);
        for(int size:new int[]{4097,65539}) {
            byte[] payload=new byte[size]; new Random(size).nextBytes(payload); values.put(values.size()+1,payload);
        }
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,source.toFile())) {
            Table table=new TableBuilder("binary_probe").addColumn(new ColumnBuilder("id",DataType.LONG))
                .addColumn(new ColumnBuilder("payload",DataType.OLE)).toTable(db);
            for(Map.Entry<Integer,byte[]> entry:values.entrySet()) table.addRow(entry.getKey(),entry.getValue());
        }
        int rows=0;
        try(Database db=new DatabaseBuilder(source.toFile()).setReadOnly(true).open()) {
            for(Row row:db.getTable("binary_probe")) {
                int id=((Number)row.get("id")).intValue(); byte[] observed=(byte[])row.get("payload");
                if(!values.containsKey(id) || !Arrays.equals(values.get(id),observed)) throw new IllegalStateException("Independent fixture readback mismatch");
                if(observed!=null) Files.write(root.resolve("payload-"+id+".bin"),observed); ++rows;
            }
        }
        if(rows!=4) throw new IllegalStateException("Wrong fixture row count");
        System.out.println("Technical Jet4 reference: 4 rows; NULL and 7/4097/65539 binary bytes independently verified");
        JetBinaryOversize.main(args);
    }
}

class JetBinaryOversize {
    public static void main(String[] args) throws Exception {
        Path source=Path.of(args[0]).resolve("binary-oversize.mdb");
        if(Files.exists(source)) throw new IllegalStateException("Refusing to replace fixture");
        byte[] payload=new byte[11*1024*1024]; new Random(11001).nextBytes(payload);
        try(Database db=DatabaseBuilder.create(Database.FileFormat.V2000,source.toFile())) {
            new TableBuilder("binary_oversize").addColumn(new ColumnBuilder("id",DataType.LONG))
                .addColumn(new ColumnBuilder("payload",DataType.OLE)).toTable(db).addRow(1,payload);
        }
        try(Database db=new DatabaseBuilder(source.toFile()).setReadOnly(true).open()) {
            Row row=db.getTable("binary_oversize").iterator().next();
            if(!Arrays.equals(payload,(byte[])row.get("payload"))) throw new IllegalStateException("Oversize readback mismatch");
        }
        System.out.println("Technical oversized Jet4 reference: 11MiB binary independently verified");
    }
}
