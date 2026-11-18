// Optional technical Jet fixtures only: no business records or runtime Java dependency.
import com.healthmarketscience.jackcess.*;
import com.healthmarketscience.jackcess.util.OleBlob;
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
        JetOlePackageReference.main(args);
        JetBinaryOversize.main(args);
    }
}

class JetOlePackageReference {
    private static final String FILE_NAME = "r\u00e9sum\u00e9-\ud83d\udcc4.bin";
    private static final byte[] PAYLOAD = {0, (byte)0xff, 44, 34, 10, 13, 1, 5, 0};

    public static void main(String[] args) throws Exception {
        Path root = Path.of(args[0]);
        Path source = root.resolve("ole-package-probe.mdb");
        if (Files.exists(source)) {
            throw new IllegalStateException("Refusing to replace OLE Package fixture");
        }

        try (Database db = DatabaseBuilder.create(Database.FileFormat.V2000, source.toFile());
             OleBlob embedded = simplePackage(FILE_NAME, PAYLOAD);
             OleBlob empty = simplePackage("empty.bin", new byte[0]);
             OleBlob link = new OleBlob.Builder()
                 .setLinkFileName("external.bin")
                 .setLinkPath("C:\\ole\\external.bin")
                 .toBlob()) {
            Table table = new TableBuilder("ole_package_probe")
                .addColumn(new ColumnBuilder("id", DataType.LONG))
                .addColumn(new ColumnBuilder("payload", DataType.OLE)).toTable(db);
            table.addRow(1, embedded);
            table.addRow(2, empty);
            table.addRow(3, link);
        }

        Set<Integer> seen = new HashSet<>();
        try (Database db = new DatabaseBuilder(source.toFile()).setReadOnly(true).open()) {
            for (Row row : db.getTable("ole_package_probe")) {
                int id = row.getInt("id");
                if (!seen.add(id)) {
                    throw new IllegalStateException("Duplicate OLE Package fixture ID");
                }
                try (OleBlob blob = row.getBlob("payload")) {
                    OleBlob.Content content = blob.getContent();
                    if (id == 3 && content.getType() == OleBlob.ContentType.LINK) {
                        Files.write(root.resolve("ole-wrapper-3.bin"), blob.getBytes(1, (int)blob.length()));
                        continue;
                    }
                    String fileName = id == 1 ? FILE_NAME : "empty.bin";
                    byte[] expected = id == 1 ? PAYLOAD : new byte[0];
                    if ((id != 1 && id != 2) || !(content instanceof OleBlob.SimplePackageContent)) {
                        throw new IllegalStateException("Unexpected OLE Package fixture type or ID");
                    }
                    OleBlob.SimplePackageContent pkg = (OleBlob.SimplePackageContent)content;
                    try (var stream = pkg.getStream()) {
                        if (!fileName.equals(pkg.getFileName()) || !Arrays.equals(expected, stream.readAllBytes())) {
                            throw new IllegalStateException("Embedded package readback mismatch");
                        }
                    }
                    Files.write(root.resolve("ole-wrapper-" + id + ".bin"), blob.getBytes(1, (int)blob.length()));
                    Files.write(root.resolve("ole-content-" + id + ".bin"), expected);
                }
            }
        }
        if (!seen.equals(Set.of(1, 2, 3))) {
            throw new IllegalStateException("Missing OLE Package fixture rows");
        }
        System.out.println("Technical Jet4 OLE reference: Unicode embedded package, empty embedded file and external link independently verified");
    }

    private static OleBlob simplePackage(String fileName, byte[] payload) throws Exception {
        return new OleBlob.Builder()
            .setSimplePackageBytes(payload)
            .setSimplePackageFileName(fileName)
            .setSimplePackageFilePath("C:\\ole\\" + fileName)
            .toBlob();
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
