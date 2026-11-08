// Optional independent read-only Double oracle for supplied test copies.
// Writes private technical bit tuples only; never logs source names/values/VBA.
import com.healthmarketscience.jackcess.*;
import java.nio.file.*;
import java.nio.charset.StandardCharsets;
import java.security.*;
import java.io.*;
import java.util.*;

class JetDoubleOracle {
    static String hex(byte[] value) { return HexFormat.of().formatHex(value); }
    static String digest(Path path) throws Exception {
        MessageDigest hash=MessageDigest.getInstance("SHA-256");
        try(InputStream in=Files.newInputStream(path)) { byte[] buffer=new byte[65536]; int count; while((count=in.read(buffer))!=-1) hash.update(buffer,0,count); }
        return hex(hash.digest());
    }
    public static void main(String[] args) {
        try { run(args); }
        catch(Throwable error) { System.err.println("Independent Double source oracle could not be verified"); System.exit(3); }
    }
    static void run(String[] args) throws Exception {
        if(args.length!=2) throw new IllegalArgumentException();
        Path source=Path.of(args[0]); String before=digest(source); Path root=Path.of(args[1]).resolve(before);
        Files.createDirectory(root);
        int tableCount=0,columns=0,nonNull=0,nulls=0,negativeZeros=0,nonfinite=0;
        try(Database db=new DatabaseBuilder(source.toFile()).setReadOnly(true).open()) {
            for(String name:db.getTableNames()) {
                Table table=db.getTable(name); List<String> fields=new ArrayList<>();
                for(Column column:table.getColumns()) if(column.getType()==DataType.DOUBLE) fields.add(column.getName());
                if(fields.isEmpty()) continue;
                List<String> tuples=new ArrayList<>();
                for(Row row:table) {
                    List<String> parts=new ArrayList<>();
                    for(String field:fields) {
                        Double value=(Double)row.get(field);
                        if(value==null) { parts.add("N"); ++nulls; }
                        else {
                            long bits=Double.doubleToRawLongBits(value); parts.add(String.format(Locale.ROOT,"%016x",bits)); ++nonNull;
                            if(bits==Long.MIN_VALUE) ++negativeZeros;
                            if(!Double.isFinite(value)) ++nonfinite;
                        }
                    }
                    tuples.add(String.join("\t",parts));
                }
                Collections.sort(tuples);
                List<String> lines=new ArrayList<>(); List<String> header=new ArrayList<>();
                for(String field:fields) header.add(hex(field.getBytes(StandardCharsets.UTF_8)));
                lines.add("COLUMNS\t"+String.join("\t",header)); lines.addAll(tuples);
                String filename=hex(MessageDigest.getInstance("SHA-256").digest(name.getBytes(StandardCharsets.UTF_8)))+".rows";
                Files.write(root.resolve(filename),lines,StandardCharsets.UTF_8,StandardOpenOption.CREATE_NEW);
                ++tableCount; columns+=fields.size();
            }
        }
        if(!digest(source).equals(before)) throw new IllegalStateException();
        Files.writeString(root.resolve("SOURCE_SHA256"),before+"\n",StandardOpenOption.CREATE_NEW);
        System.out.println("Independent Double oracle: tables="+tableCount+", columns="+columns+", nonNULL="+nonNull+", NULL="+nulls+", negativeZero="+negativeZeros+", nonfinite="+nonfinite+"; source hash unchanged");
    }
}
