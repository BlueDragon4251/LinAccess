// Independent technical Jet4 and ACE2010 MONEY fixtures. No application data or VBA.
import com.healthmarketscience.jackcess.ColumnBuilder;
import com.healthmarketscience.jackcess.DataType;
import com.healthmarketscience.jackcess.Database;
import com.healthmarketscience.jackcess.DatabaseBuilder;
import com.healthmarketscience.jackcess.Row;
import com.healthmarketscience.jackcess.Table;
import com.healthmarketscience.jackcess.TableBuilder;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

class JetCurrencyReference {
    private static final String TABLE = "currency_probe";
    private static final String DEFAULT = "1.25";
    private static final Long[] TICKS = {
        Long.MAX_VALUE, Long.MIN_VALUE, 9007199254740993L, 1L, -1L, 0L, 12500L, null
    };

    private static void writeAndVerify(Path file, Database.FileFormat format, boolean variableSlot) throws Exception {
        if (Files.exists(file)) {
            throw new IllegalStateException("Refusing to replace fixture " + file);
        }
        try (Database db = DatabaseBuilder.create(format, file.toFile())) {
            Table table = new TableBuilder(TABLE)
                .addColumn(new ColumnBuilder("id", DataType.LONG))
                .addColumn(new ColumnBuilder("amount", DataType.MONEY) {
                    @Override
                    public boolean isVariableLength() {
                        return variableSlot;
                    }
                })
                .toTable(db);
            for (int i = 0; i < TICKS.length; ++i) {
                Long ticks = TICKS[i];
                table.addRow(i + 1, ticks == null ? null : BigDecimal.valueOf(ticks, 4));
            }
            // Jackcess applies an existing default to null inputs when adding rows.
            var properties = table.getColumn("amount").getProperties();
            properties.put("DefaultValue", DataType.TEXT, DEFAULT);
            properties.save();
        }

        try (Database db = new DatabaseBuilder(file.toFile()).setReadOnly(true).open()) {
            if (db.getFileFormat() != format) {
                throw new IllegalStateException("Unexpected fixture format");
            }
            Table table = db.getTable(TABLE);
            var amount = table.getColumn("amount");
            if (amount.getType() != DataType.MONEY || amount.isVariableLength() != variableSlot ||
                    amount.getLength() != Long.BYTES || !DEFAULT.equals(amount.getProperties().getValue("DefaultValue"))) {
                throw new IllegalStateException("MONEY field, layout or default mismatch");
            }
            boolean[] seen = new boolean[TICKS.length];
            int count = 0;
            for (Row row : table) {
                int index = row.getInt("id") - 1;
                if (index < 0 || index >= TICKS.length || seen[index]) {
                    throw new IllegalStateException("Unexpected reference row");
                }
                seen[index] = true;
                BigDecimal value = row.getBigDecimal("amount");
                Long ticks = value == null ? null : value.movePointRight(4).longValueExact();
                if (!Objects.equals(ticks, TICKS[index])) {
                    throw new IllegalStateException("MONEY ticks mismatch for row " + (index + 1));
                }
                ++count;
            }
            if (count != TICKS.length) {
                throw new IllegalStateException("MONEY reference row count mismatch");
            }
        }
        System.out.println("Independently reopened MONEY reference: " + format + " " + file);
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 1) {
            throw new IllegalArgumentException("Output directory required");
        }
        Path directory = Files.createDirectories(Path.of(args[0]));
        // Jet4 can store an eight-byte MONEY value in a variable row slot.
        writeAndVerify(directory.resolve("currency-probe.mdb"), Database.FileFormat.V2000, true);
        writeAndVerify(directory.resolve("currency-probe.accdb"), Database.FileFormat.V2010, false);
    }
}
