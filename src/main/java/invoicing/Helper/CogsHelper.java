package invoicing.Helper;
import java.math.BigDecimal;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import invoicing.entities.CogsRecord;
import invoicing.enums.FiscalYear;

import java.io.InputStream;
import java.util.ArrayList;

public class CogsHelper {
    public static List<String> findGroupIdsByRate(
            BigDecimal rate,
            FiscalYear fiscalYear,
            List<CogsRecord> records) {

        if (rate == null || fiscalYear == null) {
            return List.of();
        }

        int inputWholePart = rate.intValue(); // whole number part

        return records.stream()
                .filter(record -> {

                    BigDecimal value = switch (fiscalYear) {
                        case FY25 -> record.getFy25();
                        case FY26 -> record.getFy26();
                    };

                    return value != null &&
                            value.intValue() == inputWholePart;
                })
                .map(CogsRecord::getGroupId)
                .collect(Collectors.toList());
    }

    public static List<CogsRecord> loadFromResources() throws Exception {

        List<CogsRecord> records = new ArrayList<>();

        try (InputStream is = CogsHelper.class
                .getClassLoader()
                .getResourceAsStream("Data.xlsx");
             Workbook workbook = new XSSFWorkbook(is)) {

            if (is == null) {
                throw new RuntimeException("Data.xlsx not found in resources folder");
            }

            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // skip header
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String groupId = formatter.formatCellValue(row.getCell(0));

                BigDecimal fy25 = parseDecimal(formatter.formatCellValue(row.getCell(1)));
                BigDecimal fy26 = parseDecimal(formatter.formatCellValue(row.getCell(2)));

                if (groupId != null && !groupId.isBlank()) {
                    records.add(new CogsRecord(groupId, fy25, fy26));
                }
            }
        }

        return records;
    }

    private static BigDecimal parseDecimal(String value) {
        if (value == null || value.isBlank()) return null;

        value = value.replace(",", "."); // handle comma decimals
        return new BigDecimal(value);
    }
}
