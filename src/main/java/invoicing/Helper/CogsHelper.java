package invoicing.Helper;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
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

        BigDecimal normalizedInput = rate.setScale(2, RoundingMode.HALF_UP);
        List<String> selectedYearMatches = findExactMatchesByYear(normalizedInput, fiscalYear, records);
        if (!selectedYearMatches.isEmpty()) {
            return selectedYearMatches;
        }

        FiscalYear fallbackYear = fiscalYear == FiscalYear.FY25 ? FiscalYear.FY26 : FiscalYear.FY25;
        List<String> fallbackYearMatches = findExactMatchesByYear(normalizedInput, fallbackYear, records);
        if (!fallbackYearMatches.isEmpty()) {
            return fallbackYearMatches;
        }

        // Last resort: nearest match on selected year with a safe cap to avoid very broad categories.
        List<String> nearest = findNearestMatchesByYear(rate, fiscalYear, records, new BigDecimal("0.60"));
        if (!nearest.isEmpty()) {
            return nearest;
        }

        return findNearestMatchesByYear(rate, fallbackYear, records, new BigDecimal("0.60"));
    }

    private static List<String> findExactMatchesByYear(BigDecimal normalizedInput, FiscalYear year, List<CogsRecord> records) {
        return records.stream()
                .filter(record -> {
                    BigDecimal value = getRateByYear(record, year);
                    if (value == null) {
                        return false;
                    }
                    BigDecimal normalizedValue = value.setScale(2, RoundingMode.HALF_UP);
                    return normalizedValue.compareTo(normalizedInput) == 0;
                })
                .map(CogsRecord::getGroupId)
                .distinct()
                .collect(Collectors.toList());
    }

    private static List<String> findNearestMatchesByYear(
            BigDecimal inputRate,
            FiscalYear year,
            List<CogsRecord> records,
            BigDecimal maxAllowedDistance) {
        BigDecimal minDistance = null;
        List<String> nearestGroups = new ArrayList<>();

        for (CogsRecord record : records) {
            BigDecimal value = getRateByYear(record, year);
            if (value == null) {
                continue;
            }

            BigDecimal distance = value.subtract(inputRate).abs();
            if (distance.compareTo(maxAllowedDistance) > 0) {
                continue;
            }

            if (minDistance == null || distance.compareTo(minDistance) < 0) {
                minDistance = distance;
                nearestGroups.clear();
                nearestGroups.add(record.getGroupId());
            } else if (distance.compareTo(minDistance) == 0) {
                nearestGroups.add(record.getGroupId());
            }
        }

        return nearestGroups.stream().distinct().collect(Collectors.toList());
    }

    private static BigDecimal getRateByYear(CogsRecord record, FiscalYear year) {
        return switch (year) {
            case FY25 -> record.getFy25();
            case FY26 -> record.getFy26();
        };
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
