package rate;

/// InputRowProcessor.java
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.AbstractMap;
import java.util.Map;

public class InputRowProcessor {

    private final ReferenceData referenceData;

    public InputRowProcessor(ReferenceData referenceData) {
        this.referenceData = referenceData;
    }

    /**
     * Processes one row from input file and returns the identified Grupo and corrected Horas.
     * Returns null if employee cannot be identified.
     */
    public Map.Entry<String, Double> processRow(Row row) {
        if (row == null) return null;

        String employeeNumber = getStringCellValue(row.getCell(0));  // Column A
        String nameCellContent = getStringCellValue(row.getCell(1)); // Column B (may be partial)
        double reportedHoras = getDoubleCellValue(row.getCell(3));   // Column D - Horas

        double reportedRate = getDoubleCellValue(row.getCell(6));    // Column G - Rate
        double facturation = getDoubleCellValue(row.getCell(7)); // Column H

        String group = null;
        boolean identifiedByNameOnly = false;

        // 1. Best case: Employee Number exists
        if (isNotBlank(employeeNumber)) {
            group = referenceData.getGroupByNumber(employeeNumber.trim());
        }

        // 2. No number → try name matching (exact first, then "starts with")
        if (group == null && isNotBlank(nameCellContent)) {
            String trimmedName = nameCellContent.trim();

            // Exact match
            group = referenceData.getGroupByName(trimmedName);

            // Partial match: "starts with" in either direction
            if (group == null) {
                group = referenceData.findGroupByPartialName(trimmedName);
            }

            if (group != null) {
                identifiedByNameOnly = true;
            }
        }

        // 3. Last resort: infer group from reported rate
        if (group == null && reportedRate > 0) {
            group = referenceData.getGroupByRate(reportedRate);
        }
        if(group == null && reportedHoras == 0 && facturation!=0){
            group = "Tools";
        }
        // If still no group → skip this row
        if (group == null || group.isBlank()) {
            return null;
        }

        Double correctRate = referenceData.getRateByGroup(group);
        if (correctRate == null || correctRate < 0) {
            return null; // Invalid group/rate
        }

        double finalHoras = reportedHoras;

        // 4. CORRECT HOURS: only when identified by name (no number) AND rate mismatch
        // This preserves the monetary value: Horas × Rate = constant
        if (identifiedByNameOnly
                && (employeeNumber == null || employeeNumber.isBlank())
                && Math.abs(reportedRate - correctRate) > 0.001 // tolerance
                && reportedRate > 0) {

            finalHoras = (reportedHoras * reportedRate) / correctRate;
        }

        if("Tools".equals(group))
            finalHoras = facturation;
        return new AbstractMap.SimpleEntry<>(group.trim(), finalHoras);
    }

    // ──────── Helper Methods ────────

    private static boolean isNotBlank(String str) {
        return str != null && !str.trim().isEmpty();
    }

    private static String getStringCellValue(Cell cell) {
        if (cell == null) return null;

        if (cell.getCellType() == CellType.STRING) {
            String value = cell.getStringCellValue();
            return value != null ? value.trim() : null;
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            double num = cell.getNumericCellValue();
            if (num == (long) num) {
                return String.valueOf((long) num);
            }
            return String.valueOf(num);
        }
        return null;
    }

    private static double getDoubleCellValue(Cell cell) {
        if (cell == null) return 0.0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue().trim());
            } catch (Exception e) {
                return 0.0;
            }
        }
        return 0.0;
    }
}