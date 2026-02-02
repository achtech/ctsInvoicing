package forecast.by.rate.service;

/// InputRowProcessor.java
import forecast.by.rate.util.ReferenceData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class InputRowProcessor {

    private final ReferenceData referenceData;

    public InputRowProcessor(ReferenceData referenceData) {
        this.referenceData = referenceData;
    }

    /**
     * Processes one row from input file and returns the identified Grupo, user and corrected Horas.
     * Returns null if employee cannot be identified.
     */
    public RowData processRow(Row row) {
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
            return null;
        }

        double finalHoras = reportedHoras;

        if (identifiedByNameOnly
                && (employeeNumber == null || employeeNumber.isBlank())
                && Math.abs(reportedRate - correctRate) > 0.001
                && reportedRate > 0) {

            finalHoras = (reportedHoras * reportedRate) / correctRate;
        }

        if("Tools".equals(group))
            finalHoras = facturation;

        String user;
        if (isNotBlank(employeeNumber)) {
            user = employeeNumber.trim();
        } else if (isNotBlank(nameCellContent)) {
            user = nameCellContent.trim();
        } else {
            user = "";
        }

        return new RowData(group.trim(), user, finalHoras);
    }

    public static class RowData {
        private final String group;
        private final String user;
        private final double horas;

        public RowData(String group, String user, double horas) {
            this.group = group;
            this.user = user;
            this.horas = horas;
        }

        public String getGroup() {
            return group;
        }

        public String getUser() {
            return user;
        }

        public double getHoras() {
            return horas;
        }
    }


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
