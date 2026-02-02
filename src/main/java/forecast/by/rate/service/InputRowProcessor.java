package forecast.by.rate.service;

/// InputRowProcessor.java
import forecast.by.rate.util.ReferenceData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class InputRowProcessor {

    private final ReferenceData referenceData;
    // Store both Rate and Hours for each user occurrence
    private final Map<String, List<RawDataEntry>> rawUserRates = new TreeMap<>();

    // Simple container for raw data
    private static class RawDataEntry {
        String user;
        String group;
        double rate;
        double hours;

        RawDataEntry(String user, String group, double rate, double hours) {
            this.user = user;
            this.group = group;
            this.rate = rate;
            this.hours = hours;
        }
    }

    public InputRowProcessor(ReferenceData referenceData) {
        this.referenceData = referenceData;
    }

    /**
     * Prints the raw rates collected from the Excel files, grouped by user and ordered by rate.
     */
    public void printRawRates() {
        System.out.println("\n=== RAW DATA FROM EXCEL (Ordered by Rate) ===");
        
        // Flatten the map to a list of aggregates
        Map<String, RawDataEntry> aggregatesMap = new TreeMap<>(); // Key: "User|Group|Rate"

        for (Map.Entry<String, List<RawDataEntry>> entry : rawUserRates.entrySet()) {
            String user = entry.getKey();
            List<RawDataEntry> entries = entry.getValue();
            
            for (RawDataEntry e : entries) {
                String key = user + "|" + (e.group != null ? e.group : "Unknown") + "|" + e.rate;
                aggregatesMap.compute(key, (k, v) -> {
                    if (v == null) return new RawDataEntry(user, e.group, e.rate, e.hours);
                    v.hours += e.hours;
                    return v;
                });
            }
        }

        List<RawDataEntry> sortedList = new ArrayList<>(aggregatesMap.values());
        // Sort by Rate (Ascending)
        sortedList.sort(Comparator.comparingDouble(e -> e.rate));

        for (RawDataEntry data : sortedList) {
            System.out.println("Group: " + (data.group != null ? data.group : "Unknown") + 
                               " | User: " + data.user + 
                               " | Rate: " + data.rate + 
                               " | hours : " + String.format("%.2f", data.hours));
        }
        System.out.println("=============================================\n");
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

        // Calculate Facturacion based on user logic:
        // 1. If Rate Input == Rate Table -> Revenue (Hours * Rate)
        // 2. If Rate Input != Rate Table -> Adjusted Hours (Hours * Rate Input / Rate Table)
        double finalFacturacion = 0.0;
        if ("Tools".equals(group)) {
             finalFacturacion = facturation;
        } else {
             // Use a small epsilon for double comparison
             if (Math.abs(reportedRate - correctRate) < 0.001) {
                 // Rates match: Revenue
                 finalFacturacion = reportedHoras * reportedRate;
             } else {
                 // Rates differ: Adjusted Hours (Hours * Input / Table)
                 if (correctRate != 0) {
                     finalFacturacion = (reportedHoras * reportedRate) / correctRate;
                 }
             }
        }

        String user;
        if (isNotBlank(employeeNumber)) {
            user = employeeNumber.trim();
        } else if (isNotBlank(nameCellContent)) {
            user = nameCellContent.trim();
        } else {
            user = "";
        }

        // Collect raw data for reporting (with resolved Group)
        String rawUserKey = (nameCellContent != null) ? nameCellContent : ((employeeNumber != null) ? employeeNumber : "Unknown");
        rawUserRates.computeIfAbsent(rawUserKey, k -> new ArrayList<>()).add(new RawDataEntry(rawUserKey, group, reportedRate, reportedHoras));

        return new RowData(group.trim(), user, finalHoras, finalFacturacion);
    }

    public static class RowData {
        private final String group;
        private final String user;
        private final double horas;
        private final double facturacion;

        public RowData(String group, String user, double horas, double facturacion) {
            this.group = group;
            this.user = user;
            this.horas = horas;
            this.facturacion = facturacion;
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

        public double getFacturacion() {
            return facturacion;
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
