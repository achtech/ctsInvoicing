package forecast.by.rate.util;
// ReferenceData.java
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReferenceData {
    private Map<String, Double> groupToRate = new HashMap<>();
    private Map<String, String> employeeNumberToGroup = new HashMap<>();
    private Map<String, String> employeeNameToGroup = new HashMap<>();

    public void load(String path) throws IOException {
        try (FileInputStream fis = new FileInputStream(path);
             Workbook wb = new XSSFWorkbook(fis)) {
            Sheet sheet1 = wb.getSheetAt(0); // Rate Reference Table
            for (Row row : sheet1) {
                if (row.getRowNum() == 0) continue;
                String group = getStringCellValue(row.getCell(0));
                double rate = getDoubleCellValue(row.getCell(1));
                if (group != null) {
                    groupToRate.put(group, rate);
                }
            }

            Sheet sheet2 = wb.getSheetAt(1); // Employee Reference Table
            for (Row row : sheet2) {
                if (row.getRowNum() == 0) continue;
                String number = getStringCellValue(row.getCell(0));
                String name = getStringCellValue(row.getCell(1));
                String group = getStringCellValue(row.getCell(2));
                if (number != null && group != null) {
                    employeeNumberToGroup.put(number, group);
                }
                if (name != null && group != null) {
                    employeeNameToGroup.put(name, group);
                }
            }
            
            // Add Hardcoded defaults for Guardias
            groupToRate.put("Guardias-50", 50.0);
            groupToRate.put("Guardias-25", 25.0);
            employeeNameToGroup.put("Guardias-50", "Guardias-50");
            employeeNameToGroup.put("Guardias-25", "Guardias-25");
            
        }
    }

    public Map<String, String> getEmployeeNameToGroupMap() {
        return new HashMap<>(employeeNameToGroup); // defensive copy
    }

    public String getGroupByNumber(String number) {
        return employeeNumberToGroup.get(number);
    }

    public String getGroupByName(String name) {
        return employeeNameToGroup.get(name);
    }

    public Double getRateByGroup(String group) {
        return groupToRate.get(group);
    }

    public String getGroupByRate(double rate) {
        double tolerance = 0.001;
        for (Map.Entry<String, Double> entry : groupToRate.entrySet()) {
            if (Math.abs(entry.getValue() - rate) < tolerance) {
                return entry.getKey();
            }
        }
        return null;
    }

    private static String getStringCellValue(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((int) cell.getNumericCellValue());
        }
        return null;
    }

    private static double getDoubleCellValue(Cell cell) {
        if (cell == null) return 0.0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        return 0.0;
    }

    // Add this to ReferenceData.java
    public String findGroupByPartialName(String namePart) {
        if (namePart == null || namePart.trim().isEmpty()) return null;

        String search = namePart.trim().toLowerCase();

        // Try exact match first (case-insensitive)
        for (Map.Entry<String, String> entry : employeeNameToGroup.entrySet()) {
            String knownName = entry.getKey();
            if (knownName != null && knownName.trim().equalsIgnoreCase(namePart.trim())) {
                return entry.getValue();
            }
        }

        // Then "starts with" in either direction
        for (Map.Entry<String, String> entry : employeeNameToGroup.entrySet()) {
            String knownName = entry.getKey();
            if (knownName != null) {
                String knownLower = knownName.trim().toLowerCase();
                if (knownLower.startsWith(search) || search.startsWith(knownLower)) {
                    return entry.getValue();
                }
            }
        }
        return null;
    }
}