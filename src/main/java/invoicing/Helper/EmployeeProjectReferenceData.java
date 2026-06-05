package invoicing.Helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

public class EmployeeProjectReferenceData {
    private static final String FILE_NAME = "employee-categorys.xlsx";
    private static final String CATEGORY_SHEET_NAME = "Categorys";
    private static final String PROJECTS_SHEET_NAME = "Projects";

    private final Map<String, String> categoryByEmployeeNumber = new HashMap<>();
    private final Map<String, String> descriptionByExtCode = new HashMap<>();

    public static EmployeeProjectReferenceData load() throws IOException {
        EmployeeProjectReferenceData data = new EmployeeProjectReferenceData();
        try (InputStream is = EmployeeProjectReferenceData.class.getClassLoader().getResourceAsStream(FILE_NAME)) {
            if (is == null) {
                throw new IOException(FILE_NAME + " not found in resources");
            }
            data.load(is);
        }
        return data;
    }

    private void load(InputStream is) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(is)) {
            DataFormatter formatter = new DataFormatter(Locale.US);
            loadCategories(workbook, formatter);
            loadProjects(workbook, formatter);
        }
    }

    private void loadCategories(Workbook workbook, DataFormatter formatter) {
        Sheet sheet = workbook.getSheet(CATEGORY_SHEET_NAME);
        if (sheet == null) {
            return;
        }

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            String employeeNumber = normalizeEmployeeNumber(row.getCell(0), formatter);
            String category = formatter.formatCellValue(row.getCell(2)).trim();
            if (!employeeNumber.isEmpty() && !category.isEmpty()) {
                categoryByEmployeeNumber.put(employeeNumber, category);
            }
        }
    }

    private void loadProjects(Workbook workbook, DataFormatter formatter) {
        Sheet sheet = workbook.getSheet(PROJECTS_SHEET_NAME);
        if (sheet == null) {
            return;
        }

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            String extCode = normalizeExtCode(formatter.formatCellValue(row.getCell(1)));
            String description = formatter.formatCellValue(row.getCell(2)).trim();
            if (!extCode.isEmpty() && !description.isEmpty()) {
                descriptionByExtCode.put(extCode, description);
            }
        }
    }

    public String findCategory(Cell employeeNumberCell) {
        String employeeNumber = normalizeEmployeeNumber(employeeNumberCell, new DataFormatter(Locale.US));
        if (employeeNumber.isEmpty()) {
            return "";
        }
        return categoryByEmployeeNumber.getOrDefault(employeeNumber, "");
    }

    public String findProjectDescription(String extCode) {
        return descriptionByExtCode.getOrDefault(normalizeExtCode(extCode), "");
    }

    private static String normalizeEmployeeNumber(Cell cell, DataFormatter formatter) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return "";
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((long) cell.getNumericCellValue());
        }
        String value = formatter.formatCellValue(cell).trim();
        return value.endsWith(".0") ? value.substring(0, value.length() - 2) : value;
    }

    private static String normalizeExtCode(String value) {
        if (value == null) {
            return "";
        }
        return value.trim().replace('_', '-').toUpperCase(Locale.ROOT);
    }
}
