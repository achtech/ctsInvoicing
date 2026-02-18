package forecast.by.month.service.impl;

import forecast.by.month.service.ExcelReader;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class DefaultExcelReader implements ExcelReader {
    @Override
    public Sheet getSheet(Workbook workbook, String sheetName) {
        // Try direct match first
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet != null) return sheet;

        // Try normalized match (handle accents and case)
        String searchName = normalize(sheetName.toLowerCase());
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            String currentName = normalize(workbook.getSheetName(i).toLowerCase());
            if (currentName.equals(searchName)) {
                return workbook.getSheetAt(i);
            }
        }

        throw new IllegalArgumentException("Sheet '" + sheetName + "' not found in the input file.");
    }

    private String normalize(String str) {
        return str.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u");
    }
}
