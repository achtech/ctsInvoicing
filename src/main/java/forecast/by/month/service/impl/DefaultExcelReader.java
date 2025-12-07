package forecast.by.month.service.impl;

import forecast.by.month.service.ExcelReader;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class DefaultExcelReader implements ExcelReader {
    @Override
    public Sheet getSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet '" + sheetName + "' not found in the input file.");
        }
        return sheet;
    }
}
