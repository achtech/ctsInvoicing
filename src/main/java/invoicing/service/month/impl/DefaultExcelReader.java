package invoicing.service.month.impl;


import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import invoicing.service.month.ExcelReader;

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
