package cts.service;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelReader {
    Sheet getSheet(Workbook workbook, String sheetName);
}
