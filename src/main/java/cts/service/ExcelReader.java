package cts.service;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelReader {
	 Workbook readExcelFile(String fileName) throws IOException;
	    Sheet getSheet(Workbook workbook, String sheetName);
}
