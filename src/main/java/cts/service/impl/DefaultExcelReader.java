package cts.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cts.service.ExcelReader;

public class DefaultExcelReader implements ExcelReader{
    @Override
    public Workbook readExcelFile(String fileName) throws IOException {
        File file = new File(fileName);
        if (!file.exists()) {
            throw new FileNotFoundException("Input file not found: " + fileName);
        }
        return new XSSFWorkbook(new FileInputStream(file));
    }

    @Override
    public Sheet getSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet '" + sheetName + "' not found in the input file.");
        }
        return sheet;
    }
}
