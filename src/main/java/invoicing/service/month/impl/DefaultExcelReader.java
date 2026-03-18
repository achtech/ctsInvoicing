package invoicing.service.month.impl;


import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import invoicing.service.month.ExcelReader;

import java.util.Locale;

public class DefaultExcelReader implements ExcelReader {
    @Override
    public Sheet getSheet(Workbook workbook, String sheetName) {
        if (workbook == null || sheetName == null) return null;

        Sheet exact = workbook.getSheet(sheetName);
        if (exact != null) return exact;

        String expected = normalize(sheetName);
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet s = workbook.getSheetAt(i);
            if (s == null) continue;
            if (expected.equals(normalize(s.getSheetName()))) return s;
        }

        return null;
    }

    private String normalize(String s) {
        String v = s.trim().toLowerCase(Locale.ROOT);
        v = v.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u");
        v = v.replace("à", "a").replace("è", "e").replace("ì", "i").replace("ò", "o").replace("ù", "u");
        v = v.replaceAll("\\s+", " ");
        return v;
    }
}
