package invoicing.Helper;

import org.apache.poi.ss.usermodel.*;

public class ExcelStyler {

    public static void applyStyles(Workbook workbook) {

        Sheet sheet = workbook.getSheetAt(0);

        // === HEADER STYLE ===
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.BLACK.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        // === BODY STYLE ===
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderRight(BorderStyle.THIN);

        // Apply header style (row 0)
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                cell.setCellStyle(headerStyle);
            }
        }

        // Iterate through rows
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            for (int c = 0; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                // Convert column D (index 3) to number
                if (c == 3 && cell.getCellType() == CellType.STRING) {
                    String text = cell.getStringCellValue().trim();
                    try {
                        double num = Double.parseDouble(text);
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(num);
                    } catch (Exception ignore) {}
                }

                cell.setCellStyle(bodyStyle);
            }
        }

        // Auto-size columns
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
