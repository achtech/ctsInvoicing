package global;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class ExcelWriter {

    public void write(List<ServiceTeam> items, File targetFolder) throws Exception {

        Workbook workbook = new XSSFWorkbook();
        LocalDate today = LocalDate.now();
        String sheetName = today.format(DateTimeFormatter.ofPattern("MMMM yyyy"));

        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle bodyStyle = getBodyStyle(workbook);
        CellStyle headerStyle = getHeaderStyle(workbook);
        CellStyle currencyStyle = getCurrencyStyle(workbook);

        // -----------------------------------------------------
        // HEADER CREATION
        // -----------------------------------------------------
        Row header = sheet.createRow(0);
        Cell c1 =header.createCell(0);c1.setCellValue("Project Client");c1.setCellStyle(headerStyle);
        Cell c2 =header.createCell(1);c2.setCellValue("Project EXT");c2.setCellStyle(headerStyle);
        Cell c3 =header.createCell(2);c3.setCellValue("Descr EXT");c3.setCellStyle(headerStyle);
        Cell c4 =header.createCell(3);c4.setCellValue("Total EUR");c4.setCellStyle(headerStyle);
        Cell c5 =header.createCell(4);c5.setCellValue("BU");c5.setCellStyle(headerStyle);

        // -----------------------------------------------------
        // WRITE ROW DATA
        // -----------------------------------------------------
        for (int i = 0; i < items.size(); i++) {
            ServiceTeam st = items.get(i);
            Row row = sheet.createRow(i + 1);
            Cell c10 = row.createCell(0);c10.setCellValue(st.getProjectName());c10.setCellStyle(bodyStyle);
            Cell c11 = row.createCell(1);c11.setCellValue(st.getExtCode());c11.setCellStyle(bodyStyle);
            Cell c12 = row.createCell(2);c12.setCellValue(st.getProjectDescription());c12.setCellStyle(bodyStyle);
            Cell c13 = row.createCell(3);
            double costValue = Double.parseDouble(st.getCost());  // convert String → double
            c13.setCellValue(costValue);                          // write numeric value
            c13.setCellStyle(currencyStyle);
            Cell c14 = row.createCell(4);c14.setCellValue(st.getBuDescription());c14.setCellStyle(bodyStyle);
        }
        double grandTotal = items.stream()
                .mapToDouble(item -> Double.parseDouble(item.getCost()))
                .sum();
        int grandTotalRow = items.size()+1;
        Row footer = sheet.createRow(grandTotalRow);
        Cell c20 = footer.createCell(2);c20.setCellValue("Grand Total");c20.setCellStyle(headerStyle);
        Cell c21 = footer.createCell(3);c21.setCellValue(grandTotal);c21.setCellStyle(currencyStyle);

        int title2Row = grandTotalRow+3;
        Row title2 = sheet.createRow(title2Row);
        Cell c22 = title2.createCell(1);c22.setCellValue("Amount"); c22.setCellStyle(headerStyle);
        Cell c23 = title2.createCell(2);c23.setCellValue("OneErp details");c23.setCellStyle(headerStyle);

        int cogsRow = title2Row+1;
        Row cogs = sheet.createRow(cogsRow);
        Cell c24 = cogs.createCell(0);c24.setCellValue("COGS");c24.setCellStyle(bodyStyle);
        Cell c25 = cogs.createCell(1);c25.setCellValue(grandTotal);c25.setCellStyle(currencyStyle);
        Cell c26 = cogs.createCell(2);c26.setCellValue("Split on EXTs");c26.setCellStyle(bodyStyle);

        int gaRow = cogsRow+1;
        Row ga = sheet.createRow(gaRow);
        Cell c27 = ga.createCell(0);c27.setCellValue("G&A (10% COGS)");c27.setCellStyle(bodyStyle);
        Cell c28 = ga.createCell(1);c28.setCellValue(grandTotal*0.1);c28.setCellStyle(currencyStyle);
        Cell c29 = ga.createCell(2);c29.setCellValue("INT Code1");c29.setCellStyle(bodyStyle);

        int tpRow = gaRow+1;
        Row tp = sheet.createRow(tpRow);
        Cell c30 = tp.createCell(0);c30.setCellValue("TP (5% (COGS+G&A))");c30.setCellStyle(bodyStyle);
        Cell c31 = tp.createCell(1);c31.setCellValue(grandTotal*0.055);c31.setCellStyle(currencyStyle);
        Cell c32 = tp.createCell(2);c32.setCellValue("INT Code2");c32.setCellStyle(bodyStyle);

        int totalRow = tpRow+1;
        Row total = sheet.createRow(totalRow);
        Cell c33 = total.createCell(0);c33.setCellValue("Total Cost");c33.setCellStyle(bodyStyle);
        Cell c34 = total.createCell(1);c34.setCellValue(grandTotal*1.155);c34.setCellStyle(currencyStyle);

        for (int i = 0; i < 10; i++) {
            sheet.autoSizeColumn(i);
        }
        // -----------------------------------------------------
        // WRITE FILE
        // -----------------------------------------------------
        File file = new File(targetFolder, "ForeCast IT.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.close();
        workbook.close();
    }


    private CellStyle getCurrencyStyle(Workbook workbook){
        CellStyle currencyStyle = workbook.createCellStyle();
        currencyStyle.setDataFormat(
                workbook.createDataFormat().getFormat("#,##0.00€")
        );
        currencyStyle.setBorderBottom(BorderStyle.THIN);
        currencyStyle.setBorderTop(BorderStyle.THIN);
        currencyStyle.setBorderLeft(BorderStyle.THIN);
        currencyStyle.setBorderRight(BorderStyle.THIN);
        return currencyStyle;
    }

    private CellStyle getHeaderStyle(Workbook workbook){
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
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        return headerStyle;
    }

    private CellStyle getBodyStyle(Workbook workbook){
        // === BODY STYLE ===
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderRight(BorderStyle.THIN);
        return bodyStyle;
    }
}
