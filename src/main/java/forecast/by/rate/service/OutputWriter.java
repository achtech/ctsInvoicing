package forecast.by.rate.service;

// OutputWriter.java
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import forecast.by.rate.util.GroupAggregator;
import forecast.by.rate.util.ReferenceData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OutputWriter {
    private final ReferenceData referenceData;
    private final GroupAggregator aggregator;

    public OutputWriter(ReferenceData referenceData, GroupAggregator aggregator) {
        this.referenceData = referenceData;
        this.aggregator = aggregator;
    }

    public void write(String outputPath) throws IOException {

        try (Workbook wb = new XSSFWorkbook()) {
            CellStyle bodyStyle = getBodyStyle(wb);
            CellStyle headerStyle = getHeaderStyle(wb);
            CellStyle currencyStyle = getCurrencyStyle(wb);

            Sheet sheet = wb.createSheet("Consolidated");
            Row headerRow = sheet.createRow(0);
            Cell c1 = headerRow.createCell(0);c1.setCellValue("Cliente");c1.setCellStyle(headerStyle);
            Cell c2 = headerRow.createCell(1);c2.setCellValue("Proyecto");c2.setCellStyle(headerStyle);
            Cell c3 = headerRow.createCell(2);c3.setCellValue("Grupo");c3.setCellStyle(headerStyle);
            Cell c4 = headerRow.createCell(3);c4.setCellValue("Rate");c4.setCellStyle(headerStyle);
            Cell c5 = headerRow.createCell(4);c5.setCellValue("Horas");c5.setCellStyle(headerStyle);
            Cell c6 = headerRow.createCell(5);c6.setCellValue("Facturacion");c6.setCellStyle(headerStyle);

            int rowNum = 1;
            double total = 0;
            for (Map.Entry<String, Double> entry : aggregator.getAggregates().entrySet()) {
                String group = entry.getKey();
                double horas = entry.getValue();
                Double rate = referenceData.getRateByGroup(group);

                if (rate == null) continue; // Skip if no rate
                double facturacion = !("Tools".equals(group)) ? rate * horas : horas;
                total+=facturacion;

                Row dataRow = sheet.createRow(rowNum++);
                Cell c7= dataRow.createCell(0);c7.setCellValue("Italy");c7.setCellStyle(bodyStyle);
                Cell c8= dataRow.createCell(1);c8.setCellValue("INS-026696-00003");c8.setCellStyle(bodyStyle);
                Cell c9= dataRow.createCell(2);c9.setCellValue(group);c9.setCellStyle(bodyStyle);
                Cell c10= dataRow.createCell(3);c10.setCellValue(rate);c10.setCellStyle(currencyStyle);
                Cell c11= dataRow.createCell(4);c11.setCellValue(horas);c11.setCellStyle(bodyStyle);
                Cell c12= dataRow.createCell(5);c12.setCellValue(facturacion);c12.setCellStyle(currencyStyle);
            }
            Row totalRow = sheet.createRow(rowNum++);
            Cell c13= totalRow.createCell(4);c13.setCellValue("Total");c13.setCellStyle(headerStyle);
            Cell c14= totalRow.createCell(5);c14.setCellValue(total);c14.setCellStyle(currencyStyle);

            for (int i = 0; i < 6; i++) {
                sheet.setColumnWidth(i, 4500);
            }
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                wb.write(fos);
            }
        }
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
        headerStyle.setFillForegroundColor(IndexedColors.BLUE1.getIndex());
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
