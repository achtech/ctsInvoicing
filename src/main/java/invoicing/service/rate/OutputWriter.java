package invoicing.service.rate;

// OutputWriter.java
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Map;

import invoicing.Helper.Helper;
import invoicing.Helper.GroupAggregator;
import invoicing.Helper.ReferenceData;
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
            
            // Create header row with new column names
            Row headerRow = sheet.createRow(0);
            Cell c1 = headerRow.createCell(0); c1.setCellValue("Client"); c1.setCellStyle(headerStyle);
            Cell c2 = headerRow.createCell(1); c2.setCellValue("Project"); c2.setCellStyle(headerStyle);
            Cell c3 = headerRow.createCell(2); c3.setCellValue("GroupId"); c3.setCellStyle(headerStyle);
            Cell c4 = headerRow.createCell(3); c4.setCellValue("Rate"); c4.setCellStyle(headerStyle);
            Cell c5 = headerRow.createCell(4); c5.setCellValue("Hours"); c5.setCellStyle(headerStyle);
            Cell c6 = headerRow.createCell(5); c6.setCellValue("Cost"); c6.setCellStyle(headerStyle);

            int rowNum = 1;
            double totalCost = 0;
            
            for (Map.Entry<String, Map<String, Double>> groupEntry : aggregator.getAggregates().entrySet()) {
                String groupId = groupEntry.getKey();
                BigDecimal rate = referenceData.getRateByGroup(groupId);

                if (rate == null) {
                    System.err.println("Warning: No rate found for GroupId: " + groupId);
                    continue;
                }

                // Calculate total hours for this group
                double hours = 0;
                for (Double h : groupEntry.getValue().values()) {
                    hours += h;
                }

                // Get total cost for this group
                double cost = 0;
                Map<String, Double> usersCost = aggregator.getCostAggregates().get(groupId);
                if (usersCost != null) {
                    for (Double c : usersCost.values()) {
                        cost += c;
                    }
                }
                cost = Helper.round(cost);
                
                totalCost += cost;

                // Create data row
                Row dataRow = sheet.createRow(rowNum++);
                Cell c7 = dataRow.createCell(0); c7.setCellValue("Italy"); c7.setCellStyle(bodyStyle);
                Cell c8 = dataRow.createCell(1); c8.setCellValue("INS-026696-00003"); c8.setCellStyle(bodyStyle);
                Cell c9 = dataRow.createCell(2); c9.setCellValue(groupId); c9.setCellStyle(bodyStyle);
                Cell c10 = dataRow.createCell(3); c10.setCellValue(rate.doubleValue()); c10.setCellStyle(currencyStyle);
                Cell c11 = dataRow.createCell(4); c11.setCellValue(hours); c11.setCellStyle(bodyStyle);
                Cell c12 = dataRow.createCell(5); c12.setCellValue(cost); c12.setCellStyle(currencyStyle);
            }
            
            totalCost = Helper.round(totalCost);
            // Create total row
            Row totalRow = sheet.createRow(rowNum);
            Cell c13 = totalRow.createCell(4); c13.setCellValue("Total"); c13.setCellStyle(headerStyle);
            Cell c14 = totalRow.createCell(5); c14.setCellValue(totalCost); c14.setCellStyle(currencyStyle);

            // Set column widths
            for (int i = 0; i < 6; i++) {
                sheet.setColumnWidth(i, 4500);
            }
            
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                wb.write(fos);
            }
        }
    }

    private CellStyle getBodyStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private CellStyle getHeaderStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private CellStyle getCurrencyStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        DataFormat format = wb.createDataFormat();
        style.setDataFormat(format.getFormat("#,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
}
