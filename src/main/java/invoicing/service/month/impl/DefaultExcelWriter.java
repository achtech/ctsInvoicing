package invoicing.service.month.impl;


import invoicing.service.month.ExcelWriter;
import invoicing.Helper.CogsHelper;
import invoicing.entities.CogsRecord;
import invoicing.enums.FiscalYear;
import invoicing.Helper.Helper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class DefaultExcelWriter implements ExcelWriter {
    List<CogsRecord> recogs;

    public DefaultExcelWriter() throws Exception {
        recogs = CogsHelper.loadFromResources();
    }

    @Override
    public Workbook createWorkbookWithSheets(List<String> monthNames) {
        Workbook workbook = new XSSFWorkbook();
        for (String monthName : monthNames) {
            workbook.createSheet("Service Hours Details " + monthName);
        }
        return workbook;
    }

    private List<Row> getAdjustmentSheetData(Workbook inputWorkbook, String ajustesSheetName, String serviceTeam, int month) {
        Sheet inputSheet = inputWorkbook.getSheet(ajustesSheetName);
        List<Row> rows = new ArrayList<>();
        // Copy data rows where column E matches serviceTeam
        for (Row inputRow : inputSheet) {
            if (inputRow.getRowNum() == 0) {
                continue;
            }
            Cell cellE = inputRow.getCell(4); // Column E
            Cell cellH = inputRow.getCell(7); // Column H
            if (cellE.getStringCellValue().equals(serviceTeam) &&
                    cellH.getDateCellValue().getMonth() == (month - 1)) {
                rows.add(inputRow);
            }
        }
        return rows;
    }

    @Override
    public void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam,
                                          String invoicingSheetNameES, String invoicingSheetName, String ajustesSheetName, String facturacionSheetName) {
        Sheet outputSheet = outputWorkbook.getSheet(invoicingSheetName);
        if (outputSheet == null) {
            System.err.println("Skipping invoicing details sheet: output sheet not found.");
            return;
        }
        copyServiceHoursToSheet(inputWorkbook, outputSheet, 0, serviceTeam, invoicingSheetNameES, invoicingSheetName, ajustesSheetName, facturacionSheetName);
    }

    @Override
    public int copyServiceHoursToSheet(Workbook inputWorkbook, Sheet outputSheet, int startRowIndex, String serviceTeam,
                                       String invoicingSheetNameES, String invoicingSheetName, String ajustesSheetName, String facturacionSheetName) {
        Workbook outputWorkbook = outputSheet.getWorkbook();
        Sheet inputSheet = inputWorkbook.getSheet(invoicingSheetNameES);
        if (inputSheet == null) {
            System.err.println("Skipping invoicing details sheet: input sheet not found.");
            return startRowIndex;
        }

        int outputRowIndex = startRowIndex;

        // Create styles
        CellStyle headerStyle = Helper.getHeaderStyle(outputWorkbook);
        CellStyle centerStyle = Helper.getCenterStandardStyle(outputWorkbook);
        CellStyle leftStyle = Helper.getLeftStandardStyle(outputWorkbook);
        CellStyle currencyStyle = Helper.getCurrencyStyle(outputWorkbook);
        CellStyle dateStyle = Helper.getDateStyle(outputWorkbook);
        CellStyle vacanceStyle = Helper.getVacanceStyle(outputWorkbook);
        CellStyle freedayStyle = Helper.getFreedayStyle(outputWorkbook);
        CellStyle sickLeaveStyle = Helper.getSickLeaveStyle(outputWorkbook);
        CellStyle legalAbsenceStyle = Helper.getLegalAbsenceStyle(outputWorkbook);
        CellStyle weekendStyle = Helper.getWeekendStyle(outputWorkbook);
        CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(outputWorkbook);

        // Team Name Header (if starting a new table in a consolidated sheet)
        if (startRowIndex > 0) {
            outputRowIndex++; // Add an empty row before the new table
            Row teamHeaderRow = outputSheet.createRow(outputRowIndex++);
            Cell teamCell = teamHeaderRow.createCell(0);
            teamCell.setCellValue("PROJECT: " + serviceTeam + " (" + invoicingSheetName + ")");
            CellStyle teamStyle = outputWorkbook.createCellStyle();
            Font teamFont = outputWorkbook.createFont();
            teamFont.setBold(true);
            teamFont.setFontHeightInPoints((short) 14);
            teamStyle.setFont(teamFont);
            teamCell.setCellStyle(teamStyle);
        }

        // Create and set custom header row
        Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
        String[] headers = {"Empl. N°", "Person", "Category", "Rates"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = outputHeaderRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
        int nbrDaysInThisMonths = Helper.numberOfDays(invoicingSheetName);
        for (int i = headers.length, j = 1; i < headers.length + nbrDaysInThisMonths; i++, j++) {
            Cell cell = outputHeaderRow.createCell(i);
            cell.setCellValue(j);
            cell.setCellStyle(dateStyle);
        }
        Cell cell = outputHeaderRow.createCell(headers.length + nbrDaysInThisMonths);
        cell.setCellValue("Working Hours");
        cell.setCellStyle(headerStyle);

        Cell cellAmount = outputHeaderRow.createCell(headers.length + nbrDaysInThisMonths + 1);
        cellAmount.setCellValue("Cost (Euro) ");
        cellAmount.setCellStyle(headerStyle);

        int dataStartRow = outputRowIndex + 1;

        // Iterate through rows in the input sheet
        Map<BigDecimal, List<Row>> maps = getAllData(inputSheet);
        Map<BigDecimal, List<Row>> mapsByServiceTeam = filterRowsByServiceTeam(maps, serviceTeam);
        Map<BigDecimal, Row> mergedMaps = transformRows(inputWorkbook, facturacionSheetName,mapsByServiceTeam);
        
        // WRITE DATA IN EXCEL FILE
        for (Map.Entry<BigDecimal, Row> entry : mergedMaps.entrySet()) {
            Row row = entry.getValue();
            Row outputRow = outputSheet.createRow(outputRowIndex++);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell inputCell = row.getCell(j);
                Cell outputCell = outputRow.createCell(j);
                if (inputCell != null) {
                    switch (inputCell.getCellType()) {
                        case NUMERIC:
                            double numericValue = inputCell.getNumericCellValue();
                            if (j == 3) { // Rates column
                                outputCell.setCellValue(Helper.round(numericValue));
                                outputCell.setCellStyle(currencyStyle);
                            } else {
                                outputCell.setCellValue(numericValue);
                                outputCell.setCellStyle(centerStyle);
                            }
                            break;
                        case STRING:
                            String value = inputCell.getStringCellValue();
                            outputCell.setCellValue(inputCell.getStringCellValue());
                            if (j != 1)
                                outputCell.setCellStyle("V".equals(value) ? vacanceStyle
                                        : "A".equals(value) ? legalAbsenceStyle
                                        : "F".equals(value) ? freedayStyle
                                        : "S".equals(value) ? sickLeaveStyle : weekendStyle);
                            else
                                outputCell.setCellStyle(leftStyle);
                            break;
                        default:
                            break;
                    }
                }
                if (j == headers.length + nbrDaysInThisMonths)
                    break;
            }
        }

        // Add Cost column
        int lastColumn = headers.length + nbrDaysInThisMonths;
        String letterCostHours = Helper.getColumnLetter(lastColumn);
        String letterRate = "D";

        for (int r = dataStartRow - 1; r < outputRowIndex; r++) {
            Row row = outputSheet.getRow(r);
            if (row != null) {
                Cell newCell = row.createCell(lastColumn + 1);
                Cell rateCell = row.getCell(3);
                if (rateCell != null && CellType.NUMERIC.equals(rateCell.getCellType()) && rateCell.getNumericCellValue() != 0) {
                    String formula = "ROUND(" + letterRate + (row.getRowNum() + 1) + "*" + letterCostHours + (row.getRowNum() + 1) + ", 2)";
                    newCell.setCellFormula(formula);
                }
                newCell.setCellStyle(currencyStyle);
            }
        }

        // ADD ADJUSTMENT - FIXED SECTION
        int month = Helper.getMonthFromSheetName(invoicingSheetName);
        List<Row> AjustesRows = getAdjustmentSheetData(inputWorkbook, ajustesSheetName, serviceTeam, month);
        for (Row row : AjustesRows) {
            Row outputRow = outputSheet.createRow(outputRowIndex++);

            Cell cellA = outputRow.createCell(0);
            cellA.setCellValue("");
            cellA.setCellStyle(centerStyle);

            Cell cellB = outputRow.createCell(1);
            cellB.setCellValue(row.getCell(6).getStringCellValue());
            cellB.setCellStyle(leftStyle);

            Cell cellC = outputRow.createCell(2);
            cellC.setCellValue("");
            cellC.setCellStyle(centerStyle);

            Cell cellD = outputRow.createCell(3);
            BigDecimal hourlyRate = new BigDecimal(row.getCell(15).getNumericCellValue());
            cellD.setCellValue(Helper.round(hourlyRate.doubleValue()));
            cellD.setCellStyle(currencyStyle);

            Cell cellE = outputRow.createCell(4);
            cellE.setCellValue(row.getCell(12).getNumericCellValue());
            cellE.setCellStyle(centerStyle);

            for (int i = headers.length + 1; i < lastColumn; i++) {
                Cell cellM = outputRow.createCell(i);
                cellM.setCellValue("");
                cellM.setCellStyle(centerStyle);
            }

            Cell cellAdj = outputRow.createCell(lastColumn);
            BigDecimal workingHours = new BigDecimal(row.getCell(12).getNumericCellValue());
            cellAdj.setCellValue(workingHours.doubleValue());
            cellAdj.setCellStyle(centerStyle);

            Cell cellCost = outputRow.createCell(lastColumn + 1);
            if (workingHours.compareTo(BigDecimal.ZERO) == 0) {
                cellCost.setCellValue(Helper.round(row.getCell(16).getNumericCellValue()));
            } else {
                cellCost.setCellValue(Helper.round(workingHours.multiply(hourlyRate).doubleValue()));
            }
            cellCost.setCellStyle(currencyStyle);
        }

        // ADD TOTAL ROW FOR THIS TABLE
        Row totalRow = outputSheet.createRow(outputRowIndex++);
        Cell cellTotalLabel = totalRow.createCell(lastColumn - 2);
        cellTotalLabel.setCellValue("Total " + serviceTeam);
        cellTotalLabel.setCellStyle(headerStyle);
        outputSheet.addMergedRegion(new CellRangeAddress(totalRow.getRowNum(), totalRow.getRowNum(), lastColumn - 2, lastColumn - 1));

        Cell cellTotalHours = totalRow.createCell(lastColumn);
        String formulaHours = "SUM(" + letterCostHours + dataStartRow + ":" + letterCostHours + (outputRowIndex - 1) + ")";
        cellTotalHours.setCellFormula(formulaHours);
        cellTotalHours.setCellStyle(headerStyle);

        Cell cellTotalCost = totalRow.createCell(lastColumn + 1);
        String letterTotalCost = Helper.getColumnLetter(lastColumn + 1);
        String formulaCost = "SUM(" + letterTotalCost + dataStartRow + ":" + letterTotalCost + (outputRowIndex - 1) + ")";
        cellTotalCost.setCellFormula(formulaCost);
        cellTotalCost.setCellStyle(footerCurrencyStyle);

        // Auto-size
        for (int col = 0; col < lastColumn + 2; col++) {
            outputSheet.autoSizeColumn(col);
        }

        return outputRowIndex;
    }

    private Map<BigDecimal, List<Row>> getAllData(Sheet inputSheet) {

        Map<BigDecimal, List<Row>> maps = new HashMap<>();
        BigDecimal lastId = null;

        for (Row row : inputSheet) {

            Cell empIdCell = row.getCell(0);
            Cell empNameCell = row.getCell(1);

            if (empIdCell != null && empNameCell != null) {

                lastId = BigDecimal.valueOf(empIdCell.getNumericCellValue());

                List<Row> list = new ArrayList<>();
                list.add(row);

                maps.put(lastId, list);

            } else if (!Helper.isRowEmpty(row) && lastId != null) {

                List<Row> list = maps.get(lastId);

                if (list != null) {
                    list.add(row);
                }
            }
        }

        return maps;
    }


    private static Map<BigDecimal, List<Row>> filterRowsByServiceTeam(Map<BigDecimal, List<Row>> inputMap, String serviceTeam) {
        Map<BigDecimal, List<Row>> filteredMap = new HashMap<>();

        for (Map.Entry<BigDecimal, List<Row>> entry : inputMap.entrySet()) {
            BigDecimal key = entry.getKey();
            List<Row> rows = entry.getValue();

            // Extract the first row and filter rows based on the second cell condition
            List<Row> filteredRows = rows.stream().filter(row -> {
                Cell secondCell = row.getCell(1); // Second cell (index 1)
                boolean isFirstRow = rows.indexOf(row) == 0; // Check if it's the first row
                boolean isSecondCellEmptyOrMatches = secondCell == null || secondCell.getCellType() == CellType.BLANK
                        || (secondCell.getCellType() == CellType.STRING
                        && secondCell.getStringCellValue().contains(serviceTeam));
                return isFirstRow || isSecondCellEmptyOrMatches;
            }).collect(Collectors.toList());

            if (!filteredRows.isEmpty()) {
                filteredMap.put(key, filteredRows);
            }
        }

        return filteredMap;
    }

    private Map<BigDecimal, Row> transformRows(Workbook inputWorkbook,String sheetNameEs ,Map<BigDecimal, List<Row>> inputMap) {
        Map<BigDecimal, Row> resultMap = new HashMap<>();
        for (Map.Entry<BigDecimal, List<Row>> entry : inputMap.entrySet()) {
            BigDecimal key = entry.getKey();
            List<Row> rows = entry.getValue();

            if (rows == null || rows.isEmpty()) {
                continue; // Skip if the list is empty or null
            }

            // Create a new row for the result
            Workbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
            Row newRow = workbook.createSheet().createRow(0);

            // Cell 0: First cell of the first row
            if (!rows.isEmpty()) {
                Cell firstCell = rows.get(0).getCell(0);
                newRow.createCell(0).setCellValue(firstCell != null ? firstCell.getNumericCellValue() : 0);
            }

            // Cell 1: Second cell of the first row
            if (!rows.isEmpty() && rows.get(0).getCell(1) != null) {
                Cell secondCellFirst = rows.get(0).getCell(1);
                newRow.createCell(1).setCellValue(secondCellFirst.getStringCellValue());
            }

            // Cell 1: Second cell of the first row
            if (!rows.isEmpty() && rows.size() > 1 && rows.get(1) != null && rows.get(1).getCell(1) != null) {
                Cell secondCellSecond = rows.get(1).getCell(1);
                BigDecimal input = BigDecimal.valueOf(Helper.getRates(secondCellSecond.getStringCellValue()));
                List<String>  groupsId = CogsHelper.findGroupIdsByRate(input, FiscalYear.FY25,recogs);
                newRow.createCell(2).setCellValue(groupsId.toString());
               // newRow.createCell(2).setCellValue(CogsHelper.getCategory(input));
            }

            // RATE COLUMN
            CellStyle currencyStyle = Helper.getCurrencyStyle(workbook);
            if (rows.size() > 1 && rows.get(1).getCell(1) != null) {
                Cell secondCellSecond = rows.get(1).getCell(1);
                BigDecimal input = BigDecimal.valueOf(Helper.getRates(secondCellSecond.getStringCellValue()));
                String description = rows.get(0)!=null && rows.get(0).getCell(1)!=null && CellType.STRING.equals(rows.get(0).getCell(1).getCellType()) && !rows.get(0).getCell(1).getStringCellValue().isEmpty() ? rows.get(0).getCell(1).getStringCellValue():"";
                Cell thirdCell = newRow.createCell(3);

                if(!description.isEmpty()) {
                    BigDecimal exactRate = getExactValueFromSheet(inputWorkbook, sheetNameEs, description,6);
                    thirdCell.setCellValue(Helper.round(exactRate.doubleValue()));
                }
                thirdCell.setCellStyle(currencyStyle);
            }

            // Cells N (N >= 3): Based on the third cell and beyond
            if (rows.size() > 1) {
                Row teamServiceRow = rows.get(1);
                Row vacationRow = rows.size() > 2 ? rows.get(2) : null;
                for (int i = 4; i < teamServiceRow.getLastCellNum() + 4; i++) { // Limit to 10 cells for practicality,
                    // adjust as needed
                    Cell hoursCell = teamServiceRow.getCell(i - 2);
                    Cell vacationsCell = vacationRow != null ? vacationRow.getCell(i - 2) : null;
                    Cell outputCell = newRow.createCell(i);
                    if (hoursCell != null && hoursCell.getCellType() != CellType.BLANK) {
                        switch (hoursCell.getCellType()) {
                            case NUMERIC:
                                outputCell.setCellValue(hoursCell.getNumericCellValue());
                                break;
                            case FORMULA:
                                Workbook w = teamServiceRow.getSheet().getWorkbook();
                                FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator();

                                // Evaluate the formula cell to get its numeric value
                                CellValue cellValue = evaluator.evaluate(hoursCell);
                                BigDecimal numericValue = BigDecimal.valueOf(cellValue.getNumberValue());
                                outputCell.setCellValue(Helper.round(numericValue.doubleValue()));
                                break;
                            default:
                                outputCell.setCellValue(hoursCell.getStringCellValue());
                                break;
                        }
                    }
                    if (outputCell == null || outputCell.getCellType() == CellType.BLANK
                            || (outputCell.getCellType() == CellType.STRING
                            && outputCell.getStringCellValue().isEmpty())) {
                        if (vacationsCell != null) {
                            switch (vacationsCell.getCellType()) {
                                case NUMERIC:
                                    outputCell.setCellValue(vacationsCell.getNumericCellValue());
                                    break;
                                case STRING:
                                    outputCell.setCellValue(!vacationsCell.getStringCellValue().isEmpty()
                                            ? vacationsCell.getStringCellValue()
                                            : "");
                                    break;
                                default:
                                    break;
                            }
                        }

                    }
                }
            }

            if (newRow.getCell(2) != null)
                resultMap.put(key, newRow);
        }

        return resultMap;
    }

    public BigDecimal getTotalServiceTeam(Workbook inputWorkbook, String serviceTeam, String sheetName){
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        FormulaEvaluator evaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();
        if (sheet == null) return BigDecimal.ZERO;

        BigDecimal total = BigDecimal.ZERO;
        boolean inProjectBlock = false;
        String projectBlock = "";
        for (Row row : sheet) {
            Cell projectCell = row.getCell(1); // Column B (index 1)
            Cell cell0 = row.getCell(0);
            projectBlock =cell0!=null && CellType.STRING.equals(cell0.getCellType())  && cell0.getStringCellValue() != null && !cell0.getStringCellValue().isEmpty() && cell0.getStringCellValue().equals("Número Empleado") ? projectCell.getStringCellValue() : projectBlock;
            Cell totalCell = row.getCell(7);  // Column H (index 7)
            String project = projectCell !=null ? projectCell.getStringCellValue()!=null ? projectCell.getStringCellValue().trim():"":"";
            BigDecimal val = totalCell!=null? BigDecimal.valueOf(evaluator.evaluate(totalCell).getNumberValue()) : BigDecimal.ZERO;
            if(project.isEmpty() && val !=BigDecimal.ZERO && projectBlock.contains(serviceTeam)) {
                total =  val ;
                break;
            }
            if (projectCell != null && projectCell.getCellType() == CellType.STRING) {
                String cellValue = projectCell.getStringCellValue().trim();
                if (cellValue.contains(serviceTeam)) {
                    inProjectBlock = true;
                } else if (inProjectBlock && !cellValue.isEmpty()) {
                    inProjectBlock = false; // End of project block
                }
            }
        }
        return total;
    }

    public BigDecimal getExactValueFromSheet(Workbook inputWorkbook, String sheetName, String rowDescription, int column){
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        FormulaEvaluator evaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();
        if (sheet == null) return BigDecimal.ZERO;

        BigDecimal exactValue = BigDecimal.ZERO;
        for (Row row : sheet) {
            Cell cellDescription = row.getCell(1); // Column B (index 1)
            if(cellDescription!=null && CellType.STRING.equals(cellDescription.getCellType())  && cellDescription.getStringCellValue() != null && !cellDescription.getStringCellValue().isEmpty() && cellDescription.getStringCellValue().equals(rowDescription)){
                Cell cellValue = row.getCell(column);
                if (cellValue != null && cellValue.getCellType() == CellType.NUMERIC) {
                    exactValue = BigDecimal.valueOf(cellValue.getNumericCellValue());
                }
            }
        }
        return exactValue;
    }

}