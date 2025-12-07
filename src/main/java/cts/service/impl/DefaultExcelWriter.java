package cts.service.impl;


import cts.service.ExcelWriter;
import cts.service.RateTable;
import cts.util.Utils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class DefaultExcelWriter implements ExcelWriter {

    @Override
    public Workbook createWorkbookWithSheets(String currentMonthName, String nextMonthName, String nextNextMonthName) {
        Workbook workbook = new XSSFWorkbook();
        workbook.createSheet("Service Hours Details " + currentMonthName);
        workbook.createSheet("Service Hours Details " + nextMonthName);
        workbook.createSheet("Service Hours Details " + nextNextMonthName);
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
        // TODO Auto-generated method stub
        Sheet inputSheet = inputWorkbook.getSheet(invoicingSheetNameES);
        Sheet outputSheet = outputWorkbook.getSheet(invoicingSheetName);
        if (inputSheet == null || outputSheet == null) {
            System.err.println("Skipping invoicing details sheet: input or output sheet not found.");
            return;
        }

        int outputRowIndex = 0;

        // Create header style (white font, #003399 background)
        CellStyle headerStyle = Utils.getHeaderStyle(outputWorkbook);
        CellStyle centerStyle = Utils.getCenterStandardStyle(outputWorkbook);
        CellStyle leftStyle = Utils.getLeftStandardStyle(outputWorkbook);
        CellStyle currencyStyle = Utils.getCurrencyStyle(outputWorkbook);
        CellStyle dateStyle = Utils.getDateStyle(outputWorkbook);
        CellStyle vacanceStyle = Utils.getVacanceStyle(outputWorkbook);
        CellStyle freedayStyle = Utils.getFreedayStyle(outputWorkbook);
        CellStyle sickLeaveStyle = Utils.getSickLeaveStyle(outputWorkbook);
        CellStyle legalAbsenceStyle = Utils.getLegalAbsenceStyle(outputWorkbook);
        CellStyle weekendStyle = Utils.getWeekendStyle(outputWorkbook);
        CellStyle footerCurrencyStyle = Utils.getFooterCurrencyStyle(outputWorkbook);

        // Create and set custom header row
        Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
        String[] headers = {"Empl. N°", "Person", "Category", "Rates"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = outputHeaderRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
        int nbrDaysInThisMonths = Utils.numberOfDays(invoicingSheetName);
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

        // Iterate through rows in the input sheet
        Map<Double, List<Row>> maps = getAllData(inputSheet);
        Map<Double, List<Row>> mapsByServiceTeam = filterRowsByServiceTeam(maps, serviceTeam);
        Map<Double, Row> mergedMaps = transformRows(inputWorkbook, facturacionSheetName,mapsByServiceTeam);
        // WRITE DTAA IN EXCEL FILE
        for (Map.Entry<Double, Row> entry : mergedMaps.entrySet()) {
            Row row = entry.getValue();
            Row outputRow = outputSheet.createRow(outputRowIndex++);
            ALL:
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell inputCell = row.getCell(j);
                Cell outputCell = outputRow.createCell(j);
                if (inputCell != null) {
                    switch (inputCell.getCellType()) {
                        case NUMERIC:
                            outputCell.setCellValue(inputCell.getNumericCellValue());
                            outputCell.setCellStyle(centerStyle);
                            if (j == 3) {
                                outputCell.setCellStyle(currencyStyle);
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
        int lastColumn = 1;
        Row row1 = outputSheet.getRow(1);
        lastColumn = row1!=null? row1.getLastCellNum() - 1 : 1;
        String letterCost = Utils.getColumnLetter(lastColumn);
        String letterCost2 = Utils.getColumnLetter(lastColumn + 1);
        String letterTotalHours = Utils.getColumnLetter(lastColumn);

        for (Row row : outputSheet) {
            if (row.getRowNum() != 0) {
                // Create a new cell in the last column (new column)
                Cell newCell = row.createCell(lastColumn + 1);
                Cell rateCell = row.getCell(3);
                Cell descCell = row.getCell(1);
                if(rateCell!=null && CellType.NUMERIC.equals(rateCell.getCellType()) && rateCell.getNumericCellValue() != 0) {
                    // Create formula: D(row_number) * last_column
                    String formula = "D" + (row.getRowNum() + 1) + "*" + letterCost + (row.getRowNum() + 1);
                    newCell.setCellFormula(formula);
                } else {
//                    Double totalValue = getExactRowValueFromSheet();
                    newCell.setCellValue(descCell.getStringCellValue());
                }
                newCell.setCellStyle(currencyStyle);
            }
        }
        // ADD ADJUSTMENT
        int month = Utils.getMonthFromSheetName(invoicingSheetName);
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
            cellD.setCellValue(row.getCell(15).getNumericCellValue());
            cellD.setCellStyle(currencyStyle);

            Cell cellE = outputRow.createCell(4);
            cellE.setCellValue(row.getCell(12).getNumericCellValue());
            cellE.setCellStyle(centerStyle);
            int lastCol = headers.length + nbrDaysInThisMonths;
            for (int i = headers.length + 1; i < lastCol; i++) {
                Cell cellM = outputRow.createCell(i);
                cellM.setCellValue("");
                cellM.setCellStyle(centerStyle);
            }
            Cell cellAdj = outputRow.createCell(lastCol);
            Cell cellCost = outputRow.createCell(lastCol + 1);

        //    if(CellType.NUMERIC.equals(cellE.getCellType()) && cellE.getNumericCellValue()!=0){
                cellAdj.setCellValue(row.getCell(15).getNumericCellValue());
                String formulaCost = "D" + outputRowIndex + "*E" + outputRowIndex;
                cellCost.setCellFormula(formulaCost);
                if(cellCost.getNumericCellValue()==0){
                    cellCost.setCellValue(row.getCell(16).getNumericCellValue());
                }
//            } else {
//                String desc = cellB.getStringCellValue();
//                double agustoValue = getAgustoExactValueFromSheet(inputWorkbook, ajustesSheetName, serviceTeam, desc);
//                cellAdj.setCellValue(9999);
//                cellCost.setCellValue(9999);
//            }

            cellAdj.setCellStyle(centerStyle);
            cellCost.setCellStyle(currencyStyle);


        }
        // ADD TOTAL ROW
        Row lastRow = outputSheet.createRow(outputRowIndex);
        if(lastRow !=null && lastColumn >2) {
            Cell cellTotal = lastRow.createCell(lastColumn - 2) ;

            cellTotal.setCellValue("Total");
            cellTotal.setCellStyle(headerStyle);
            outputSheet.addMergedRegion(new CellRangeAddress(
                    outputRowIndex, // First row (0-based)
                    outputRowIndex, // Last row (0-based)
                    lastColumn - 2, // First column (0-based)
                    lastColumn - 1  // Last column (0-based)
            ));
        }
        Cell cellTotalHours = lastRow.createCell(lastColumn);
        String formula = "SUM(" + letterTotalHours + "2:" + letterTotalHours + (outputRowIndex) + ")";
        cellTotalHours.setCellFormula(formula);
        cellTotalHours.setCellStyle(headerStyle);

        Cell cellTotalCost = lastRow.createCell(lastColumn + 1);
        double total = getTotalServiceTeam(inputWorkbook,serviceTeam,facturacionSheetName);
        cellTotalCost.setCellValue(total);
        cellTotalCost.setCellStyle(footerCurrencyStyle);

        // Auto-size columns after all data is written
        for (int col = 0; col < 40; col++) { // Adjust up to column G (index 6)
            outputSheet.autoSizeColumn(col);
        }

    }

    private Map<Double, List<Row>> getAllData(Sheet inputSheet) {
        List<Double> ids = new ArrayList<>();
        Map<Double, List<Row>> maps = new HashMap<>();
        Double lastId = 0D;
        for (Row row : inputSheet) {
            Cell empIdCell = row.getCell(0);
            Cell empNameCell = row.getCell(1);
            if (empIdCell != null && empNameCell != null) {
                lastId = empIdCell.getNumericCellValue();
                ids.add(lastId);
                List<Row> list = new ArrayList<>();
                list.add(row);
                maps.put(empIdCell.getNumericCellValue(), list);
            } else {
                if (!Utils.isRowEmpty(row) && lastId != 0) {
                    List<Row> list = maps.get(lastId);
                    list.add(row);
                    maps.put(lastId, list);
                }
            }
        }
        return maps;
    }

    private static Map<Double, List<Row>> filterRowsByServiceTeam(Map<Double, List<Row>> inputMap, String serviceTeam) {
        Map<Double, List<Row>> filteredMap = new HashMap<>();

        for (Map.Entry<Double, List<Row>> entry : inputMap.entrySet()) {
            Double key = entry.getKey();
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

    private Map<Double, Row> transformRows(Workbook inputWorkbook,String sheetNameEs ,Map<Double, List<Row>> inputMap) {
        Map<Double, Row> resultMap = new HashMap<>();
        for (Map.Entry<Double, List<Row>> entry : inputMap.entrySet()) {
            Double key = entry.getKey();
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
                Double input = Utils.getRates(secondCellSecond.getStringCellValue());

                newRow.createCell(2).setCellValue(RateTable.getCategory(input));
            }

            // RATE COLUMN
            CellStyle currencyStyle = Utils.getCurrencyStyle(workbook);
            if (rows.size() > 1 && rows.get(1).getCell(1) != null) {
                Cell secondCellSecond = rows.get(1).getCell(1);
                Double input = Utils.getRates(secondCellSecond.getStringCellValue());
                String description = rows.get(0)!=null && rows.get(0).getCell(1)!=null && CellType.STRING.equals(rows.get(0).getCell(1).getCellType()) && !rows.get(0).getCell(1).getStringCellValue().isEmpty() ? rows.get(0).getCell(1).getStringCellValue():"";
                Cell thirdCell = newRow.createCell(3);

                if(!description.isEmpty()) {
                    Double exactRate = getExactValueFromSheet(inputWorkbook, sheetNameEs, description,6);
                    thirdCell.setCellValue(exactRate);
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
                                double numericValue = cellValue.getNumberValue();
                                outputCell.setCellValue(numericValue);
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

    public Double getTotalServiceTeam(Workbook inputWorkbook, String serviceTeam, String sheetName){
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        FormulaEvaluator evaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();
        if (sheet == null) return 0.0;

        Double total = 0.0;
        boolean inProjectBlock = false;
        String projectBlock = "";
        for (Row row : sheet) {
            Cell projectCell = row.getCell(1); // Column B (index 1)
            Cell cell0 = row.getCell(0);
            projectBlock =cell0!=null && CellType.STRING.equals(cell0.getCellType())  && cell0.getStringCellValue() != null && !cell0.getStringCellValue().isEmpty() && cell0.getStringCellValue().equals("Número Empleado") ? projectCell.getStringCellValue() : projectBlock;
            Cell totalCell = row.getCell(7);  // Column H (index 7)
            String project = projectCell !=null ? projectCell.getStringCellValue()!=null ? projectCell.getStringCellValue().trim():"":"";
            double val = totalCell!=null? evaluator.evaluate(totalCell).getNumberValue() : 0;
            if(project.isEmpty() && val !=0 && projectBlock.contains(serviceTeam)) {
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

    public Double getExactValueFromSheet(Workbook inputWorkbook, String sheetName, String rowDescription, int column){
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        FormulaEvaluator evaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();
        if (sheet == null) return 0.0;

        double exactValue = 0.0;
        for (Row row : sheet) {
            Cell cellDescription = row.getCell(1); // Column B (index 1)
            if(cellDescription!=null && CellType.STRING.equals(cellDescription.getCellType())  && cellDescription.getStringCellValue() != null && !cellDescription.getStringCellValue().isEmpty() && cellDescription.getStringCellValue().equals(rowDescription)){
                Cell cellValue = row.getCell(column);
                exactValue = cellValue.getNumericCellValue();
            }
        }
        return exactValue;
    }

}