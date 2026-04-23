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
        if (inputSheet == null) {
            return rows;
        }
        for (Row inputRow : inputSheet) {
            if (inputRow.getRowNum() == 0) {
                continue;
            }
            Cell cellE = inputRow.getCell(4); // Column E
            Cell cellH = inputRow.getCell(7); // Column H
            if (cellE == null || cellH == null) {
                continue;
            }
            if (cellE.getCellType() != CellType.STRING || cellH.getCellType() != CellType.NUMERIC) {
                continue;
            }
            if (!cellE.getStringCellValue().equals(serviceTeam)) {
                continue;
            }
            if (cellH.getDateCellValue().getMonth() == (month - 1)) {
                rows.add(inputRow);
            }
        }
        return rows;
    }

    @Override
    public void copyServiceHoursSheetData(Workbook inputWorkbook, Workbook outputWorkbook, String serviceTeam,
                                          String invoicingSheetNameES, String invoicingSheetName, String ajustesSheetName, String facturacionSheetName) {
        Sheet inputSheet = inputWorkbook.getSheet(invoicingSheetNameES);
        Sheet outputSheet = outputWorkbook.getSheet(invoicingSheetName);
        if (inputSheet == null || outputSheet == null) {
            System.err.println("Skipping invoicing details sheet: input or output sheet not found.");
            return;
        }

        int outputRowIndex = 0;

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

        Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
        String[] headers = {"Empl. N°", "Person", "Category", "Rate"};
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

        Map<BigDecimal, List<Row>> maps = getAllData(inputSheet);
        Map<BigDecimal, List<Row>> mapsByServiceTeam = filterRowsByServiceTeam(maps, serviceTeam);
        Map<BigDecimal, Row> mergedMaps = transformRows(inputWorkbook, facturacionSheetName, mapsByServiceTeam);

        for (Map.Entry<BigDecimal, Row> entry : mergedMaps.entrySet()) {
            Row row = entry.getValue();
            Row outputRow = outputSheet.createRow(outputRowIndex++);
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

        int lastColumn = 1;
        Row row1 = outputSheet.getRow(1);
        lastColumn = row1 != null ? row1.getLastCellNum() - 1 : 1;
        String letterTotalHours = Helper.getColumnLetter(lastColumn);

        for (Row row : outputSheet) {
            if (row.getRowNum() != 0) {
                Cell newCell = row.createCell(lastColumn + 1);
                Cell rateCell = row.getCell(3);
                Cell descCell = row.getCell(1);
                if (rateCell != null && CellType.NUMERIC.equals(rateCell.getCellType()) && rateCell.getNumericCellValue() != 0) {
                    String rowNumber = String.valueOf(row.getRowNum() + 1);
                    String formula = "IF(D" + rowNumber + "=0,1,D" + rowNumber + ")" +
                            "*IF(" + letterTotalHours + rowNumber + "=0,1," + letterTotalHours + rowNumber + ")";
                    newCell.setCellFormula(formula);
                } else {
                    if (descCell != null) {
                        newCell.setCellValue(descCell.getStringCellValue());
                    }
                }
                newCell.setCellStyle(currencyStyle);
            }
        }

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

            int lastCol = headers.length + nbrDaysInThisMonths;
            for (int i = headers.length + 1; i < lastCol; i++) {
                Cell cellM = outputRow.createCell(i);
                cellM.setCellValue("");
                cellM.setCellStyle(centerStyle);
            }

            Cell cellAdj = outputRow.createCell(lastCol);
            Cell cellCost = outputRow.createCell(lastCol + 1);

            BigDecimal workingHours = new BigDecimal(row.getCell(12).getNumericCellValue());
            BigDecimal adjustmentCost = new BigDecimal(row.getCell(16).getNumericCellValue());
            BigDecimal computedHours = workingHours;
            if (workingHours.compareTo(BigDecimal.ZERO) == 0 && adjustmentCost.compareTo(BigDecimal.ZERO) != 0) {
                // Align month-hours behavior with Rate module:
                // - if rate exists, derive hours = cost / rate
                // - if rate is zero, treat as Tools-like row and use hours = cost
                computedHours = hourlyRate.compareTo(BigDecimal.ZERO) == 0
                        ? adjustmentCost
                        : adjustmentCost.divide(hourlyRate, 10, java.math.RoundingMode.HALF_UP);
            }

            cellAdj.setCellValue(Helper.round(computedHours.doubleValue()));
            cellAdj.setCellStyle(centerStyle);

            if (workingHours.compareTo(BigDecimal.ZERO) == 0) {
                cellCost.setCellValue(Helper.round(adjustmentCost.doubleValue()));
            } else {
                cellCost.setCellValue(Helper.round(workingHours.multiply(hourlyRate).doubleValue()));
            }
            cellCost.setCellStyle(currencyStyle);
        }

        // ADD TOTAL ROW
        Row lastRow = outputSheet.createRow(outputRowIndex);
        if (lastRow != null && lastColumn > 2) {
            Cell cellTotal = lastRow.createCell(lastColumn - 2);
            cellTotal.setCellValue("Total");
            cellTotal.setCellStyle(headerStyle);
            outputSheet.addMergedRegion(new CellRangeAddress(
                    outputRowIndex,
                    outputRowIndex,
                    lastColumn - 2,
                    lastColumn - 1
            ));
        }
        Cell cellTotalHours = lastRow.createCell(lastColumn);
        String formulaHours = "SUM(" + letterTotalHours + "2:" + letterTotalHours + (outputRowIndex) + ")";
        cellTotalHours.setCellFormula(formulaHours);
        cellTotalHours.setCellStyle(headerStyle);

        Cell cellTotalRate = lastRow.createCell(3);
        String formulaRateSum = "SUM(D2:D" + (outputRowIndex) + ")";
        cellTotalRate.setCellFormula(formulaRateSum);
        cellTotalRate.setCellStyle(footerCurrencyStyle);

        String letterTotalCost = Helper.getColumnLetter(lastColumn + 1);
        Cell cellTotalCost = lastRow.createCell(lastColumn + 1);
        String formulaCost = "SUM(" + letterTotalCost + "2:" + letterTotalCost + (outputRowIndex) + ")";
        cellTotalCost.setCellFormula(formulaCost);
        cellTotalCost.setCellStyle(footerCurrencyStyle);

        for (int col = 0; col < 40; col++) {
            outputSheet.autoSizeColumn(col);
        }
    }

    @Override
    public int copyServiceHoursToConsolidatedSheet(
            Workbook inputWorkbook,
            Sheet consolidatedSheet,
            int startRow,
            String serviceTeam,
            String invoicingSheetNameES,
            String invoicingSheetNameEN,
            String ajustesSheetName,
            String facturacionSheetName
    ) {
        Sheet inputSheet = inputWorkbook.getSheet(invoicingSheetNameES);
        if (inputSheet == null) {
            System.err.println("Skipping consolidated sheet: input sheet not found: " + invoicingSheetNameES);
            return startRow;
        }

        Workbook wb = consolidatedSheet.getWorkbook();
        CellStyle headerStyle = Helper.getHeaderStyle(wb);
        CellStyle leftStyle = Helper.getLeftStandardStyle(wb);
        CellStyle centerStyle = Helper.getCenterStandardStyle(wb);
        CellStyle currencyStyle = Helper.getCurrencyStyle(wb);
        CellStyle dateStyle = Helper.getDateStyle(wb);
        CellStyle vacanceStyle = Helper.getVacanceStyle(wb);
        CellStyle freedayStyle = Helper.getFreedayStyle(wb);
        CellStyle sickLeaveStyle = Helper.getSickLeaveStyle(wb);
        CellStyle legalAbsenceStyle = Helper.getLegalAbsenceStyle(wb);
        CellStyle weekendStyle = Helper.getWeekendStyle(wb);
        CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(wb);

        int rowIdx = startRow;
        int nbrDaysInThisMonth = Helper.numberOfDays(invoicingSheetNameEN);
        int hoursCol = 4 + nbrDaysInThisMonth;
        int costCol = hoursCol + 1;

        Row headerRow = consolidatedSheet.createRow(rowIdx++);
        String[] headers = {"Empl. N", "Person", "Category", "Rate"};
        for (int c = 0; c < headers.length; c++) {
            Cell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(headerStyle);
        }
        for (int day = 1; day <= nbrDaysInThisMonth; day++) {
            Cell dayCell = headerRow.createCell(3 + day);
            dayCell.setCellValue(day);
            dayCell.setCellStyle(dateStyle);
        }
        Cell hoursHeaderCell = headerRow.createCell(hoursCol);
        hoursHeaderCell.setCellValue("Working Hours");
        hoursHeaderCell.setCellStyle(headerStyle);

        Cell costHeaderCell = headerRow.createCell(costCol);
        costHeaderCell.setCellValue("Cost (Euro)");
        costHeaderCell.setCellStyle(headerStyle);

        Map<BigDecimal, List<Row>> allData = getAllData(inputSheet);
        Map<BigDecimal, List<Row>> teamData = filterRowsByServiceTeam(allData, serviceTeam);
        Map<BigDecimal, Row> mergedData = transformRows(inputWorkbook, facturacionSheetName, teamData);
        int hoursColInTransformed = 4 + nbrDaysInThisMonth;

        List<Integer> dataRowIndices = new ArrayList<>();

        for (Map.Entry<BigDecimal, Row> entry : mergedData.entrySet()) {
            Row srcRow = entry.getValue();
            Row outRow = consolidatedSheet.createRow(rowIdx);
            dataRowIndices.add(rowIdx);
            rowIdx++;

            Cell empCell = outRow.createCell(0);
            empCell.setCellValue(srcRow.getCell(0) != null ? srcRow.getCell(0).getNumericCellValue() : 0);
            empCell.setCellStyle(centerStyle);

            Cell nameCell = outRow.createCell(1);
            nameCell.setCellValue(srcRow.getCell(1) != null ? srcRow.getCell(1).getStringCellValue() : "");
            nameCell.setCellStyle(leftStyle);

            Cell catCell = outRow.createCell(2);
            catCell.setCellValue(srcRow.getCell(2) != null ? srcRow.getCell(2).getStringCellValue() : "");
            catCell.setCellStyle(leftStyle);

            double rate = srcRow.getCell(3) != null ? srcRow.getCell(3).getNumericCellValue() : 0;
            Cell rateCell = outRow.createCell(3);
            if (rate != 0) {
                rateCell.setCellValue(Helper.round(rate));
            } else {
                rateCell.setCellValue("");
            }
            rateCell.setCellStyle(currencyStyle);

            FormulaEvaluator evaluator = srcDayCellEvaluator(srcRow);
            for (int dayOffset = 0; dayOffset < nbrDaysInThisMonth; dayOffset++) {
                int srcCol = 4 + dayOffset;
                int outCol = 4 + dayOffset;
                Cell srcDayCell = srcRow.getCell(srcCol);
                Cell outDayCell = outRow.createCell(outCol);
                if (srcDayCell == null) {
                    outDayCell.setCellStyle(centerStyle);
                    continue;
                }

                switch (srcDayCell.getCellType()) {
                    case NUMERIC:
                        outDayCell.setCellValue(srcDayCell.getNumericCellValue());
                        outDayCell.setCellStyle(centerStyle);
                        break;
                    case STRING:
                        String value = srcDayCell.getStringCellValue();
                        outDayCell.setCellValue(value);
                        outDayCell.setCellStyle("V".equals(value) ? vacanceStyle
                                : "A".equals(value) ? legalAbsenceStyle
                                : "F".equals(value) ? freedayStyle
                                : "S".equals(value) ? sickLeaveStyle : weekendStyle);
                        break;
                    case FORMULA:
                        CellValue evaluated = evaluator.evaluate(srcDayCell);
                        if (evaluated != null && evaluated.getCellType() == CellType.NUMERIC) {
                            outDayCell.setCellValue(evaluated.getNumberValue());
                        } else if (evaluated != null && evaluated.getCellType() == CellType.STRING) {
                            outDayCell.setCellValue(evaluated.getStringValue());
                        } else {
                            outDayCell.setCellValue("");
                        }
                        outDayCell.setCellStyle(centerStyle);
                        break;
                    default:
                        outDayCell.setCellStyle(centerStyle);
                        break;
                }
            }

            double hours = 0;
            Cell srcHours = srcRow.getCell(hoursColInTransformed);
            if (srcHours != null && srcHours.getCellType() == CellType.NUMERIC) {
                hours = srcHours.getNumericCellValue();
            }
            Cell hoursCell = outRow.createCell(hoursCol);
            hoursCell.setCellValue(hours);
            hoursCell.setCellStyle(centerStyle);

            Cell costCell = outRow.createCell(costCol);
            if (rate != 0) {
                costCell.setCellValue(Helper.round(rate * hours));
            } else {
                costCell.setCellValue("");
            }
            costCell.setCellStyle(currencyStyle);
        }

        int month = Helper.getMonthFromSheetName(invoicingSheetNameEN);
        List<Row> adjustments = getAdjustmentSheetData(inputWorkbook, ajustesSheetName, serviceTeam, month);

        for (Row adjRow : adjustments) {
            Row outRow = consolidatedSheet.createRow(rowIdx);
            dataRowIndices.add(rowIdx);
            rowIdx++;

            outRow.createCell(0).setCellStyle(centerStyle);

            Cell descCell = outRow.createCell(1);
            descCell.setCellValue(adjRow.getCell(6) != null ? adjRow.getCell(6).getStringCellValue() : "");
            descCell.setCellStyle(leftStyle);

            outRow.createCell(2).setCellStyle(centerStyle);

            BigDecimal workingHours = new BigDecimal(adjRow.getCell(12).getNumericCellValue());
            BigDecimal hourlyRate = new BigDecimal(adjRow.getCell(15).getNumericCellValue());
            BigDecimal adjustmentCost = new BigDecimal(adjRow.getCell(16).getNumericCellValue());
            BigDecimal computedHours = workingHours;
            if (workingHours.compareTo(BigDecimal.ZERO) == 0 && adjustmentCost.compareTo(BigDecimal.ZERO) != 0) {
                computedHours = hourlyRate.compareTo(BigDecimal.ZERO) == 0
                        ? adjustmentCost
                        : adjustmentCost.divide(hourlyRate, 10, java.math.RoundingMode.HALF_UP);
            }

            Cell adjRateCell = outRow.createCell(3);
            if (hourlyRate.compareTo(BigDecimal.ZERO) != 0) {
                adjRateCell.setCellValue(Helper.round(hourlyRate.doubleValue()));
            } else {
                adjRateCell.setCellValue("");
            }
            adjRateCell.setCellStyle(currencyStyle);

            for (int dayCol = 4; dayCol < hoursCol; dayCol++) {
                Cell dayBlankCell = outRow.createCell(dayCol);
                dayBlankCell.setCellValue("");
                dayBlankCell.setCellStyle(centerStyle);
            }

            Cell adjHoursCell = outRow.createCell(hoursCol);
            adjHoursCell.setCellValue(Helper.round(computedHours.doubleValue()));
            adjHoursCell.setCellStyle(centerStyle);

            Cell adjCostCell = outRow.createCell(costCol);
            if (workingHours.compareTo(BigDecimal.ZERO) == 0) {
                adjCostCell.setCellValue(Helper.round(adjustmentCost.doubleValue()));
            } else {
                adjCostCell.setCellValue(Helper.round(workingHours.multiply(hourlyRate).doubleValue()));
            }
            adjCostCell.setCellStyle(currencyStyle);
        }

        Row totalRow = consolidatedSheet.createRow(rowIdx);
        rowIdx++;

        Cell totalLabelCell = totalRow.createCell(0);
        totalLabelCell.setCellValue("Total " + serviceTeam);
        totalLabelCell.setCellStyle(headerStyle);
        consolidatedSheet.addMergedRegion(new CellRangeAddress(totalRow.getRowNum(), totalRow.getRowNum(), 0, 2));

        String hoursColumnLetter = Helper.getColumnLetter(hoursCol);
        String costColumnLetter = Helper.getColumnLetter(costCol);

        Cell totalRateCell = totalRow.createCell(3);
        String totalRowNumber = String.valueOf(totalRow.getRowNum() + 1);
        totalRateCell.setCellFormula("IF(" + hoursColumnLetter + totalRowNumber + "=0,\"\","
                + costColumnLetter + totalRowNumber + "/" + hoursColumnLetter + totalRowNumber + ")");
        totalRateCell.setCellStyle(footerCurrencyStyle);

        StringBuilder hoursFormula = new StringBuilder("SUM(");
        for (int i = 0; i < dataRowIndices.size(); i++) {
            hoursFormula.append(hoursColumnLetter).append(dataRowIndices.get(i) + 1);
            if (i < dataRowIndices.size() - 1) {
                hoursFormula.append(",");
            }
        }
        hoursFormula.append(")");
        Cell totalHoursCell = totalRow.createCell(hoursCol);
        totalHoursCell.setCellFormula(hoursFormula.toString());
        totalHoursCell.setCellStyle(headerStyle);

        StringBuilder costFormula = new StringBuilder("SUM(");
        for (int i = 0; i < dataRowIndices.size(); i++) {
            costFormula.append(costColumnLetter).append(dataRowIndices.get(i) + 1);
            if (i < dataRowIndices.size() - 1) {
                costFormula.append(",");
            }
        }
        costFormula.append(")");
        Cell totalCostCell = totalRow.createCell(costCol);
        totalCostCell.setCellFormula(costFormula.toString());
        totalCostCell.setCellStyle(footerCurrencyStyle);

        consolidatedSheet.createRow(rowIdx++);
        consolidatedSheet.createRow(rowIdx++);

        return rowIdx;
    }

    private FormulaEvaluator srcDayCellEvaluator(Row srcRow) {
        return srcRow.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
    }

    private Map<BigDecimal, List<Row>> getAllData(Sheet inputSheet) {
        Map<BigDecimal, List<Row>> maps = new HashMap<>();
        BigDecimal lastId = null;

        for (Row row : inputSheet) {
            Cell empIdCell   = row.getCell(0);
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
            BigDecimal key  = entry.getKey();
            List<Row>  rows = entry.getValue();

            List<Row> filteredRows = rows.stream().filter(row -> {
                Cell secondCell = row.getCell(1);
                boolean isFirstRow = rows.indexOf(row) == 0;
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

    private Map<BigDecimal, Row> transformRows(Workbook inputWorkbook, String sheetNameEs, Map<BigDecimal, List<Row>> inputMap) {
        Map<BigDecimal, Row> resultMap = new HashMap<>();

        for (Map.Entry<BigDecimal, List<Row>> entry : inputMap.entrySet()) {
            BigDecimal key  = entry.getKey();
            List<Row>  rows = entry.getValue();

            if (rows == null || rows.isEmpty()) {
                continue;
            }

            Workbook workbook = new XSSFWorkbook();
            Row newRow = workbook.createSheet().createRow(0);

            if (!rows.isEmpty()) {
                Cell firstCell = rows.get(0).getCell(0);
                newRow.createCell(0).setCellValue(firstCell != null ? firstCell.getNumericCellValue() : 0);
            }

            if (!rows.isEmpty() && rows.get(0).getCell(1) != null) {
                Cell secondCellFirst = rows.get(0).getCell(1);
                newRow.createCell(1).setCellValue(secondCellFirst.getStringCellValue());
            }

            if (!rows.isEmpty() && rows.size() > 1 && rows.get(1) != null && rows.get(1).getCell(1) != null) {
                Cell secondCellSecond = rows.get(1).getCell(1);
                BigDecimal input = BigDecimal.valueOf(Helper.getRates(secondCellSecond.getStringCellValue()));
                List<String> groupsId = CogsHelper.findGroupIdsByRate(input, FiscalYear.FY25, recogs);
                newRow.createCell(2).setCellValue(groupsId.toString());
            }

            CellStyle currencyStyle = Helper.getCurrencyStyle(workbook);
            if (rows.size() > 1 && rows.get(1).getCell(1) != null) {
                Cell secondCellSecond = rows.get(1).getCell(1);
                BigDecimal input = BigDecimal.valueOf(Helper.getRates(secondCellSecond.getStringCellValue()));
                String description = rows.get(0) != null && rows.get(0).getCell(1) != null
                        && CellType.STRING.equals(rows.get(0).getCell(1).getCellType())
                        && !rows.get(0).getCell(1).getStringCellValue().isEmpty()
                        ? rows.get(0).getCell(1).getStringCellValue() : "";
                Cell thirdCell = newRow.createCell(3);

                if (!description.isEmpty()) {
                    BigDecimal exactRate = getExactValueFromSheet(inputWorkbook, sheetNameEs, description, 6);
                    thirdCell.setCellValue(Helper.round(exactRate.doubleValue()));
                }
                thirdCell.setCellStyle(currencyStyle);
            }

            if (rows.size() > 1) {
                Row teamServiceRow = rows.get(1);
                Row vacationRow    = rows.size() > 2 ? rows.get(2) : null;
                for (int i = 4; i < teamServiceRow.getLastCellNum() + 4; i++) {
                    Cell hoursCell    = teamServiceRow.getCell(i - 2);
                    Cell vacationsCell = vacationRow != null ? vacationRow.getCell(i - 2) : null;
                    Cell outputCell   = newRow.createCell(i);
                    if (hoursCell != null && hoursCell.getCellType() != CellType.BLANK) {
                        switch (hoursCell.getCellType()) {
                            case NUMERIC:
                                outputCell.setCellValue(hoursCell.getNumericCellValue());
                                break;
                            case FORMULA:
                                Workbook w = teamServiceRow.getSheet().getWorkbook();
                                FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator();
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
                                            ? vacationsCell.getStringCellValue() : "");
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

    public BigDecimal getTotalServiceTeam(Workbook inputWorkbook, String serviceTeam, String sheetName) {
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        FormulaEvaluator evaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();
        if (sheet == null) return BigDecimal.ZERO;

        BigDecimal total = BigDecimal.ZERO;
        boolean inProjectBlock = false;
        String projectBlock = "";
        for (Row row : sheet) {
            Cell projectCell = row.getCell(1);
            Cell cell0 = row.getCell(0);
            projectBlock = cell0 != null && CellType.STRING.equals(cell0.getCellType())
                    && cell0.getStringCellValue() != null && !cell0.getStringCellValue().isEmpty()
                    && cell0.getStringCellValue().equals("Número Empleado")
                    ? projectCell.getStringCellValue() : projectBlock;
            Cell totalCell = row.getCell(7);
            String project = projectCell != null ? (projectCell.getStringCellValue() != null ? projectCell.getStringCellValue().trim() : "") : "";
            BigDecimal val = totalCell != null ? BigDecimal.valueOf(evaluator.evaluate(totalCell).getNumberValue()) : BigDecimal.ZERO;
            if (project.isEmpty() && val != BigDecimal.ZERO && projectBlock.contains(serviceTeam)) {
                total = val;
                break;
            }
            if (projectCell != null && projectCell.getCellType() == CellType.STRING) {
                String cellValue = projectCell.getStringCellValue().trim();
                if (cellValue.contains(serviceTeam)) {
                    inProjectBlock = true;
                } else if (inProjectBlock && !cellValue.isEmpty()) {
                    inProjectBlock = false;
                }
            }
        }
        return total;
    }

    public BigDecimal getExactValueFromSheet(Workbook inputWorkbook, String sheetName, String rowDescription, int column) {
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        FormulaEvaluator evaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();
        if (sheet == null) return BigDecimal.ZERO;

        BigDecimal exactValue = BigDecimal.ZERO;
        for (Row row : sheet) {
            Cell cellDescription = row.getCell(1);
            if (cellDescription != null && CellType.STRING.equals(cellDescription.getCellType())
                    && cellDescription.getStringCellValue() != null
                    && !cellDescription.getStringCellValue().isEmpty()
                    && cellDescription.getStringCellValue().equals(rowDescription)) {
                Cell cellValue = row.getCell(column);
                if (cellValue != null && cellValue.getCellType() == CellType.NUMERIC) {
                    exactValue = BigDecimal.valueOf(cellValue.getNumericCellValue());
                }
            }
        }
        return exactValue;
    }
}
