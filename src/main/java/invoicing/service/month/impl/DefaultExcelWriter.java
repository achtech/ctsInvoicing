package invoicing.service.month.impl;

import invoicing.Helper.CogsHelper;
import invoicing.Helper.Helper;
import invoicing.entities.CogsRecord;
import invoicing.enums.FiscalYear;
import invoicing.service.month.ExcelWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
            Cell cellE = inputRow.getCell(4);
            Cell cellH = inputRow.getCell(7);
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
        int nbrDaysInThisMonth = Helper.numberOfDays(invoicingSheetName);
        int transformedHoursCol = 4 + nbrDaysInThisMonth;
        int transformedCostCol = transformedHoursCol + 1;
        int hoursCol = 4;
        int costCol = 5;

        CellStyle headerStyle = Helper.getHeaderStyle(outputWorkbook);
        CellStyle centerStyle = Helper.getCenterStandardStyle(outputWorkbook);
        CellStyle leftStyle = Helper.getLeftStandardStyle(outputWorkbook);
        CellStyle currencyStyle = Helper.getCurrencyStyle(outputWorkbook);
        CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(outputWorkbook);

        Row headerRow = outputSheet.createRow(outputRowIndex++);
        String[] headers = {"Empl. NÂ°", "Person", "Category", "Rate", "Working Hours", "Cost (Euro)"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        Map<BigDecimal, List<Row>> maps = getAllData(inputSheet);
        Map<BigDecimal, List<Row>> mapsByServiceTeam = filterRowsByServiceTeam(maps, serviceTeam);
        Map<BigDecimal, Row> mergedMaps = transformRows(inputWorkbook, facturacionSheetName, mapsByServiceTeam);

        for (Map.Entry<BigDecimal, Row> entry : mergedMaps.entrySet()) {
            Row outputRow = outputSheet.createRow(outputRowIndex++);
            writeCompactDataRow(outputRow, entry.getValue(), leftStyle, centerStyle, currencyStyle, transformedHoursCol, transformedCostCol, nbrDaysInThisMonth);
        }

        int month = Helper.getMonthFromSheetName(invoicingSheetName);
        List<Row> ajustesRows = getAdjustmentSheetData(inputWorkbook, ajustesSheetName, serviceTeam, month);
        for (Row row : ajustesRows) {
            Row outputRow = outputSheet.createRow(outputRowIndex++);

            outputRow.createCell(0).setCellStyle(centerStyle);

            Cell personCell = outputRow.createCell(1);
            personCell.setCellValue(row.getCell(6).getStringCellValue());
            personCell.setCellStyle(leftStyle);

            outputRow.createCell(2).setCellStyle(centerStyle);

            BigDecimal hourlyRate = new BigDecimal(row.getCell(15).getNumericCellValue());
            Cell rateCell = outputRow.createCell(3);
            rateCell.setCellValue(Helper.round(hourlyRate.doubleValue()));
            rateCell.setCellStyle(currencyStyle);

            BigDecimal workingHours = new BigDecimal(row.getCell(12).getNumericCellValue());
            BigDecimal adjustmentCost = new BigDecimal(row.getCell(16).getNumericCellValue());
            BigDecimal computedHours = workingHours;
            if (workingHours.compareTo(BigDecimal.ZERO) == 0 && adjustmentCost.compareTo(BigDecimal.ZERO) != 0) {
                computedHours = hourlyRate.compareTo(BigDecimal.ZERO) == 0
                        ? adjustmentCost
                        : adjustmentCost.divide(hourlyRate, 10, java.math.RoundingMode.HALF_UP);
            }

            Cell hoursCell = outputRow.createCell(hoursCol);
            hoursCell.setCellValue(Helper.round(computedHours.doubleValue()));
            hoursCell.setCellStyle(centerStyle);

            Cell costCell = outputRow.createCell(costCol);
            if (workingHours.compareTo(BigDecimal.ZERO) == 0) {
                costCell.setCellValue(Helper.round(adjustmentCost.doubleValue()));
            } else {
                costCell.setCellValue(Helper.round(workingHours.multiply(hourlyRate).doubleValue()));
            }
            costCell.setCellStyle(currencyStyle);
        }

        Row totalRow = outputSheet.createRow(outputRowIndex);
        Cell totalLabelCell = totalRow.createCell(2);
        totalLabelCell.setCellValue("Total");
        totalLabelCell.setCellStyle(headerStyle);
        outputSheet.addMergedRegion(new CellRangeAddress(outputRowIndex, outputRowIndex, 2, 3));

        String hoursColumnLetter = Helper.getColumnLetter(hoursCol);
        Cell totalHoursCell = totalRow.createCell(hoursCol);
        totalHoursCell.setCellFormula("SUM(" + hoursColumnLetter + "2:" + hoursColumnLetter + outputRowIndex + ")");
        totalHoursCell.setCellStyle(headerStyle);

        String costColumnLetter = Helper.getColumnLetter(costCol);
        Cell totalCostCell = totalRow.createCell(costCol);
        totalCostCell.setCellFormula("SUM(" + costColumnLetter + "2:" + costColumnLetter + outputRowIndex + ")");
        totalCostCell.setCellStyle(footerCurrencyStyle);

        for (int col = 0; col <= costCol; col++) {
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

        int nbrDaysInThisMonth = Helper.numberOfDays(invoicingSheetNameEN);
        int transformedHoursCol = 4 + nbrDaysInThisMonth;
        int transformedCostCol = transformedHoursCol + 1;
        int firstDayCol = 4;
        int hoursCol = firstDayCol + nbrDaysInThisMonth;
        int costCol = hoursCol + 1;
        int rowIdx = startRow;

        Workbook wb = consolidatedSheet.getWorkbook();
        CellStyle headerStyle = Helper.getHeaderStyle(wb);
        CellStyle projectBandStyle = Helper.getProjectBandStyle(wb);
        CellStyle leftStyle = Helper.getLeftStandardStyle(wb);
        CellStyle centerStyle = Helper.getCenterStandardStyle(wb);
        CellStyle currencyStyle = Helper.getCurrencyStyle(wb);
        CellStyle categoryStyle = Helper.getCategoryStyle(wb);
        CellStyle dayValueStyle = Helper.getDayValueStyle(wb);
        CellStyle emptyDayStyle = Helper.getEmptyDayStyle(wb);
        CellStyle dateStyle = Helper.getDateStyle(wb);
        CellStyle vacanceStyle = Helper.getVacanceStyle(wb);
        CellStyle freedayStyle = Helper.getFreedayStyle(wb);
        CellStyle sickLeaveStyle = Helper.getSickLeaveStyle(wb);
        CellStyle legalAbsenceStyle = Helper.getLegalAbsenceStyle(wb);
        CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(wb);

        configureConsolidatedSheetLayout(consolidatedSheet, nbrDaysInThisMonth);

        Row projectRow = consolidatedSheet.createRow(rowIdx++);
        projectRow.setHeightInPoints(20f);
        Cell projectNameCell = projectRow.createCell(0);
        projectNameCell.setCellValue(serviceTeam);
        projectNameCell.setCellStyle(projectBandStyle);
        for (int col = 1; col <= 2; col++) {
            Cell fillerCell = projectRow.createCell(col);
            fillerCell.setCellStyle(projectBandStyle);
        }
        consolidatedSheet.addMergedRegion(new CellRangeAddress(projectRow.getRowNum(), projectRow.getRowNum(), 0, 2));

        Row headerRow = consolidatedSheet.createRow(rowIdx++);
        headerRow.setHeightInPoints(20f);
        String[] headers = {"Empl. N" + '\u00B0', "Person", "Category", "Rates"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
        for (int day = 1; day <= nbrDaysInThisMonth; day++) {
            Cell dayCell = headerRow.createCell(firstDayCol + day - 1);
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
        List<Integer> dataRowIndices = new ArrayList<>();

        for (Map.Entry<BigDecimal, Row> entry : mergedData.entrySet()) {
            Row srcRow = entry.getValue();
            Row outRow = consolidatedSheet.createRow(rowIdx);
            outRow.setHeightInPoints(20f);
            dataRowIndices.add(rowIdx);
            rowIdx++;
            writeConsolidatedDataRow(
                    outRow,
                    srcRow,
                    leftStyle,
                    centerStyle,
                    currencyStyle,
                    categoryStyle,
                    dayValueStyle,
                    emptyDayStyle,
                    vacanceStyle,
                    freedayStyle,
                    sickLeaveStyle,
                    legalAbsenceStyle,
                    transformedHoursCol,
                    transformedCostCol,
                    nbrDaysInThisMonth,
                    firstDayCol,
                    hoursCol,
                    costCol
            );
        }

        int month = Helper.getMonthFromSheetName(invoicingSheetNameEN);
        List<Row> adjustments = getAdjustmentSheetData(inputWorkbook, ajustesSheetName, serviceTeam, month);
        for (Row adjRow : adjustments) {
            Row outRow = consolidatedSheet.createRow(rowIdx);
            outRow.setHeightInPoints(20f);
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

            for (int dayCol = firstDayCol; dayCol < hoursCol; dayCol++) {
                outRow.createCell(dayCol).setCellStyle(centerStyle);
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

        String hoursColumnLetter = Helper.getColumnLetter(hoursCol);
        String costColumnLetter = Helper.getColumnLetter(costCol);

        StringBuilder hoursFormula = new StringBuilder("SUM(");
        for (int i = 0; i < dataRowIndices.size(); i++) {
            hoursFormula.append(hoursColumnLetter).append(dataRowIndices.get(i) + 1);
            if (i < dataRowIndices.size() - 1) {
                hoursFormula.append(",");
            }
        }
        hoursFormula.append(")");

        StringBuilder costFormula = new StringBuilder("SUM(");
        for (int i = 0; i < dataRowIndices.size(); i++) {
            costFormula.append(costColumnLetter).append(dataRowIndices.get(i) + 1);
            if (i < dataRowIndices.size() - 1) {
                costFormula.append(",");
            }
        }
        costFormula.append(")");

        Row bottomTotalRow = consolidatedSheet.createRow(rowIdx++);
        bottomTotalRow.setHeightInPoints(20f);
        int totalLabelStartCol = Math.max(2, hoursCol - 2);
        int totalLabelEndCol = Math.max(totalLabelStartCol, hoursCol - 1);
        Cell bottomLabelCell = bottomTotalRow.createCell(totalLabelStartCol);
        bottomLabelCell.setCellValue("Total");
        bottomLabelCell.setCellStyle(headerStyle);
        consolidatedSheet.addMergedRegion(new CellRangeAddress(
                bottomTotalRow.getRowNum(),
                bottomTotalRow.getRowNum(),
                totalLabelStartCol,
                totalLabelEndCol
        ));

        Cell totalHoursCell = bottomTotalRow.createCell(hoursCol);
        totalHoursCell.setCellFormula(hoursFormula.toString());
        totalHoursCell.setCellStyle(headerStyle);

        Cell totalCostCell = bottomTotalRow.createCell(costCol);
        totalCostCell.setCellFormula(costFormula.toString());
        totalCostCell.setCellStyle(footerCurrencyStyle);

        consolidatedSheet.createRow(rowIdx++);

        return rowIdx;
    }

    private void configureConsolidatedSheetLayout(Sheet sheet, int nbrDaysInThisMonth) {
        sheet.setDisplayGridlines(false);
        sheet.setColumnWidth(0, 11 * 256);
        sheet.setColumnWidth(1, 40 * 256);
        sheet.setColumnWidth(2, 20 * 256);
        sheet.setColumnWidth(3, 9 * 256);
        for (int i = 0; i < nbrDaysInThisMonth; i++) {
            sheet.setColumnWidth(4 + i, 3 * 256);
        }
        sheet.setColumnWidth(4 + nbrDaysInThisMonth, 13 * 256);
        sheet.setColumnWidth(5 + nbrDaysInThisMonth, 13 * 256);
    }

    private void writeConsolidatedDataRow(
            Row outputRow,
            Row sourceRow,
            CellStyle leftStyle,
            CellStyle centerStyle,
            CellStyle currencyStyle,
            CellStyle categoryStyle,
            CellStyle dayValueStyle,
            CellStyle emptyDayStyle,
            CellStyle vacanceStyle,
            CellStyle freedayStyle,
            CellStyle sickLeaveStyle,
            CellStyle legalAbsenceStyle,
            int transformedHoursCol,
            int transformedCostCol,
            int nbrDaysInThisMonth,
            int firstDayCol,
            int hoursCol,
            int costCol
    ) {
        Cell emplCell = outputRow.createCell(0);
        Cell srcEmp = sourceRow.getCell(0);
        if (srcEmp != null && srcEmp.getCellType() == CellType.NUMERIC) {
            emplCell.setCellValue(srcEmp.getNumericCellValue());
        }
        emplCell.setCellStyle(centerStyle);

        Cell personCell = outputRow.createCell(1);
        personCell.setCellValue(sourceRow.getCell(1) != null ? sourceRow.getCell(1).toString() : "");
        personCell.setCellStyle(leftStyle);

        Cell categoryCell = outputRow.createCell(2);
        categoryCell.setCellValue(sourceRow.getCell(2) != null ? sourceRow.getCell(2).toString() : "");
        categoryCell.setCellStyle(categoryStyle);

        double rate = getNumericCellValue(sourceRow.getCell(3));
        Cell rateCell = outputRow.createCell(3);
        if (rate != 0) {
            rateCell.setCellValue(Helper.round(rate));
        } else {
            rateCell.setCellValue("");
        }
        rateCell.setCellStyle(currencyStyle);

        boolean adjustmentLike = isAdjustmentLike(sourceRow);
        for (int dayOffset = 0; dayOffset < nbrDaysInThisMonth; dayOffset++) {
            Cell sourceDayCell = sourceRow.getCell(4 + dayOffset);
            Cell targetDayCell = outputRow.createCell(firstDayCol + dayOffset);
            styleDayCell(targetDayCell, sourceDayCell, adjustmentLike, centerStyle, dayValueStyle, emptyDayStyle, vacanceStyle, freedayStyle, sickLeaveStyle, legalAbsenceStyle);
        }

        double hours = getWorkingHoursFromTransformedRow(sourceRow, transformedHoursCol, nbrDaysInThisMonth);
        Cell hoursCell = outputRow.createCell(hoursCol);
        hoursCell.setCellValue(hours);
        hoursCell.setCellStyle(centerStyle);

        Cell costCell = outputRow.createCell(costCol);
        costCell.setCellValue(getCostFromTransformedRow(sourceRow, transformedCostCol, rate, hours));
        costCell.setCellStyle(currencyStyle);
    }

    private void styleDayCell(
            Cell targetDayCell,
            Cell sourceDayCell,
            boolean adjustmentLike,
            CellStyle centerStyle,
            CellStyle dayValueStyle,
            CellStyle emptyDayStyle,
            CellStyle vacanceStyle,
            CellStyle freedayStyle,
            CellStyle sickLeaveStyle,
            CellStyle legalAbsenceStyle
    ) {
        if (sourceDayCell == null || sourceDayCell.getCellType() == CellType.BLANK) {
            targetDayCell.setCellStyle(adjustmentLike ? centerStyle : emptyDayStyle);
            return;
        }

        if (sourceDayCell.getCellType() == CellType.NUMERIC) {
            targetDayCell.setCellValue(sourceDayCell.getNumericCellValue());
            targetDayCell.setCellStyle(dayValueStyle);
            return;
        }

        if (sourceDayCell.getCellType() == CellType.FORMULA) {
            FormulaEvaluator evaluator = sourceDayCell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellValue evaluated = evaluator.evaluate(sourceDayCell);
            if (evaluated != null && evaluated.getCellType() == CellType.NUMERIC) {
                targetDayCell.setCellValue(Helper.round(evaluated.getNumberValue()));
                targetDayCell.setCellStyle(dayValueStyle);
                return;
            }
            if (evaluated != null && evaluated.getCellType() == CellType.STRING) {
                applyDayStringStyle(targetDayCell, evaluated.getStringValue(), centerStyle, dayValueStyle, emptyDayStyle, vacanceStyle, freedayStyle, sickLeaveStyle, legalAbsenceStyle, adjustmentLike);
                return;
            }
        }

        applyDayStringStyle(targetDayCell, sourceDayCell.toString(), centerStyle, dayValueStyle, emptyDayStyle, vacanceStyle, freedayStyle, sickLeaveStyle, legalAbsenceStyle, adjustmentLike);
    }

    private void applyDayStringStyle(
            Cell targetDayCell,
            String value,
            CellStyle centerStyle,
            CellStyle dayValueStyle,
            CellStyle emptyDayStyle,
            CellStyle vacanceStyle,
            CellStyle freedayStyle,
            CellStyle sickLeaveStyle,
            CellStyle legalAbsenceStyle,
            boolean adjustmentLike
    ) {
        String normalized = value == null ? "" : value.trim().toUpperCase();
        targetDayCell.setCellValue(value == null ? "" : value);
        switch (normalized) {
            case "V":
                targetDayCell.setCellStyle(vacanceStyle);
                break;
            case "F":
                targetDayCell.setCellStyle(freedayStyle);
                break;
            case "S":
                targetDayCell.setCellStyle(sickLeaveStyle);
                break;
            case "A":
                targetDayCell.setCellStyle(legalAbsenceStyle);
                break;
            default:
                if (normalized.isEmpty()) {
                    targetDayCell.setCellStyle(adjustmentLike ? centerStyle : emptyDayStyle);
                } else if (isNumericText(normalized)) {
                    targetDayCell.setCellStyle(dayValueStyle);
                } else {
                    targetDayCell.setCellStyle(adjustmentLike ? centerStyle : emptyDayStyle);
                }
                break;
        }
    }

    private boolean isAdjustmentLike(Row sourceRow) {
        Cell employeeCell = sourceRow.getCell(0);
        Cell categoryCell = sourceRow.getCell(2);
        boolean hasEmployeeId = employeeCell != null && employeeCell.getCellType() == CellType.NUMERIC;
        boolean hasCategory = categoryCell != null && !categoryCell.toString().trim().isEmpty();
        return !hasEmployeeId && !hasCategory;
    }

    private boolean isNumericText(String value) {
        try {
            Double.parseDouble(value);
            return true;
        } catch (NumberFormatException ex) {
            return false;
        }
    }

    private void writeCompactDataRow(
            Row outputRow,
            Row sourceRow,
            CellStyle leftStyle,
            CellStyle centerStyle,
            CellStyle currencyStyle,
            int transformedHoursCol,
            int transformedCostCol,
            int nbrDaysInThisMonth
    ) {
        Cell emplCell = outputRow.createCell(0);
        Cell srcEmp = sourceRow.getCell(0);
        if (srcEmp != null && srcEmp.getCellType() == CellType.NUMERIC) {
            emplCell.setCellValue(srcEmp.getNumericCellValue());
        }
        emplCell.setCellStyle(centerStyle);

        Cell personCell = outputRow.createCell(1);
        personCell.setCellValue(sourceRow.getCell(1) != null ? sourceRow.getCell(1).toString() : "");
        personCell.setCellStyle(leftStyle);

        Cell categoryCell = outputRow.createCell(2);
        categoryCell.setCellValue(sourceRow.getCell(2) != null ? sourceRow.getCell(2).toString() : "");
        categoryCell.setCellStyle(leftStyle);

        double rate = getNumericCellValue(sourceRow.getCell(3));
        Cell rateCell = outputRow.createCell(3);
        if (rate != 0) {
            rateCell.setCellValue(Helper.round(rate));
        } else {
            rateCell.setCellValue("");
        }
        rateCell.setCellStyle(currencyStyle);

        double hours = getWorkingHoursFromTransformedRow(sourceRow, transformedHoursCol, nbrDaysInThisMonth);
        Cell hoursCell = outputRow.createCell(4);
        hoursCell.setCellValue(hours);
        hoursCell.setCellStyle(centerStyle);

        Cell costCell = outputRow.createCell(5);
        costCell.setCellValue(getCostFromTransformedRow(sourceRow, transformedCostCol, rate, hours));
        costCell.setCellStyle(currencyStyle);
    }

    private double getWorkingHoursFromTransformedRow(Row row, int transformedHoursCol, int nbrDaysInThisMonth) {
        double explicitHours = getNumericCellValue(row.getCell(transformedHoursCol));
        if (explicitHours != 0) {
            return Helper.round(explicitHours);
        }

        double summedHours = 0;
        FormulaEvaluator evaluator = srcDayCellEvaluator(row);
        for (int col = 4; col < 4 + nbrDaysInThisMonth; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }
            if (cell.getCellType() == CellType.NUMERIC) {
                summedHours += cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.FORMULA) {
                CellValue evaluated = evaluator.evaluate(cell);
                if (evaluated != null && evaluated.getCellType() == CellType.NUMERIC) {
                    summedHours += evaluated.getNumberValue();
                }
            }
        }
        return Helper.round(summedHours);
    }

    private double getCostFromTransformedRow(Row row, int transformedCostCol, double rate, double hours) {
        double explicitCost = getNumericCellValue(row.getCell(transformedCostCol));
        if (explicitCost != 0) {
            return Helper.round(explicitCost);
        }
        return Helper.round(rate * hours);
    }

    private FormulaEvaluator srcDayCellEvaluator(Row srcRow) {
        return srcRow.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
    }

    private double getNumericCellValue(Cell cell) {
        if (cell == null) {
            return 0;
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (cell.getCellType() == CellType.FORMULA) {
            FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellValue evaluated = evaluator.evaluate(cell);
            if (evaluated != null && evaluated.getCellType() == CellType.NUMERIC) {
                return evaluated.getNumberValue();
            }
        }
        return 0;
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
            BigDecimal key = entry.getKey();
            List<Row> rows = entry.getValue();

            if (rows == null || rows.isEmpty()) {
                continue;
            }

            Workbook workbook = new XSSFWorkbook();
            Row newRow = workbook.createSheet().createRow(0);

            Cell firstCell = rows.get(0).getCell(0);
            newRow.createCell(0).setCellValue(firstCell != null ? firstCell.getNumericCellValue() : 0);

            if (rows.get(0).getCell(1) != null) {
                newRow.createCell(1).setCellValue(rows.get(0).getCell(1).getStringCellValue());
            }

            if (rows.size() > 1 && rows.get(1) != null && rows.get(1).getCell(1) != null) {
                Cell secondCellSecond = rows.get(1).getCell(1);
                BigDecimal input = BigDecimal.valueOf(Helper.getRates(secondCellSecond.getStringCellValue()));
                List<String> groupsId = CogsHelper.findGroupIdsByRate(input, FiscalYear.FY25, recogs);
                newRow.createCell(2).setCellValue(groupsId.toString());
            }

            CellStyle currencyStyle = Helper.getCurrencyStyle(workbook);
            if (rows.size() > 1 && rows.get(1).getCell(1) != null) {
                Cell secondCellSecond = rows.get(1).getCell(1);
                BigDecimal input = BigDecimal.valueOf(Helper.getRates(secondCellSecond.getStringCellValue()));
                String description = rows.get(0).getCell(1) != null
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
                Row vacationRow = rows.size() > 2 ? rows.get(2) : null;
                for (int i = 4; i < teamServiceRow.getLastCellNum() + 4; i++) {
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
                                CellValue cellValue = evaluator.evaluate(hoursCell);
                                BigDecimal numericValue = BigDecimal.valueOf(cellValue.getNumberValue());
                                outputCell.setCellValue(Helper.round(numericValue.doubleValue()));
                                break;
                            default:
                                outputCell.setCellValue(hoursCell.getStringCellValue());
                                break;
                        }
                    }

                    if ((outputCell.getCellType() == CellType.BLANK)
                            || (outputCell.getCellType() == CellType.STRING && outputCell.getStringCellValue().isEmpty())) {
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

            if (newRow.getCell(2) != null) {
                resultMap.put(key, newRow);
            }
        }

        return resultMap;
    }

    public BigDecimal getTotalServiceTeam(Workbook inputWorkbook, String serviceTeam, String sheetName) {
        Sheet sheet = inputWorkbook.getSheet(sheetName);
        FormulaEvaluator evaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();
        if (sheet == null) {
            return BigDecimal.ZERO;
        }

        BigDecimal total = BigDecimal.ZERO;
        boolean inProjectBlock = false;
        String projectBlock = "";
        for (Row row : sheet) {
            Cell projectCell = row.getCell(1);
            Cell cell0 = row.getCell(0);
            projectBlock = cell0 != null && CellType.STRING.equals(cell0.getCellType())
                    && cell0.getStringCellValue() != null && !cell0.getStringCellValue().isEmpty()
                    && cell0.getStringCellValue().equals("NÃºmero Empleado")
                    ? projectCell.getStringCellValue() : projectBlock;
            Cell totalCell = row.getCell(7);
            String project = projectCell != null ? (projectCell.getStringCellValue() != null ? projectCell.getStringCellValue().trim() : "") : "";
            BigDecimal val = totalCell != null ? BigDecimal.valueOf(evaluator.evaluate(totalCell).getNumberValue()) : BigDecimal.ZERO;
            if (project.isEmpty() && val.compareTo(BigDecimal.ZERO) != 0 && projectBlock.contains(serviceTeam)) {
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
        if (sheet == null) {
            return BigDecimal.ZERO;
        }

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
