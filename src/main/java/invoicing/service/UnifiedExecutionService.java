package invoicing.service;

import invoicing.Helper.GroupAggregator;
import invoicing.Helper.Helper;
import invoicing.Helper.ReferenceData;
import invoicing.entities.ServiceTeam;
import invoicing.service.ext.ExcelReader;
import invoicing.service.ext.ServiceTeamParser;
import invoicing.service.month.ExecuteService;
import invoicing.service.rate.InputFilesReader;
import invoicing.service.rate.InputRowProcessor;
import invoicing.service.rate.OutputWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.Normalizer;
import java.time.Month;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

public class UnifiedExecutionService {
    private static final int MERGED_FILE_SEPARATOR_ROWS = 3;

    public interface Listener {
        void log(String message);

        void setProgress(int value, String barLabel, String detail);
    }

    public File runUnified(File targetDir, List<File> inputs, int months, boolean useManual, Listener listener) {
        LocalDateTime now = LocalDateTime.now();
        List<String> monthNamesSpa = detectRequestedMonthsSpanish(inputs, months, useManual, listener);
        String currentMonthStr = buildPeriodToken(monthNamesSpa, now.getYear());
        String periodDisplay = buildPeriodDisplay(monthNamesSpa, now.getYear());
        String runStamp = now.format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

        File mainOutputFolder = new File(targetDir, "forecast_italy_" + currentMonthStr + "_" + runStamp);
        mainOutputFolder.mkdirs();

        File rateFolder = new File(mainOutputFolder, "forecast_it_rate_" + currentMonthStr);
        File extFolder = new File(mainOutputFolder, "forecast_EXT_" + currentMonthStr);
        File monthFolder = new File(mainOutputFolder, "forecast_month_" + currentMonthStr);
        rateFolder.mkdirs();
        extFolder.mkdirs();
        monthFolder.mkdirs();

        listener.log("=== STARTING UNIFIED EXECUTION ===");
        listener.log("Output Folder : " + mainOutputFolder.getAbsolutePath());
        listener.log("Months mode   : " + (useManual ? "MANUAL (" + months + " months)" : "AUTO-DETECT from Facturacion sheets"));
        listener.log("Detected months: " + (monthNamesSpa.isEmpty() ? "(none)" : String.join(", ", monthNamesSpa)));

        runRateModule(now, rateFolder, inputs, months, useManual, periodDisplay, listener);
        runExtModule(extFolder, inputs, months, useManual, periodDisplay, listener);
        runMonthModule(monthFolder, inputs, months, useManual, listener);

        listener.setProgress(3, "Completed", "All modules finished successfully.");
        listener.log("\n=== EXECUTION COMPLETED ===");

        return mainOutputFolder;
    }

    private void runRateModule(LocalDateTime now, File rateFolder, List<File> inputs, int months, boolean useManual, String periodDisplay, Listener listener) {
        listener.setProgress(0, "Step 1/3 - Rate", "Running Forecast By Rate...");
        listener.log("\n[1/3] Running Forecast By Rate...");
        try {
            List<String> monthNamesSpa = detectRequestedMonthsSpanish(inputs, months, useManual, listener);
            if (monthNamesSpa.isEmpty()) {
                listener.log("  - Rate Warning: No requested Facturacion month sheets found.");
                return;
            }

            try (Workbook outputWorkbook = new XSSFWorkbook()) {
                for (String monthSpa : monthNamesSpa) {
                    ReferenceData referenceData = new ReferenceData();
                    try (InputStream dataStream = getClass().getClassLoader().getResourceAsStream("Data.xlsx")) {
                        if (dataStream == null) {
                            listener.log("  ! Rate Error: Data.xlsx not found inside the JAR. Check build resources.");
                            return;
                        }
                        referenceData.load(dataStream);
                    } catch (Exception e) {
                        listener.log("  ! Rate Error loading Data.xlsx: " + e.getMessage());
                        return;
                    }

                    GroupAggregator aggregator = new GroupAggregator();
                    InputRowProcessor rowProcessor = new InputRowProcessor(referenceData);
                    InputFilesReader filesReader = new InputFilesReader(rowProcessor, aggregator);

                    for (File f : inputs) {
                        try {
                            boolean found = filesReader.processFile(f.getAbsolutePath(), monthSpa);
                            if (!found) {
                                listener.log("  - Rate Info: Facturacion " + monthSpa + " not found in " + f.getName());
                            }
                        } catch (Exception e) {
                            listener.log("  - Rate Warning: Failed to process " + f.getName() + " for " + monthSpa);
                        }
                    }

                    OutputWriter writer = new OutputWriter(referenceData, aggregator);
                    writer.writeSheet(outputWorkbook, toEnglishMonthName(monthSpa, now), filesReader.getProcessedRows());
                }

                String rateOut = new File(rateFolder, "Rate Forecast " + periodDisplay + ".xlsx").getAbsolutePath();
                try (FileOutputStream fos = new FileOutputStream(rateOut)) {
                    outputWorkbook.write(fos);
                }
                listener.log("  > Rate Report created: " + rateOut);
            }
        } catch (Exception e) {
            listener.log("  ! Rate Module Failed: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void runExtModule(File extFolder, List<File> inputs, int months, boolean useManual, String periodDisplay, Listener listener) {
        listener.setProgress(1, "Step 2/3 - ExtCode", "Running Forecast By ExtCode...");
        listener.log("\n[2/3] Running Forecast By ExtCode...");
        try {
            List<String> monthNamesSpa = detectRequestedMonthsSpanish(inputs, months, useManual, listener);
            if (monthNamesSpa.isEmpty()) {
                listener.log("  - ExtCode Warning: No requested Facturacion month sheets found.");
                return;
            }

            ExcelReader reader = new ExcelReader();
            ServiceTeamParser parser = new ServiceTeamParser();
            invoicing.service.ext.ExcelWriter writer = new invoicing.service.ext.ExcelWriter();

            try (Workbook outputWorkbook = new XSSFWorkbook()) {
                for (String monthSpa : monthNamesSpa) {
                    List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();
                    for (File f : inputs) {
                        try {
                            List<ExcelReader.ServiceTeamRaw> monthRaw = reader.extractRawServiceTeams(f, monthSpa);
                            if (monthRaw.isEmpty()) {
                                listener.log("  - ExtCode Info: Facturacion " + monthSpa + " not found in " + f.getName());
                            }
                            rawItems.addAll(monthRaw);
                        } catch (Exception e) {
                            listener.log("  - ExtCode Warning: Failed to process " + f.getName() + " for " + monthSpa);
                        }
                    }

                    List<String> labels = new ArrayList<>();
                    for (ExcelReader.ServiceTeamRaw raw : rawItems) {
                        labels.add(raw.getLabel());
                    }
                    List<ServiceTeam> parsed = parser.parse(labels);
                    for (int i = 0; i < parsed.size(); i++) {
                        parsed.get(i).setCost(rawItems.get(i).getCost() == null ? "" : String.valueOf(rawItems.get(i).getCost()));
                        parsed.get(i).setStyle(rawItems.get(i).getCost() == null ? null : rawItems.get(i).getStyle());
                    }
                    writer.writeSheet(outputWorkbook, toEnglishMonthName(monthSpa, LocalDateTime.now()), parsed);
                }

                File file = new File(extFolder, "ForeCast IT " + periodDisplay + ".xlsx");
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    outputWorkbook.write(fos);
                }
                listener.log("  > ExtCode Report created in: " + extFolder.getAbsolutePath());
            }
        } catch (Exception e) {
            listener.log("  ! ExtCode Module Failed: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void runMonthModule(File monthFolder, List<File> inputs, int months, boolean useManual, Listener listener) {
        listener.setProgress(2, "Step 3/3 - Month", "Running Forecast By Month...");
        listener.log("\n[3/3] Running Forecast By Month...");
        try {
            for (File f : inputs) {
                try {
                    int currentMonths = useManual ? months : countMonthSheets(f, listener);
                    if (currentMonths <= 0) {
                        listener.log("  - Month Warning: No Facturacion sheets found in " + f.getName() + ". Skipping file.");
                        continue;
                    }
                    listener.log("  - Processing " + f.getName() + " with " + currentMonths + " months...");
                    ExecuteService.executeScript(
                            f.getAbsolutePath(),
                            monthFolder.getAbsolutePath(),
                            currentMonths,
                            listener::log
                    );
                } catch (Exception e) {
                    listener.log("  - Month Warning: Failed to process " + f.getName() + ": " + e.getMessage());
                }
            }

            mergeConsolidatedMonthFiles(monthFolder, listener);
            listener.log("  > Month processing finished.");
        } catch (Exception e) {
            listener.log("  ! Month Module Critical Error: " + e.getMessage());
        }
    }

    private void mergeConsolidatedMonthFiles(File monthFolder, Listener listener) {
        File[] consolidatedFiles = monthFolder.listFiles((dir, name) ->
                name != null
                        && name.startsWith("Consolidated_Month_Forecast_")
                        && name.toLowerCase().endsWith(".xlsx")
                        && !name.contains("_ALL_PROJECTS"));

        if (consolidatedFiles == null || consolidatedFiles.length <= 1) {
            return;
        }

        Arrays.sort(consolidatedFiles, Comparator.comparing(File::getName));
        String monthFolderName = monthFolder.getName();
        String periodToken = monthFolderName.startsWith("forecast_month_")
                ? monthFolderName.substring("forecast_month_".length())
                : "period";
        File mergedFile = new File(monthFolder, "Consolidated_Month_Forecast_ALL_PROJECTS_" + periodToken + ".xlsx");

        try (Workbook mergedWorkbook = new XSSFWorkbook()) {
            CellStyle headerStyle = Helper.getHeaderStyle(mergedWorkbook);
            CellStyle projectBandStyle = Helper.getProjectBandStyle(mergedWorkbook);
            CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(mergedWorkbook);
            CellStyle leftStyle = Helper.getLeftStandardStyle(mergedWorkbook);
            CellStyle centerStyle = Helper.getCenterStandardStyle(mergedWorkbook);
            CellStyle currencyStyle = Helper.getCurrencyStyle(mergedWorkbook);
            CellStyle categoryStyle = Helper.getCategoryStyle(mergedWorkbook);
            CellStyle dayValueStyle = Helper.getDayValueStyle(mergedWorkbook);
            CellStyle emptyDayStyle = Helper.getEmptyDayStyle(mergedWorkbook);
            CellStyle dateStyle = Helper.getDateStyle(mergedWorkbook);
            CellStyle vacanceStyle = Helper.getVacanceStyle(mergedWorkbook);
            CellStyle freedayStyle = Helper.getFreedayStyle(mergedWorkbook);
            CellStyle sickLeaveStyle = Helper.getSickLeaveStyle(mergedWorkbook);
            CellStyle legalAbsenceStyle = Helper.getLegalAbsenceStyle(mergedWorkbook);

            Map<String, Sheet> targetSheets = new LinkedHashMap<>();
            Map<String, Integer> sheetRowIndex = new LinkedHashMap<>();
            Map<String, Double> sheetHoursTotal = new LinkedHashMap<>();
            Map<String, Double> sheetCostTotal = new LinkedHashMap<>();
            Map<String, Integer> sheetRateCol = new LinkedHashMap<>();
            Map<String, Integer> sheetHoursCol = new LinkedHashMap<>();
            Map<String, Integer> sheetCostCol = new LinkedHashMap<>();
            Map<String, Integer> sheetMaxCol = new LinkedHashMap<>();

            for (File file : consolidatedFiles) {
                try (Workbook sourceWorkbook = WorkbookFactory.create(file)) {
                    List<Sheet> sourceSheets = new ArrayList<>();
                    for (int s = 0; s < sourceWorkbook.getNumberOfSheets(); s++) {
                        Sheet candidate = sourceWorkbook.getSheetAt(s);
                        if (candidate.getSheetName().startsWith("All Teams Forecast")) {
                            sourceSheets.add(candidate);
                        }
                    }
                    if (sourceSheets.isEmpty()) {
                        if (sourceWorkbook.getNumberOfSheets() == 0) {
                            listener.log("  - Month Merge Warning: " + file.getName() + " has no sheets.");
                            continue;
                        }
                        sourceSheets.add(sourceWorkbook.getSheetAt(0));
                    }

                    FormulaEvaluator evaluator = sourceWorkbook.getCreationHelper().createFormulaEvaluator();
                    for (Sheet sourceSheet : sourceSheets) {
                        String sheetKey = sourceSheet.getSheetName();
                        Sheet mergedSheet = targetSheets.get(sheetKey);
                        if (mergedSheet == null) {
                            String targetName = sheetKey.length() > 31 ? sheetKey.substring(0, 31) : sheetKey;
                            mergedSheet = mergedWorkbook.createSheet(targetName);
                            targetSheets.put(sheetKey, mergedSheet);
                            sheetRowIndex.put(sheetKey, 0);
                            sheetHoursTotal.put(sheetKey, 0.0);
                            sheetCostTotal.put(sheetKey, 0.0);
                            sheetRateCol.put(sheetKey, 3);
                            sheetHoursCol.put(sheetKey, 4);
                            sheetCostCol.put(sheetKey, 5);
                            sheetMaxCol.put(sheetKey, 5);
                        }

                        int mergedRowIndex = sheetRowIndex.get(sheetKey);
                        int activeRateCol = sheetRateCol.get(sheetKey);
                        int activeHoursCol = sheetHoursCol.get(sheetKey);
                        int activeCostCol = sheetCostCol.get(sheetKey);
                        int activeMaxCol = sheetMaxCol.get(sheetKey);

                        double runningHours = sheetHoursTotal.get(sheetKey);
                        double runningCost = sheetCostTotal.get(sheetKey);

                        for (int r = 0; r <= sourceSheet.getLastRowNum(); r++) {
                            Row sourceRow = sourceSheet.getRow(r);
                            if (sourceRow == null) {
                                continue;
                            }

                            if (isHeaderRow(sourceRow, evaluator)) {
                                int detectedRateCol = findColumnIndexByHeader(sourceRow, evaluator, "rate");
                                int detectedHoursCol = findColumnIndexByHeader(sourceRow, evaluator, "working hours");
                                int detectedCostCol = findColumnIndexByHeader(sourceRow, evaluator, "cost");
                                if (detectedRateCol >= 0) {
                                    activeRateCol = detectedRateCol;
                                }
                                if (detectedHoursCol >= 0) {
                                    activeHoursCol = detectedHoursCol;
                                }
                                if (detectedCostCol >= 0) {
                                    activeCostCol = detectedCostCol;
                                }
                            }

                            if (isGrandTotalRow(sourceRow, evaluator)) {
                                runningHours += getNumericCellValue(sourceRow.getCell(activeHoursCol), evaluator);
                                runningCost += getNumericCellValue(sourceRow.getCell(activeCostCol), evaluator);
                                continue;
                            }

                            Row targetRow = mergedSheet.createRow(mergedRowIndex++);
                            short lastCellNum = sourceRow.getLastCellNum();
                            if (lastCellNum <= 0) {
                                continue;
                            }

                            for (int c = 0; c < lastCellNum; c++) {
                                Cell sourceCell = sourceRow.getCell(c);
                                if (sourceCell == null) {
                                    continue;
                                }
                                Cell targetCell = targetRow.createCell(c);
                                copyCellValue(sourceCell, targetCell, evaluator);
                            }

                            if (lastCellNum - 1 > activeMaxCol) {
                                activeMaxCol = lastCellNum - 1;
                            }

                            applyMergedRowStyle(
                                    targetRow,
                                    sourceRow,
                                    evaluator,
                                    projectBandStyle,
                                    headerStyle,
                                    footerCurrencyStyle,
                                    leftStyle,
                                    centerStyle,
                                    currencyStyle,
                                    categoryStyle,
                                    dayValueStyle,
                                    emptyDayStyle,
                                    dateStyle,
                                    vacanceStyle,
                                    freedayStyle,
                                    sickLeaveStyle,
                                    legalAbsenceStyle,
                                    activeRateCol,
                                    activeHoursCol,
                                    activeCostCol,
                                    activeMaxCol
                            );

                            if (isProjectBandRow(sourceRow, evaluator)) {
                                mergedSheet.addMergedRegion(new CellRangeAddress(
                                        targetRow.getRowNum(),
                                        targetRow.getRowNum(),
                                        0,
                                        2
                                ));
                            }
                        }

                        for (int i = 0; i < MERGED_FILE_SEPARATOR_ROWS; i++) {
                            mergedSheet.createRow(mergedRowIndex++);
                        }

                        sheetRowIndex.put(sheetKey, mergedRowIndex);
                        sheetHoursTotal.put(sheetKey, runningHours);
                        sheetCostTotal.put(sheetKey, runningCost);
                        sheetRateCol.put(sheetKey, activeRateCol);
                        sheetHoursCol.put(sheetKey, activeHoursCol);
                        sheetCostCol.put(sheetKey, activeCostCol);
                        sheetMaxCol.put(sheetKey, activeMaxCol);
                    }
                } catch (Exception e) {
                    listener.log("  - Month Merge Warning: Failed to merge " + file.getName() + ": " + e.getMessage());
                }
            }

            for (Map.Entry<String, Sheet> entry : targetSheets.entrySet()) {
                String sheetKey = entry.getKey();
                Sheet mergedSheet = entry.getValue();
                int mergedRowIndex = sheetRowIndex.get(sheetKey);
                double allProjectsHours = sheetHoursTotal.get(sheetKey);
                double allProjectsCost = sheetCostTotal.get(sheetKey);
                int rateCol = sheetRateCol.get(sheetKey);
                int hoursCol = sheetHoursCol.get(sheetKey);
                int costCol = sheetCostCol.get(sheetKey);
                int maxCol = sheetMaxCol.get(sheetKey);

                if (allProjectsHours != 0 || allProjectsCost != 0) {
                    mergedRowIndex++;
                    Row grandTotalRow = mergedSheet.createRow(mergedRowIndex);

                    int grandLabelStartCol = Math.max(2, hoursCol - 2);
                    int grandLabelEndCol = Math.max(grandLabelStartCol, hoursCol - 1);
                    Cell labelCell = grandTotalRow.createCell(grandLabelStartCol);
                    labelCell.setCellValue("GRAND TOTAL (ALL PROJECTS)");
                    labelCell.setCellStyle(headerStyle);
                    mergedSheet.addMergedRegion(new CellRangeAddress(
                            grandTotalRow.getRowNum(),
                            grandTotalRow.getRowNum(),
                            grandLabelStartCol,
                            grandLabelEndCol
                    ));

                    Cell hoursCell = grandTotalRow.createCell(hoursCol);
                    hoursCell.setCellValue(Helper.round(allProjectsHours));
                    hoursCell.setCellStyle(headerStyle);

                    Cell costCell = grandTotalRow.createCell(costCol);
                    costCell.setCellValue(Helper.round(allProjectsCost));
                    costCell.setCellStyle(footerCurrencyStyle);
                }

                for (int i = 0; i <= maxCol; i++) {
                    if (i == 0) {
                        mergedSheet.setColumnWidth(i, 11 * 256);
                    } else if (i == 1) {
                        mergedSheet.setColumnWidth(i, 40 * 256);
                    } else if (i == 2) {
                        mergedSheet.setColumnWidth(i, 20 * 256);
                    } else if (i == rateCol) {
                        mergedSheet.setColumnWidth(i, 9 * 256);
                    } else if (i > rateCol && i < hoursCol) {
                        mergedSheet.setColumnWidth(i, 5 * 256);
                    } else if (i == hoursCol || i == costCol) {
                        mergedSheet.setColumnWidth(i, 13 * 256);
                    } else {
                        mergedSheet.autoSizeColumn(i);
                    }
                }
                mergedSheet.setDisplayGridlines(true);
            }

            try (FileOutputStream fos = new FileOutputStream(mergedFile)) {
                mergedWorkbook.write(fos);
            }

            for (File file : consolidatedFiles) {
                if (!file.getAbsolutePath().equals(mergedFile.getAbsolutePath()) && !file.delete()) {
                    listener.log("  - Month Merge Warning: Could not delete old consolidated file: " + file.getName());
                }
            }

            listener.log("  > Month consolidated merge created: " + mergedFile.getAbsolutePath());
        } catch (Exception e) {
            listener.log("  ! Month Merge Failed: " + e.getMessage());
        }
    }

    private void applyMergedRowStyle(Row targetRow, Row sourceRow, FormulaEvaluator evaluator,
                                     CellStyle projectBandStyle, CellStyle headerStyle, CellStyle footerCurrencyStyle,
                                     CellStyle leftStyle, CellStyle centerStyle, CellStyle currencyStyle,
                                     CellStyle categoryStyle, CellStyle dayValueStyle, CellStyle emptyDayStyle, CellStyle dateStyle,
                                     CellStyle vacanceStyle, CellStyle freedayStyle, CellStyle sickLeaveStyle, CellStyle legalAbsenceStyle,
                                     int rateCol, int hoursCol, int costCol, int maxCol) {
        boolean projectRow = isProjectBandRow(sourceRow, evaluator);
        boolean headerRow = rowContainsToken(sourceRow, evaluator, "empl");
        boolean totalRow = rowContainsToken(sourceRow, evaluator, "total");
        boolean adjustmentLike = isAdjustmentLike(sourceRow);

        for (int c = 0; c <= maxCol; c++) {
            Cell cell = targetRow.getCell(c);
            if (cell == null) {
                cell = targetRow.createCell(c);
            }

            if (projectRow) {
                if (c <= 2) {
                    cell.setCellStyle(projectBandStyle);
                }
                continue;
            }

            if (headerRow) {
                if (c > rateCol && c < hoursCol) {
                    cell.setCellStyle(dateStyle);
                } else {
                    cell.setCellStyle(headerStyle);
                }
                continue;
            }

            if (totalRow) {
                int totalLabelStartCol = Math.max(2, hoursCol - 2);
                if (c < totalLabelStartCol) {
                    cell.setCellStyle(centerStyle);
                } else if (c == costCol) {
                    cell.setCellStyle(footerCurrencyStyle);
                } else {
                    cell.setCellStyle(headerStyle);
                }
                continue;
            }

            if (c == 0) {
                cell.setCellStyle(centerStyle);
            } else if (c == 1) {
                cell.setCellStyle(leftStyle);
            } else if (c == 2) {
                cell.setCellStyle(categoryStyle);
            } else if (c == rateCol || c == costCol) {
                cell.setCellStyle(currencyStyle);
            } else if (c == hoursCol) {
                cell.setCellStyle(centerStyle);
            } else if (c > rateCol && c < hoursCol) {
                applyMergedDayStyle(cell, adjustmentLike, centerStyle, dayValueStyle, emptyDayStyle, vacanceStyle, freedayStyle, sickLeaveStyle, legalAbsenceStyle);
            } else {
                cell.setCellStyle(centerStyle);
            }
        }
    }

    private void applyMergedDayStyle(Cell cell, boolean adjustmentLike,
                                     CellStyle centerStyle, CellStyle dayValueStyle, CellStyle emptyDayStyle,
                                     CellStyle vacanceStyle, CellStyle freedayStyle,
                                     CellStyle sickLeaveStyle, CellStyle legalAbsenceStyle) {
        String value = cell == null ? "" : cell.toString();
        String normalized = value.trim().toUpperCase(Locale.ROOT);
        switch (normalized) {
            case "V":
                cell.setCellStyle(vacanceStyle);
                break;
            case "F":
                cell.setCellStyle(freedayStyle);
                break;
            case "S":
                cell.setCellStyle(sickLeaveStyle);
                break;
            case "A":
                cell.setCellStyle(legalAbsenceStyle);
                break;
            default:
                if (normalized.isEmpty()) {
                    cell.setCellStyle(adjustmentLike ? centerStyle : emptyDayStyle);
                } else if (isNumericText(normalized)) {
                    cell.setCellStyle(dayValueStyle);
                } else {
                    cell.setCellStyle(adjustmentLike ? centerStyle : emptyDayStyle);
                }
                break;
        }
    }

    private boolean isProjectBandRow(Row row, FormulaEvaluator evaluator) {
        if (row == null || isHeaderRow(row, evaluator) || isGrandTotalRow(row, evaluator)) {
            return false;
        }

        String firstCell = getStringCellValue(row.getCell(0), evaluator).trim();
        if (firstCell.isEmpty()) {
            return false;
        }

        String secondCell = getStringCellValue(row.getCell(1), evaluator).trim();
        String thirdCell = getStringCellValue(row.getCell(2), evaluator).trim();
        String fourthCell = getStringCellValue(row.getCell(3), evaluator).trim();
        return secondCell.isEmpty() && thirdCell.isEmpty() && fourthCell.isEmpty();
    }

    private boolean isAdjustmentLike(Row row) {
        Cell employeeCell = row.getCell(0);
        Cell categoryCell = row.getCell(2);
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

    private boolean isHeaderRow(Row row, FormulaEvaluator evaluator) {
        return rowContainsToken(row, evaluator, "empl");
    }

    private int findColumnIndexByHeader(Row row, FormulaEvaluator evaluator, String token) {
        short last = row.getLastCellNum();
        if (last <= 0) {
            return -1;
        }
        String normalizedToken = token.toLowerCase(Locale.ROOT);
        for (int i = 0; i < last; i++) {
            String value = getStringCellValue(row.getCell(i), evaluator).trim().toLowerCase(Locale.ROOT);
            if (!value.isEmpty() && value.contains(normalizedToken)) {
                return i;
            }
        }
        return -1;
    }

    private boolean isGrandTotalRow(Row row, FormulaEvaluator evaluator) {
        return rowContainsToken(row, evaluator, "grand total");
    }

    private boolean rowContainsToken(Row row, FormulaEvaluator evaluator, String token) {
        if (row == null) {
            return false;
        }
        String normalizedToken = token.toLowerCase(Locale.ROOT);
        short last = row.getLastCellNum();
        if (last <= 0) {
            return false;
        }
        for (int i = 0; i < last; i++) {
            String value = getStringCellValue(row.getCell(i), evaluator).trim().toLowerCase(Locale.ROOT);
            if (!value.isEmpty() && value.contains(normalizedToken)) {
                return true;
            }
        }
        return false;
    }

    private String getStringCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }

        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            CellValue evaluated = evaluator.evaluate(cell);
            if (evaluated == null) {
                return "";
            }
            if (evaluated.getCellType() == CellType.STRING) {
                return evaluated.getStringValue();
            }
            if (evaluated.getCellType() == CellType.NUMERIC) {
                return String.valueOf(evaluated.getNumberValue());
            }
            if (evaluated.getCellType() == CellType.BOOLEAN) {
                return String.valueOf(evaluated.getBooleanValue());
            }
            return "";
        }

        if (type == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (type == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        }
        if (type == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        }
        return "";
    }

    private double getNumericCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return 0;
        }
        CellType type = cell.getCellType();
        if (type == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (type == CellType.FORMULA) {
            CellValue evaluated = evaluator.evaluate(cell);
            if (evaluated != null && evaluated.getCellType() == CellType.NUMERIC) {
                return evaluated.getNumberValue();
            }
        }
        return 0;
    }

    private void copyCellValue(Cell sourceCell, Cell targetCell, FormulaEvaluator evaluator) {
        CellType sourceType = sourceCell.getCellType();
        if (sourceType == CellType.FORMULA) {
            CellValue evaluated = evaluator.evaluate(sourceCell);
            if (evaluated == null) {
                targetCell.setBlank();
                return;
            }
            switch (evaluated.getCellType()) {
                case STRING:
                    targetCell.setCellValue(evaluated.getStringValue());
                    return;
                case NUMERIC:
                    targetCell.setCellValue(evaluated.getNumberValue());
                    return;
                case BOOLEAN:
                    targetCell.setCellValue(evaluated.getBooleanValue());
                    return;
                default:
                    targetCell.setBlank();
                    return;
            }
        }

        switch (sourceType) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            default:
                targetCell.setBlank();
                break;
        }
    }

    private int countMonthSheets(File f, Listener listener) {
        int count = 0;
        try (Workbook wb = WorkbookFactory.create(f)) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String n = Normalizer.normalize(wb.getSheetName(i).toLowerCase(), Normalizer.Form.NFD)
                        .replaceAll("\\p{M}+", "");
                if (n.contains("facturacion")) {
                    count++;
                }
            }
        } catch (Exception e) {
            listener.log("  - Error counting sheets in " + f.getName() + ": " + e.getMessage());
        }
        return count;
    }

    private List<String> detectRequestedMonthsSpanish(List<File> inputs, int requestedMonths, boolean useManual, Listener listener) {
        Set<String> months = new LinkedHashSet<>();
        for (File f : inputs) {
            try (Workbook wb = WorkbookFactory.create(f)) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    String sheetName = wb.getSheetName(i);
                    String normalized = normalize(sheetName);
                    if (!normalized.startsWith("facturacion ")) {
                        continue;
                    }
                    String month = normalized.substring("facturacion ".length()).trim();
                    if (!month.isEmpty()) {
                        months.add(month);
                    }
                }
            } catch (Exception e) {
                listener.log("  - Warning: Failed month detection in " + f.getName() + ": " + e.getMessage());
            }
            if (useManual && months.size() >= requestedMonths) {
                break;
            }
        }

        List<String> ordered = new ArrayList<>(months);
        if (useManual && ordered.size() > requestedMonths) {
            return ordered.subList(0, requestedMonths);
        }
        return ordered;
    }

    private String toEnglishMonthName(String monthSpanish, LocalDateTime now) {
        int year = now.getYear();
        int monthNum = monthNumber(monthSpanish);
        if (monthNum < 1 || monthNum > 12) {
            return monthSpanish;
        }
        return Month.of(monthNum).getDisplayName(java.time.format.TextStyle.FULL, Locale.US) + " " + year;
    }

    private int monthNumber(String monthSpanish) {
        String m = normalize(monthSpanish);
        switch (m) {
            case "enero":
                return 1;
            case "febrero":
                return 2;
            case "marzo":
                return 3;
            case "abril":
                return 4;
            case "mayo":
                return 5;
            case "junio":
                return 6;
            case "julio":
                return 7;
            case "agosto":
                return 8;
            case "septiembre":
                return 9;
            case "octubre":
                return 10;
            case "noviembre":
                return 11;
            case "diciembre":
                return 12;
            default:
                return -1;
        }
    }

    private String normalize(String value) {
        if (value == null) {
            return "";
        }
        return Normalizer.normalize(value, Normalizer.Form.NFD)
                .replaceAll("\\p{M}+", "")
                .toLowerCase(Locale.ROOT);
    }

    private String buildPeriodToken(List<String> monthNamesSpa, int year) {
        if (monthNamesSpa == null || monthNamesSpa.isEmpty()) {
            return LocalDateTime.now().format(DateTimeFormatter.ofPattern("MMM_yyyy"));
        }
        String joinedMonths = String.join("_", monthNamesSpa).replace(" ", "_");
        return joinedMonths + "_" + year;
    }

    private String buildPeriodDisplay(List<String> monthNamesSpa, int year) {
        if (monthNamesSpa == null || monthNamesSpa.isEmpty()) {
            return LocalDateTime.now().format(DateTimeFormatter.ofPattern("MMMM yyyy"));
        }
        List<String> englishMonths = new ArrayList<>();
        for (String monthSpa : monthNamesSpa) {
            int monthNum = monthNumber(monthSpa);
            if (monthNum >= 1 && monthNum <= 12) {
                englishMonths.add(Month.of(monthNum).getDisplayName(java.time.format.TextStyle.FULL, Locale.US));
            } else {
                englishMonths.add(monthSpa);
            }
        }
        return String.join("-", englishMonths) + " " + year;
    }
}
