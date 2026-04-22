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
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;

public class UnifiedExecutionService {
    private static final int MERGED_FILE_SEPARATOR_ROWS = 3;

    public interface Listener {
        void log(String message);

        void setProgress(int value, String barLabel, String detail);
    }

    public File runUnified(File targetDir, List<File> inputs, int months, boolean useManual, Listener listener) {
        LocalDateTime now = LocalDateTime.now();
        String currentMonthStr = now.format(DateTimeFormatter.ofPattern("MMM_yyyy"));
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

        runRateModule(now, rateFolder, inputs, listener);
        runExtModule(extFolder, inputs, listener);
        runMonthModule(monthFolder, inputs, months, useManual, listener);

        listener.setProgress(3, "Completed", "All modules finished successfully.");
        listener.log("\n=== EXECUTION COMPLETED ===");

        return mainOutputFolder;
    }

    private void runRateModule(LocalDateTime now, File rateFolder, List<File> inputs, Listener listener) {
        listener.setProgress(0, "Step 1/3 - Rate", "Running Forecast By Rate...");
        listener.log("\n[1/3] Running Forecast By Rate...");
        try {
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
                    filesReader.processFile(f.getAbsolutePath());
                } catch (Exception e) {
                    listener.log("  - Rate Warning: Failed to process " + f.getName());
                }
            }

            if (!aggregator.getAggregates().isEmpty()) {
                OutputWriter writer = new OutputWriter(referenceData, aggregator);
                String fullMonth = now.format(DateTimeFormatter.ofPattern("MMMM"));
                String rateOut = new File(rateFolder, "Rate Forecast " + fullMonth + ".xlsx").getAbsolutePath();
                writer.write(rateOut);
                listener.log("  > Rate Report created: " + rateOut);
            } else {
                listener.log("  - Rate Warning: No valid data found for Rate module.");
            }
        } catch (Exception e) {
            listener.log("  ! Rate Module Failed: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void runExtModule(File extFolder, List<File> inputs, Listener listener) {
        listener.setProgress(1, "Step 2/3 - ExtCode", "Running Forecast By ExtCode...");
        listener.log("\n[2/3] Running Forecast By ExtCode...");
        try {
            ExcelReader reader = new ExcelReader();
            ServiceTeamParser parser = new ServiceTeamParser();
            invoicing.service.ext.ExcelWriter writer = new invoicing.service.ext.ExcelWriter();

            List<ExcelReader.ServiceTeamRaw> rawItems = new ArrayList<>();
            for (File f : inputs) {
                try {
                    rawItems.addAll(reader.extractRawServiceTeams(f));
                } catch (Exception e) {
                    listener.log("  - ExtCode Warning: Failed to process " + f.getName());
                }
            }

            if (!rawItems.isEmpty()) {
                List<String> labels = new ArrayList<>();
                for (ExcelReader.ServiceTeamRaw raw : rawItems) {
                    labels.add(raw.getLabel());
                }
                List<ServiceTeam> parsed = parser.parse(labels);
                for (int i = 0; i < parsed.size(); i++) {
                    parsed.get(i).setCost(rawItems.get(i).getCost() == null ? "" : String.valueOf(rawItems.get(i).getCost()));
                    parsed.get(i).setStyle(rawItems.get(i).getCost() == null ? null : rawItems.get(i).getStyle());
                }
                writer.write(parsed, extFolder);
                listener.log("  > ExtCode Report created in: " + extFolder.getAbsolutePath());
            } else {
                listener.log("  - ExtCode Warning: No valid data found.");
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
                    int currentMonths = months;
                    if (!useManual) {
                        int detected = countMonthSheets(f, listener);
                        if (detected > 0) {
                            currentMonths = detected;
                            listener.log("  - Auto-detected months for " + f.getName() + ": " + currentMonths);
                        } else {
                            listener.log("  - Warning: No Facturacion sheets found in " + f.getName() + ". Using default: " + months);
                        }
                    }
                    listener.log("  - Processing " + f.getName() + " with " + currentMonths + " months...");
                    ExecuteService.executeScript(f.getAbsolutePath(), monthFolder.getAbsolutePath(), currentMonths);
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
        File mergedFile = new File(monthFolder, "Consolidated_Month_Forecast_ALL_PROJECTS.xlsx");

        try (Workbook mergedWorkbook = new XSSFWorkbook()) {
            Sheet mergedSheet = mergedWorkbook.createSheet("All Teams Forecast");
            CellStyle headerStyle = Helper.getHeaderStyle(mergedWorkbook);
            CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(mergedWorkbook);
            CellStyle leftStyle = Helper.getLeftStandardStyle(mergedWorkbook);
            CellStyle centerStyle = Helper.getCenterStandardStyle(mergedWorkbook);
            CellStyle currencyStyle = Helper.getCurrencyStyle(mergedWorkbook);

            int mergedRowIndex = 0;
            double allProjectsHours = 0;
            double allProjectsCost = 0;

            for (File file : consolidatedFiles) {
                try (Workbook sourceWorkbook = WorkbookFactory.create(file)) {
                    Sheet sourceSheet = sourceWorkbook.getSheet("All Teams Forecast");
                    if (sourceSheet == null) {
                        if (sourceWorkbook.getNumberOfSheets() == 0) {
                            listener.log("  - Month Merge Warning: " + file.getName() + " has no sheets.");
                            continue;
                        }
                        sourceSheet = sourceWorkbook.getSheetAt(0);
                    }

                    Row titleRow = mergedSheet.createRow(mergedRowIndex++);
                    Cell projectTitleCell = titleRow.createCell(0);
                    projectTitleCell.setCellValue("Project source: " + file.getName());
                    projectTitleCell.setCellStyle(headerStyle);
                    mergedSheet.addMergedRegion(new CellRangeAddress(
                            titleRow.getRowNum(),
                            titleRow.getRowNum(),
                            0,
                            5
                    ));

                    FormulaEvaluator evaluator = sourceWorkbook.getCreationHelper().createFormulaEvaluator();
                    for (int r = 0; r <= sourceSheet.getLastRowNum(); r++) {
                        Row sourceRow = sourceSheet.getRow(r);
                        if (sourceRow == null) {
                            continue;
                        }

                        if (isGrandTotalRow(sourceRow, evaluator)) {
                            allProjectsHours += getNumericCellValue(sourceRow.getCell(4), evaluator);
                            allProjectsCost += getNumericCellValue(sourceRow.getCell(5), evaluator);
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

                        applyMergedRowStyle(targetRow, sourceRow, evaluator, headerStyle, footerCurrencyStyle, leftStyle, centerStyle, currencyStyle);
                    }

                    for (int i = 0; i < MERGED_FILE_SEPARATOR_ROWS; i++) {
                        mergedSheet.createRow(mergedRowIndex++);
                    }
                } catch (Exception e) {
                    listener.log("  - Month Merge Warning: Failed to merge " + file.getName() + ": " + e.getMessage());
                }
            }

            if (allProjectsHours != 0 || allProjectsCost != 0) {
                mergedRowIndex++;
                Row grandTotalRow = mergedSheet.createRow(mergedRowIndex);

                Cell labelCell = grandTotalRow.createCell(0);
                labelCell.setCellValue("GRAND TOTAL (ALL PROJECTS)");
                labelCell.setCellStyle(headerStyle);
                mergedSheet.addMergedRegion(new CellRangeAddress(
                        grandTotalRow.getRowNum(),
                        grandTotalRow.getRowNum(),
                        0,
                        2
                ));

                Cell rateCell = grandTotalRow.createCell(3);
                if (allProjectsHours != 0) {
                    rateCell.setCellValue(Helper.round(allProjectsCost / allProjectsHours));
                } else {
                    rateCell.setCellValue("");
                }
                rateCell.setCellStyle(footerCurrencyStyle);

                Cell hoursCell = grandTotalRow.createCell(4);
                hoursCell.setCellValue(Helper.round(allProjectsHours));
                hoursCell.setCellStyle(headerStyle);

                Cell costCell = grandTotalRow.createCell(5);
                costCell.setCellValue(Helper.round(allProjectsCost));
                costCell.setCellStyle(footerCurrencyStyle);
            }

            for (int i = 0; i < 6; i++) {
                mergedSheet.autoSizeColumn(i);
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
                                     CellStyle headerStyle, CellStyle footerCurrencyStyle,
                                     CellStyle leftStyle, CellStyle centerStyle, CellStyle currencyStyle) {
        String label = getStringCellValue(sourceRow.getCell(0), evaluator).trim().toLowerCase();
        boolean headerRow = label.startsWith("empl");
        boolean totalRow = label.startsWith("total ");

        for (int c = 0; c <= 5; c++) {
            Cell cell = targetRow.getCell(c);
            if (cell == null) {
                cell = targetRow.createCell(c);
            }

            if (headerRow) {
                cell.setCellStyle(headerStyle);
                continue;
            }

            if (totalRow) {
                if (c == 3 || c == 5) {
                    cell.setCellStyle(footerCurrencyStyle);
                } else {
                    cell.setCellStyle(headerStyle);
                }
                continue;
            }

            if (c == 0) {
                cell.setCellStyle(centerStyle);
            } else if (c == 1 || c == 2) {
                cell.setCellStyle(leftStyle);
            } else {
                cell.setCellStyle(currencyStyle);
            }
        }
    }

    private boolean isGrandTotalRow(Row row, FormulaEvaluator evaluator) {
        String label = getStringCellValue(row.getCell(0), evaluator).trim().toLowerCase();
        return label.contains("grand total");
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
}
