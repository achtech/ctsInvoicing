package invoicing.service.month;

import invoicing.Helper.Helper;
import invoicing.service.month.impl.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.StringJoiner;

import static invoicing.service.month.ExcelFileNameGenerator.*;

public class ExecuteService {
    private final DateProvider dateProvider;
    private final ExcelFileNameGenerator fileNameGenerator;
    private final ExcelReader excelReader;
    private final ServiceTeamExtractor serviceTeamExtractor;
    private final ExcelWriter excelWriter;

    public ExecuteService(
            DateProvider dateProvider,
            ExcelFileNameGenerator fileNameGenerator,
            ExcelReader excelReader,
            ServiceTeamExtractor serviceTeamExtractor,
            ExcelWriter excelWriter) {
        this.dateProvider = dateProvider;
        this.fileNameGenerator = fileNameGenerator;
        this.excelReader = excelReader;
        this.serviceTeamExtractor = serviceTeamExtractor;
        this.excelWriter = excelWriter;
    }

    public void process(String inputExcelFilePath, String outputGeneratedExcelsFilePath, int monthsToProcess) {
        // Normalize the path for cross-platform compatibility
        File file = new File(inputExcelFilePath).getAbsoluteFile();
        System.out.println("Resolved path: " + file.getAbsolutePath());

        // Validate file
        if (!file.exists()) {
            System.err.println("Error: File does not exist: " + file.getAbsolutePath());
            return;
        }
        if (!file.isFile()) {
            System.err.println("Error: Path is not a file: " + file.getAbsolutePath());
            return;
        }
        if (!file.getName().toLowerCase().endsWith(".xlsx") && !file.getName().toLowerCase().endsWith(".xls")) {
            System.err.println("Error: File is not an Excel file (.xlsx or .xls): " + file.getAbsolutePath());
            return;
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook inputWorkbook = new XSSFWorkbook(fis)) {

            LocalDate currentDate = dateProvider.getCurrentDate(inputExcelFilePath);
            int currentYear = dateProvider.getYear(currentDate);
            int currentMonth = dateProvider.getMonthValue(currentDate);
            String currentMonthSpanish = dateProvider.getMonthNameSpanish(currentDate);

            String outputDirectory = outputGeneratedExcelsFilePath;

            Sheet facturacionSheet = excelReader.getSheet(inputWorkbook, "Facturación " + currentMonthSpanish);

            List<String> fullServiceTeamNames = serviceTeamExtractor.extractFullServiceTeamNames(facturacionSheet, inputWorkbook);
            List<String> serviceTeamNames = serviceTeamExtractor.extractServiceTeamNames(fullServiceTeamNames);

            // 1. GENERATE SEPARATE FILES
            for (String serviceTeam : serviceTeamNames) {
                // Collect month names for this run
                List<String> monthNames = new ArrayList<>();
                for (int i = 0; i < monthsToProcess; i++) {
                    monthNames.add(dateProvider.getMonthNameEnglish(currentDate.plusMonths(i)));
                }

                try (Workbook outputWorkbook = excelWriter.createWorkbookWithSheets(monthNames)) {
                    for (int i = 0; i < monthsToProcess; i++) {
                        LocalDate dateForSheet = currentDate.plusMonths(i);
                        String monthNameEng = dateProvider.getMonthNameEnglish(dateForSheet);
                        String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);

                        excelWriter.copyServiceHoursSheetData(
                                inputWorkbook,
                                outputWorkbook,
                                serviceTeam,
                                SHEET_HORAS_SERVICIO + " " + monthNameSpa,
                                SHEET_SERVICE_HOURS_DETAILS + " " + monthNameEng,
                                SHEET_AJUSTES,
                                SHEET_FACTURACIÓN + " " + monthNameSpa
                        );
                    }

                    String outputFileName = fileNameGenerator.generateOutputFileName(currentMonth, currentYear, serviceTeam, outputDirectory);
                    Helper.writeWorkbook(outputWorkbook, outputFileName);
                }
            }

            // 2. GENERATE CONSOLIDATED FILE (one sheet, tables stacked under each other)
            String inputFileName = new File(inputExcelFilePath).getName().replace(".xlsx", "");
            String consolidatedFileName = outputDirectory
                    + "/Consolidated_Month_Forecast_" + currentMonthSpanish + "_" + inputFileName + ".xlsx";

            try (Workbook consolidatedWorkbook = new XSSFWorkbook()) {

                Sheet allTeamsSheet = consolidatedWorkbook.createSheet("All Teams Forecast");

                // Track grand-total row positions across all teams (1-based Excel row for each "Total" row)
                List<Integer> grandTotalHoursRows = new ArrayList<>();
                List<Integer> grandTotalCostRows  = new ArrayList<>();

                int currentRow = 0;

                for (String serviceTeam : serviceTeamNames) {
                    for (int i = 0; i < monthsToProcess; i++) {
                        LocalDate dateForSheet = currentDate.plusMonths(i);
                        String monthNameEng = dateProvider.getMonthNameEnglish(dateForSheet);
                        String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);

                        int nextFreeRow = excelWriter.copyServiceHoursToConsolidatedSheet(
                                inputWorkbook,
                                allTeamsSheet,
                                currentRow,
                                serviceTeam,
                                SHEET_HORAS_SERVICIO + " " + monthNameSpa,
                                SHEET_SERVICE_HOURS_DETAILS + " " + monthNameEng,
                                SHEET_AJUSTES,
                                SHEET_FACTURACIÓN + " " + monthNameSpa
                                
                        );

                        // The Total row is 2 blank rows before nextFreeRow (0-based),
                        // converted to 1-based Excel row: (nextFreeRow - 2) + 1 = nextFreeRow - 1
                        int totalRowExcel1Based = nextFreeRow - 1;
                        grandTotalHoursRows.add(totalRowExcel1Based);
                        grandTotalCostRows.add(totalRowExcel1Based);

                        currentRow = nextFreeRow;
                    }
                }

                // ── Grand Total row ───────────────────────────────────────────
                if (currentRow > 0) {
                    currentRow += 1; // one extra blank before grand total

                    Row grandTotalRow = allTeamsSheet.createRow(currentRow);

                    CellStyle headerStyle         = Helper.getHeaderStyle(consolidatedWorkbook);
                    CellStyle footerCurrencyStyle  = Helper.getFooterCurrencyStyle(consolidatedWorkbook);

                    // Label spanning cols 0-2
                    Cell labelCell = grandTotalRow.createCell(0);
                    labelCell.setCellValue("GRAND TOTAL (ALL PROJECTS)");
                    labelCell.setCellStyle(headerStyle);
                    allTeamsSheet.addMergedRegion(new CellRangeAddress(
                            currentRow, currentRow, 0, 2));

                    // Sum of all team "Total" hours (column E = index 4)
                    StringJoiner hoursFormula = new StringJoiner("+");
                    for (Integer rowIdx : grandTotalHoursRows) {
                        hoursFormula.add("E" + (rowIdx - 1));
                    }
                    String grandRowNumber = String.valueOf(currentRow + 1);
                    Cell grandRateCell = grandTotalRow.createCell(3);
                    grandRateCell.setCellFormula("IF(E" + grandRowNumber + "=0,\"\",F" + grandRowNumber + "/E" + grandRowNumber + ")");
                    grandRateCell.setCellStyle(footerCurrencyStyle);
                    Cell grandHoursCell = grandTotalRow.createCell(4);
                    grandHoursCell.setCellFormula(hoursFormula.toString());
                    grandHoursCell.setCellStyle(headerStyle);

                    // Sum of all team "Total" cost (column F = index 5)
                    StringJoiner costFormula = new StringJoiner("+");
                    for (Integer rowIdx : grandTotalCostRows) {
                        costFormula.add("F" + (rowIdx - 1));
                    }
                    Cell grandCostCell = grandTotalRow.createCell(5);
                    grandCostCell.setCellFormula(costFormula.toString());
                    grandCostCell.setCellStyle(footerCurrencyStyle);
                }

                // Auto-size the consolidated columns
                for (int col = 0; col < 6; col++) {
                    allTeamsSheet.autoSizeColumn(col);
                }

                Helper.writeWorkbook(consolidatedWorkbook, consolidatedFileName);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void executeScript(String inputExcelFilePath, String outputExcelsFilePath, int monthsToProcess) throws Exception {
        ExecuteService processor = new ExecuteService(
                new DefaultDateProvider(),
                new DefaultExcelFileNameGenerator(),
                new DefaultExcelReader(),
                new ServiceTeamExtractorImpl(),
                new DefaultExcelWriter()
        );
        processor.process(inputExcelFilePath, outputExcelsFilePath, monthsToProcess);
    }
}
