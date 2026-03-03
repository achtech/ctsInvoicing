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

            String outputDirectory = outputGeneratedExcelsFilePath + "/Version_" + dateProvider.getCurrentDateTime();

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

            // 2. GENERATE CONSOLIDATED FILE (ONE SHEET, TABLES UNDER EACH OTHER)
            String consolidatedFileName = outputDirectory + "/Consolidated_Month_Forecast_" + currentMonthSpanish + ".xlsx";
            try (Workbook consolidatedWorkbook = new XSSFWorkbook()) {
                Sheet allTeamsSheet = consolidatedWorkbook.createSheet("All Teams Forecast");
                int currentRow = 0;
                List<Integer> teamTotalCostRows = new ArrayList<>();
                List<Integer> teamTotalHoursRows = new ArrayList<>();
                int costColumnIndex = -1;
                int hoursColumnIndex = -1;

                for (String serviceTeam : serviceTeamNames) {
                    for (int i = 0; i < monthsToProcess; i++) {
                        LocalDate dateForSheet = currentDate.plusMonths(i);
                        String monthNameEng = dateProvider.getMonthNameEnglish(dateForSheet);
                        String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);

                        currentRow = excelWriter.copyServiceHoursToSheet(
                                inputWorkbook,
                                allTeamsSheet,
                                currentRow,
                                serviceTeam,
                                SHEET_HORAS_SERVICIO + " " + monthNameSpa,
                                monthNameEng,
                                SHEET_AJUSTES,
                                SHEET_FACTURACIÓN + " " + monthNameSpa
                        );
                        
                        // After copying, the last row of the table is currentRow - 1
                        // We need to keep track of these rows for the final grand total
                        teamTotalCostRows.add(currentRow); 
                        
                        // Find cost and hours columns (they are at the end)
                        if (costColumnIndex == -1) {
                            Row lastRow = allTeamsSheet.getRow(currentRow - 1);
                            if (lastRow != null) {
                                costColumnIndex = lastRow.getLastCellNum() - 1;
                                hoursColumnIndex = costColumnIndex - 1;
                            }
                        }
                    }
                }
                
                // Add Grand Total at the bottom of the consolidated sheet
                if (currentRow > 0 && costColumnIndex != -1) {
                    currentRow += 2;
                    Row grandTotalRow = allTeamsSheet.createRow(currentRow++);
                    
                    CellStyle headerStyle = Helper.getHeaderStyle(consolidatedWorkbook);
                    CellStyle footerCurrencyStyle = Helper.getFooterCurrencyStyle(consolidatedWorkbook);
                    
                    Cell labelCell = grandTotalRow.createCell(hoursColumnIndex - 2);
                    labelCell.setCellValue("GRAND TOTAL (ALL PROJECTS)");
                    labelCell.setCellStyle(headerStyle);
                    allTeamsSheet.addMergedRegion(new CellRangeAddress(grandTotalRow.getRowNum(), grandTotalRow.getRowNum(), hoursColumnIndex - 2, hoursColumnIndex - 1));
                    
                    // Sum of all "Total" rows for Hours
                    Cell grandHoursCell = grandTotalRow.createCell(hoursColumnIndex);
                    StringJoiner hoursFormula = new StringJoiner("+");
                    for (Integer rowIdx : teamTotalCostRows) {
                        hoursFormula.add(Helper.getColumnLetter(hoursColumnIndex) + rowIdx);
                    }
                    grandHoursCell.setCellFormula(hoursFormula.toString());
                    grandHoursCell.setCellStyle(headerStyle);
                    
                    // Sum of all "Total" rows for Cost
                    Cell grandCostCell = grandTotalRow.createCell(costColumnIndex);
                    StringJoiner costFormula = new StringJoiner("+");
                    for (Integer rowIdx : teamTotalCostRows) {
                        costFormula.add(Helper.getColumnLetter(costColumnIndex) + rowIdx);
                    }
                    grandCostCell.setCellFormula(costFormula.toString());
                    grandCostCell.setCellStyle(footerCurrencyStyle);
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