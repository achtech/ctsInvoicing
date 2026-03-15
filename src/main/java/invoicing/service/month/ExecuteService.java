package invoicing.service.month;

import invoicing.Helper.Helper;
import invoicing.service.month.impl.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

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

            // Collect all projects/service teams across the requested months.
            // Some projects can start next month, so they may not appear in the current Facturación sheet.
            java.util.Set<String> serviceTeamNamesSet = new java.util.LinkedHashSet<>();
            for (int i = 0; i < monthsToProcess; i++) {
                LocalDate dateForSheet = currentDate.plusMonths(i);
                String monthNameSpa = dateProvider.getMonthNameSpanish(dateForSheet);
                String facturacionName = SHEET_FACTURACIÓN + " " + monthNameSpa;
                Sheet facturacionSheet = excelReader.getSheet(inputWorkbook, facturacionName);
                if (facturacionSheet == null) {
                    System.err.println("Warning: Sheet not found: " + facturacionName);
                    continue;
                }
                List<String> fullServiceTeamNames = serviceTeamExtractor.extractFullServiceTeamNames(facturacionSheet, inputWorkbook);
                serviceTeamNamesSet.addAll(serviceTeamExtractor.extractServiceTeamNames(fullServiceTeamNames));
            }

            List<String> serviceTeamNames = new java.util.ArrayList<>(serviceTeamNamesSet);

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

                    String outputFileName = fileNameGenerator.generateOutputFileName(currentMonth, currentYear, serviceTeam, outputGeneratedExcelsFilePath);
                    Helper.writeWorkbook(outputWorkbook, outputFileName);
                }
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
